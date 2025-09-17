#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de mallas HTML interactivas desde Excel.

Uso:
    python3 generar_mallas.py archivo.xlsx
    python3 generar_mallas.py archivo.xlsx --outdir dist
    python3 generar_mallas.py archivo.xlsx --randomize-colors --seed 123
    python3 generar_mallas.py --selftest

Requisitos:
    pip install pandas openpyxl
"""

import argparse, os, re, sys, json, html, random, math
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("Falta pandas. Instala con: pip install pandas openpyxl", file=sys.stderr)
    sys.exit(1)

ROMANS = {"I":1,"II":2,"III":3,"IV":4,"V":5,"VI":6,"VII":7,"VIII":8,"IX":9,"X":10}

# Bolsa base de 15 colores
COLOR_BAG = [
    "#2563eb", "#a855f7", "#06b6d4", "#ef4444", "#10b981",
    "#f97316", "#84cc16", "#eab308", "#14b8a6", "#3b82f6",
    "#ec4899", "#22d3ee", "#f59e0b", "#34d399", "#9333ea",
]

# ---------- Utilidades de parseo ----------
def parse_level(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    try:
        return int(float(s))
    except:
        pass
    up = s.upper()
    return ROMANS.get(up, None)

def parse_int(x, default=0):
    if pd.isna(x) or str(x).strip()=="":
        return default
    try:
        return int(float(str(x).strip()))
    except:
        return default

def norm_cols(df):
    mapping = {c: re.sub(r"\s+","", c.strip().upper()) for c in df.columns}
    return df.rename(columns=mapping)

def infer_program_from_sheet(sheet_name):
    m = re.match(r"^(.*)\s*-\s*([A-Za-z0-9_]+)\s*$", sheet_name.strip())
    if m:
        title = m.group(1).strip()
        code = m.group(2).strip()
        return title, code
    parts = re.split(r"\s+", sheet_name.strip())
    code = None
    if parts:
        last = parts[-1]
        if re.fullmatch(r"[A-Z0-9_]{3,}", last):
            code = last
    return sheet_name.strip(), (code or sheet_name.strip().upper()[:4])

def infer_program_from_filename(filename):
    base = os.path.basename(filename).replace(".xlsx","")
    m = re.match(r"^(.*)\s*-\s*([A-Za-z0-9_]+)\s*$", base.strip())
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return base.strip(), base.strip().upper()[:4]

# ---------- Generación de colores ----------
def hsl_to_hex(h, s, l):
    """Convierte HSL (0-360, 0-1, 0-1) a #RRGGBB"""
    h = (h % 360) / 360.0
    def f(n):
        k = (n + 12*h) % 12
        a = s * min(l, 1 - l)
        c = l - a * max(-1, min(k-3, 9-k, 1))
        return int(round(255 * c))
    return "#{:02x}{:02x}{:02x}".format(f(0), f(8), f(4))

def expand_color_bag(min_size, randomize=False, seed=None):
    """
    Devuelve una lista de colores HEX de tamaño >= min_size.
    - Comienza con COLOR_BAG.
    - Si se requieren más, genera colores adicionales uniformemente separados (ángulo áureo),
      variando H y alternando L para buena separación. En modo aleatorio usa semilla.
    """
    bag = COLOR_BAG[:]
    if len(bag) >= min_size:
        return bag

    # Parámetros de generación
    phi = 137.508  # golden angle en grados
    base_h = 197.0
    base_s = 0.70
    l_steps = [0.55, 0.50, 0.60, 0.45]  # variaciones de lightness para separarlos aún más

    rng = random.Random(seed) if (randomize and seed is not None) else None
    start_h = (rng.uniform(0, 360) if rng else base_h)

    needed = min_size - len(bag)
    for i in range(needed):
        # hue espaciado por ángulo áureo
        h = (start_h + phi * i) % 360
        # alterna levels de lightness de forma determinista/reproducible
        l = l_steps[i % len(l_steps)]
        s = base_s
        col = hsl_to_hex(h, s, l)
        bag.append(col)
    return bag

def assign_colors_to_areas(areas, randomize=False, seed=None):
    """
    Devuelve lista de tuplas (var_css, color) para cada área.
    - randomize=False: asignación estable (determinística) por hash + bolsa expandida si hace falta.
    - randomize=True: baraja (con semilla si la hay) la bolsa expandida y asigna secuencialmente.
    """
    # normalizar áreas
    areas = [a for a in sorted(set(a or "OTRO" for a in areas))]
    bag = expand_color_bag(len(areas), randomize=randomize, seed=seed)

    if randomize:
        rng = random.Random(seed) if seed is not None else random.Random()
        rng.shuffle(bag)

    assigned = {}
    used = set()

    if not randomize:
        # determinístico por hash + manejo de colisiones, ahora con bolsa siempre suficiente
        for a in areas:
            idx = abs(hash(a.upper())) % len(bag)
            for step in range(len(bag)):
                c = bag[(idx + step) % len(bag)]
                if c not in used:
                    assigned[a] = c
                    used.add(c)
                    break
    else:
        # secuencial sobre bolsa barajada (bag >= áreas)
        i = 0
        for a in areas:
            assigned[a] = bag[i]
            i += 1

    # Variables CSS
    pairs = []
    for a in areas:
        var_name = f"--AREA-{re.sub(r'[^A-Z0-9_-]', '_', a.upper())}"
        pairs.append((var_name, assigned[a]))
    return pairs

# ---------- Construcción de cursos ----------
def build_courses(df):
    df = norm_cols(df).fillna("")
    req = ["LEVEL","ID","NAME","CREDITS","AREA"]
    missing = [c for c in req if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}. Necesitas {req} (más PRE1, PRE2,... opcional).")
    pre_cols = [c for c in df.columns if c.startswith("PRE")]
    courses = []
    for _, r in df.iterrows():
        level = parse_level(r["LEVEL"])
        cid = str(r["ID"]).strip()
        name = str(r["NAME"]).strip()
        area = str(r["AREA"]).strip() or "OTRO"
        credits = parse_int(r["CREDITS"], 0)
        if not level or not cid or not name:
            continue
        prereq = []
        for pc in pre_cols:
            val = str(r[pc]).strip()
            if val:
                prereq.append(val)
        courses.append({
            "id": cid,
            "name": name,
            "area": area,
            "level": level,
            "prereq": prereq,
            "credits": credits
        })
    if not courses:
        raise ValueError("No se construyeron cursos. Verifica que las filas tengan LEVEL/ID/NAME válidos.")
    return courses

# ---------- Plantilla HTML ----------
def build_html(program_title, program_code, courses, area_vars):
    # Variables CSS por área
    css_vars = "".join([f"{var}:{col};" for var, col in area_vars])
    courses_json = json.dumps(courses, ensure_ascii=False)
    code_json = json.dumps(program_code, ensure_ascii=False)

    template = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Planificador Interactivo – __TITLE__</title>
<style>
  :root{
    --bg:#0b1220; --paper:#0f172a; --ink:#e5e7eb; --muted:#94a3b8;
    --ok:#22c55e; --warn:#f59e0b; --chip:#1f2937;
    --OTRO:#9ca3af;
    __AREA_VARS__
    --levels: 10;      /* nº de columnas (JS lo actualiza) */
    --col-min: 180px;  /* ancho fijo de cada columna */
  }
  body{margin:0; font:15px/1.4 system-ui,Segoe UI,Roboto,Helvetica,Arial,sans-serif; color:var(--ink); background:var(--bg)}
  header{padding:12px; background:#111827; position:sticky; top:0; z-index:10}
  .controls{margin-top:8px; display:flex; gap:8px; flex-wrap:wrap; align-items:center}
  .controls input{background:#0b1322; color:var(--ink); border:1px solid #203249; border-radius:8px; padding:6px 10px}
  .legend{margin-top:8px; display:flex; gap:8px; flex-wrap:wrap; align-items:center}
  .pill{display:inline-flex; align-items:center; gap:6px; padding:4px 10px; border-radius:999px; background:#1f2937; border:1px solid #2a3443; font-size:12px; color:#cbd5e1}
  .dot{width:10px; height:10px; border-radius:50%; background:#9ca3af}
  button{background:#1f2937; color:var(--ink); border:1px solid #374151; padding:6px 10px; border-radius:8px; cursor:pointer}

  /* Contenedor con scroll horizontal */
  .grid-scroll{
    overflow-x:auto;
    overscroll-behavior-x: contain;
    -webkit-overflow-scrolling: touch;
    padding: 0 8px;
  }

  /* Grid con columnas de ancho fijo */
  .grid{
    display:grid;
    gap:14px;
    grid-template-columns: repeat(var(--levels, 10), var(--col-min));
    align-items:start;
    padding:16px;
    width:max-content;  /* usa ancho intrínseco → aparece scroll si no caben */
    min-width:100%;
  }

  .col{display:flex; flex-direction:column; gap:14px}
  .col-head{font-weight:700; color:#cbd5e1; text-align:center}
  .course{position:relative; background:#0f172a; border:1px solid #1f2a40; border-radius:12px; padding:18px 12px 8px 32px; cursor:pointer; aspect-ratio:2/1; opacity:.4; pointer-events:none}
  .course.available{opacity:1; pointer-events:auto}
  .course.done{border-color:var(--ok)}
  .course.unlock{outline:2px dashed var(--warn)}
  .course.locked{opacity:.45; filter:grayscale(.1); cursor:not-allowed}
  .tick{position:absolute; left:6px; top:6px; width:16px; height:16px; border-radius:4px; border:1px solid #273244; display:grid; place-items:center}
  .tick svg{stroke:#93c5fd; opacity:0; fill:none; stroke-width:2.5; stroke-linecap:round; stroke-linejoin:round}
  .course.done .tick svg{opacity:1}
  .area{position:absolute; right:6px; top:6px; font-size:11px; padding:2px 6px; border-radius:999px; font-weight:600; background:var(--OTRO); color:#0b1220}
  .course::before{content:""; position:absolute; left:0; top:0; bottom:0; width:6px; border-radius:12px 0 0 12px; background:var(--leftbar,var(--OTRO))}
  #panel{background:#111827; padding:12px 16px; margin:0 16px 16px; border-radius:8px;}
  #panel span{display:inline-block; margin-right:12px;}
  .credits{margin:8px 16px 8px; padding:14px; border:1px solid #1f2a40; border-radius:12px; background:#0f172a}
  .credits h2{margin:0 0 10px 0; font-size:16px}
  .credits .row{display:flex; flex-wrap:wrap; gap:12px}
  .stat{background:#0b1322; border:1px solid #203249; border-radius:10px; padding:10px 12px; min-width:160px}
  .stat b{display:block; font-size:12px; color:#cbd5e1; margin-bottom:4px}
  .stat .val{font-size:18px; font-weight:800}
  .tests{margin:8px 16px 16px; font-size:12px}
  .ok{color:#22c55e} .fail{color:#ef4444}

  /* Ajustes responsivos */
  @media (max-width: 1280px){
    :root{ --col-min: 170px; }
    .grid { column-gap: 12px; row-gap: 12px; }
    .course { padding: 14px 10px 8px 28px; }
  }
  @media (max-width: 900px){
    :root{ --col-min: 160px; }
    .course { aspect-ratio: 7/5; }
    .col-head { font-size: 14px; }
    .course .name { font-size: 14px; }
    .course .meta { font-size: 12px; }
  }
  @media (max-width: 600px){
    :root{ --col-min: 150px; }
    .course { aspect-ratio: 1.4/1; padding: 12px 8px 6px 24px; }
    .controls { gap: 6px; }
    .controls input, button { padding: 6px 8px; }
  }

  /* Tema claro */
  body[data-theme="light"]{color:#0b1220; background:#f8fafc}
  body[data-theme="light"] header{background:#e2e8f0}
  body[data-theme="light"] .course{background:#ffffff; border-color:#d1d5db}
  body[data-theme="light"] .credits, body[data-theme="light"] #panel{background:#ffffff; border-color:#e2e8f0}
  body[data-theme="light"] .pill{background:#f1f5f9; border-color:#cbd5e1; color:#334155}
  body[data-theme="light"] .controls input{background:#ffffff; border-color:#cbd5e1; color:#0b1220}
  body[data-theme="light"] button{background:#ffffff; border-color:#cbd5e1; color:#0b1220}
</style>
</head>
<body>
<header>
  <h1>Planificador Interactivo – __TITLE__</h1>
  <div class="controls">
    <input id="inpStudentCode" placeholder="Código estudiante" />
    <input id="inpStudentName" placeholder="Nombre" />
    <button id="btnReset">Reiniciar</button>
    <button id="btnExportCSV">Exportar CSV</button>
    <button id="btnTheme">Tema: Oscuro</button>
  </div>
  <div id="legend" class="legend"></div>
</header>
<div id="panel"></div>
<div class="grid-scroll">
  <div id="grid" class="grid"></div>
</div>
<section class="credits" id="creditsPanel">
  <h2>Progreso</h2>
  <div class="row">
    <div class="stat"><b>Créditos aprobados</b><div class="val" id="crEarned">0</div></div>
    <div class="stat"><b>Créditos totales</b><div class="val" id="crDefined">0</div></div>
    <div class="stat"><b>Créditos faltantes</b><div class="val" id="crRemain">0</div></div>
    <div class="stat"><b>Cursos aprobados</b><div class="val" id="countDone">0</div></div>
    <div class="stat"><b>Cursos pendientes</b><div class="val" id="countTodo">0</div></div>
  </div>
</section>
<div id="tests" class="tests"></div>

<script>
"use strict";
const PROGRAM_CODE = __CODE_JSON__;
const COURSES = __COURSES_JSON__;

function $$(s){return Array.from(document.querySelectorAll(s));}
function $(s){return document.querySelector(s);}
function save(progress){ localStorage.setItem('progress:'+PROGRAM_CODE, JSON.stringify([...progress])); }
function load(){ return new Set(JSON.parse(localStorage.getItem('progress:'+PROGRAM_CODE)||'[]')); }
function colorForArea(a){
  const v='--AREA-'+String(a||'OTRO').toUpperCase();
  const val=(getComputedStyle(document.documentElement).getPropertyValue(v)||'').trim();
  return val || getComputedStyle(document.documentElement).getPropertyValue('--OTRO').trim();
}
function isUnlocked(c, progress){
  return !progress.has(c.id) && (c.prereq||[]).every(p=>progress.has(p));
}

function render(){
  const grid=$('#grid'); grid.innerHTML='';
  const maxLevel=Math.max(...COURSES.map(c=>c.level||1));
  // fija el número de columnas (niveles)
  document.documentElement.style.setProperty('--levels', String(maxLevel));

  for(let lvl=1; lvl<=maxLevel; lvl++) {
    const col=document.createElement('div'); col.className='col';
    const head=document.createElement('div'); head.className='col-head'; head.textContent='NIVEL '+lvl;
    col.appendChild(head);
    COURSES.filter(c=>c.level===lvl).forEach(c=>{
      const el=document.createElement('div'); el.className='course'; el.dataset.id=c.id;
      el.style.setProperty('--leftbar',colorForArea(c.area));
      el.innerHTML = `
        <div class="tick"><svg viewBox="0 0 24 24" width="14" height="14"><polyline points="20 6 9 17 4 12"/></svg></div>
        <div class="area" style="background:${colorForArea(c.area)}">${c.area}</div>
        <div class="code">${c.id}</div>
        <div class="name">${c.name}</div>
        <div class="meta">${(c.prereq&&c.prereq.length?('Requiere: '+c.prereq.join(', ')):'Sin prereq')} | Créditos: ${c.credits}</div>`;
      el.addEventListener('click', ()=>{
        const progress = load();
        const canToggle = progress.has(c.id) || isUnlocked(c, progress);
        if(!canToggle) return;
        if(progress.has(c.id)) progress.delete(c.id); else progress.add(c.id);
        save(progress); update(); computeCredits();
      });
      col.appendChild(el);
    });
    grid.appendChild(col);
  }
  renderLegend();
  update();
  computeCredits();
}

function update(){
  const progress = load();
  $$('.course').forEach(el=>{
    el.classList.remove('done','unlock','available','locked');
    const c = COURSES.find(x=>x.id===el.dataset.id); if(!c) return;
    if(progress.has(c.id)) {
      el.classList.add('done','available');
      const svg=el.querySelector('svg'); if(svg) svg.style.opacity=1;
    } else if(isUnlocked(c, progress)) {
      el.classList.add('unlock','available');
    } else {
      el.classList.add('locked');
    }
  });
  computeStats();
}

function computeCredits(){
  const progress = load();
  const total = COURSES.reduce((a,c)=>a+(c.credits||0),0);
  const earned = [...progress].map(id=>COURSES.find(c=>c.id===id)).filter(Boolean).reduce((a,c)=>a+(c.credits||0),0);
  const pending = total-earned;
  const approved = progress.size;
  const missing = COURSES.length - approved;
  document.getElementById('panel').innerHTML =
    `<span><b>Créditos aprobados:</b> ${earned}</span>`+
    `<span><b>Créditos faltantes:</b> ${pending}</span>`+
    `<span><b>Cursos aprobados:</b> ${approved}</span>`+
    `<span><b>Cursos faltantes:</b> ${missing}</span>`;
}

function computeStats(){
  const progress = load();
  let earned=0,total=0,done=0;
  for(const c of COURSES){ total+=(c.credits||0); if(progress.has(c.id)){ done++; earned+=(c.credits||0); } }
  const remain=Math.max(0,total-earned);
  document.getElementById('crEarned').textContent=earned;
  document.getElementById('crDefined').textContent=total;
  document.getElementById('crRemain').textContent=remain;
  document.getElementById('countDone').textContent = `${done} / ${COURSES.length}`;
  document.getElementById('countTodo').textContent = `${COURSES.length - done}`;
}

function renderLegend(){
  const L=document.getElementById('legend');
  const areas=[...new Set(COURSES.map(c=>c.area))].sort();
  L.innerHTML = areas.map(a=>`<span class="pill"><span class="dot" style="background:${colorForArea(a)}"></span>${a}</span>`).join('');
}

function csvEscape(val){
  const s = String(val==null? '': val);
  return /[",\\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s;
}
function buildCSV(progress){
  const headers=['ID','Nombre','Área','Nivel','Créditos','Prerrequisitos','Estado'];
  const rows = COURSES.map(c=>{
    const estado = progress.has(c.id)? 'Aprobado' : (isUnlocked(c,progress)? 'Disponible' : 'Bloqueado');
    return [c.id, c.name, c.area, c.level, c.credits, (c.prereq||[]).join(' | '), estado];
  });
  const totalCred=COURSES.reduce((s,c)=>s+(c.credits||0),0);
  const earned=[...progress].map(id=>COURSES.find(c=>c.id===id)).filter(Boolean).reduce((s,c)=>s+(c.credits||0),0);
  const remain=Math.max(0,totalCred-earned);
  const done=[...progress].length; const todo=COURSES.length - done;
  rows.push([]);
  rows.push(['RESUMEN','','','','','','']);
  rows.push(['Créditos aprobados', earned]);
  rows.push(['Créditos totales', totalCred]);
  rows.push(['Créditos faltantes', remain]);
  rows.push(['Cursos aprobados', done]);
  rows.push(['Cursos pendientes', todo]);
  return [headers, ...rows];
}
function exportCSV(){
  const progress = load();
  const rows = buildCSV(progress);
  // Metadatos de estudiante/programa
  const student = loadStudent();
  const meta = [
    ['Programa', PROGRAM_CODE],
    ['Código estudiante', student.code||''],
    ['Nombre', student.name||''],
    ['Fecha export', new Date().toISOString()],
    []
  ];
  const all = [...meta, ...rows];
  const csv = all.map(r=>r.map(csvEscape).join(',')).join('\\n');
  const blob = new Blob(['\\ufeff'+csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download=PROGRAM_CODE+'_progreso.csv'; a.click();
  URL.revokeObjectURL(url);
}

// ===== Student meta + Theme =====
function studentKey(){ return 'student:'+PROGRAM_CODE; }
function loadStudent(){ try{ return JSON.parse(localStorage.getItem(studentKey())||'{}'); }catch(e){ return {}; } }
function saveStudent(obj){ localStorage.setItem(studentKey(), JSON.stringify(obj||{})); }
function bindStudentInputs(){
  const codeEl=document.getElementById('inpStudentCode');
  const nameEl=document.getElementById('inpStudentName');
  const data=loadStudent();
  if(codeEl) codeEl.value=data.code||'';
  if(nameEl) nameEl.value=data.name||'';
  codeEl&&codeEl.addEventListener('input',()=>{ const d=loadStudent(); d.code=codeEl.value; saveStudent(d); });
  nameEl&&nameEl.addEventListener('input',()=>{ const d=loadStudent(); d.name=nameEl.value; saveStudent(d); });
}

function themeKey(){ return 'theme:'+PROGRAM_CODE; }
function applyTheme(){
  const t = localStorage.getItem(themeKey())||'dark';
  document.body.setAttribute('data-theme', t==='light'?'light':'dark');
  const btn=document.getElementById('btnTheme'); if(btn) btn.textContent = 'Tema: ' + (t==='light'?'Claro':'Oscuro');
}
function toggleTheme(){
  const curr = localStorage.getItem(themeKey())||'dark';
  const next = curr==='light'?'dark':'light';
  localStorage.setItem(themeKey(), next);
  applyTheme();
}

document.getElementById('btnReset').onclick = ()=>{ save(new Set()); render(); };
document.getElementById('btnExportCSV').onclick = exportCSV;
document.getElementById('btnTheme').onclick = ()=>{ toggleTheme(); };
bindStudentInputs();
render();
applyTheme();

// Tests básicos (consola)
(function runTests(){
  const out=[]; const ok=n=>out.push('OK '+n); const fail=(n,m)=>out.push('FAIL '+n+': '+m);
  let progress = new Set();
  try {
    const unlocked = COURSES.filter(c=>isUnlocked(c,progress));
    const should   = COURSES.filter(c=>!c.prereq||c.prereq.length===0);
    if(unlocked.length===should.length && unlocked.every(u=>should.some(s=>s.id===u.id))) ok('Inicial unlocked = sin PRE');
    else fail('Inicial','count mismatch');
  } finally {
    console.log('[TESTS]', out.join(' | '));
  }
})();
</script>
</body>
</html>
"""

    html_out = (template
                .replace("__TITLE__", html.escape(program_title))
                .replace("__AREA_VARS__", css_vars)
                .replace("__COURSES_JSON__", courses_json)
                .replace("__CODE_JSON__", code_json))
    return html_out



# ---------- Escritura & proceso ----------
def write_program_html(outdir, title, code, courses, randomize_colors=False, seed=None):
    area_vars = assign_colors_to_areas([c["area"] for c in courses], randomize=randomize_colors, seed=seed)
    code_sanitized = re.sub(r"[^A-Za-z0-9_-]+","_", code.strip() or "PROG")
    filename = f"{code_sanitized}.html"
    path = os.path.join(outdir, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(build_html(title, code, courses, area_vars))
    return path

def process_excel(path, outdir, randomize_colors=False, seed=None, selftest=False):
    if selftest:
        print("[SELFTEST] parse_level I ->", parse_level("I"))
        print("[SELFTEST] parse_level 3 ->", parse_level("3"))
        t, c = infer_program_from_filename("CIENCIA DE DATOS - CDAT.xlsx")
        print("[SELFTEST] infer filename ->", t, c)
        # reproducibilidad de colores en aleatorio
        demo = ["CDAT","CBAS","ISCO","ECON","CHUL","CHUM","IIND","AEMP","ECOU","OTRO","X1","X2","X3","X4","X5","X6","X7"]
        a1 = assign_colors_to_areas(demo, randomize=True, seed=123)
        a2 = assign_colors_to_areas(demo, randomize=True, seed=123)
        assert a1 == a2, "Asignación aleatoria con misma seed debe ser idéntica"
        print("[SELFTEST] OK reproducibilidad y expansión de colores")

    if not os.path.exists(path):
        raise FileNotFoundError(path)

    xl = pd.ExcelFile(path)
    outputs = []

    def has_program_cols(df):
        cols = set(norm_cols(df).columns)
        return "PROGRAM" in cols or "PROGRAM_CODE" in cols

    multi_handled = False
    for sheet in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str)
        if has_program_cols(df):
            df = norm_cols(df).fillna("")
            group_key = "PROGRAM_CODE" if "PROGRAM_CODE" in df.columns else "PROGRAM"
            for g, gdf in df.groupby(group_key):
                code_guess = str(g).strip() or infer_program_from_filename(path)[1]
                title_guess = (df["PROGRAM"].iloc[0].strip()
                    if "PROGRAM" in df.columns and str(df["PROGRAM"].iloc[0]).strip()
                    else infer_program_from_filename(path)[0])
                courses = build_courses(gdf)
                if "PROGRAM" in gdf.columns and str(gdf["PROGRAM"].iloc[0]).strip():
                    title_guess = str(gdf["PROGRAM"].iloc[0]).strip()
                if "PROGRAM_CODE" in gdf.columns and str(gdf["PROGRAM_CODE"].iloc[0]).strip():
                    code_guess = str(gdf["PROGRAM_CODE"].iloc[0]).strip()
                outp = write_program_html(outdir, title_guess, code_guess, courses,
                                          randomize_colors=randomize_colors, seed=seed)
                outputs.append(outp)
            multi_handled = True

    if not multi_handled:
        for sheet in xl.sheet_names:
            df = pd.read_excel(path, sheet_name=sheet, dtype=str)
            courses = build_courses(df)
            title_guess, code_guess = infer_program_from_sheet(sheet)
            if not code_guess or code_guess.upper() == sheet.upper()[:4]:
                ftitle, fcode = infer_program_from_filename(path)
                if ftitle and fcode:
                    if title_guess == sheet:
                        title_guess = ftitle
                    if (not code_guess) or code_guess == sheet.upper()[:4]:
                        code_guess = fcode
            outp = write_program_html(outdir, title_guess, code_guess, courses,
                                      randomize_colors=randomize_colors, seed=seed)
            outputs.append(outp)

    return outputs

def main():
    ap = argparse.ArgumentParser(description="Genera mallas HTML interactivas por programa desde Excel.")
    ap.add_argument("excel", nargs="?", help="Ruta al archivo .xlsx")
    ap.add_argument("--outdir", default="dist", help="Directorio de salida (por defecto: dist)")
    ap.add_argument("--randomize-colors", action="store_true",
                    help="Asigna colores aleatorios por área (usa --seed para reproducibilidad)")
    ap.add_argument("--seed", type=int, default=None,
                    help="Semilla para la aleatoriedad (sólo aplica si usas --randomize-colors)")
    ap.add_argument("--selftest", action="store_true", help="Ejecuta pruebas internas de parseo y salida")
    args = ap.parse_args()

    if args.selftest and not args.excel:
        demo = ["CDAT","CBAS","ISCO","ECON","CHUL","CHUM","IIND","AEMP","ECOU","OTRO","X1","X2","X3","X4","X5","X6","X7"]
        a1 = assign_colors_to_areas(demo, randomize=True, seed=42)
        a2 = assign_colors_to_areas(demo, randomize=True, seed=42)
        assert a1 == a2, "Con la misma seed la asignación debe coincidir"
        print("[SELFTEST] OK seed reproducible")
        sys.exit(0)

    if not args.excel:
        ap.error("Debes indicar el archivo Excel. Ej: python3 generar_mallas.py archivo.xlsx")

    os.makedirs(args.outdir, exist_ok=True)
    outs = process_excel(args.excel, args.outdir,
                         randomize_colors=args.randomize_colors,
                         seed=args.seed,
                         selftest=args.selftest)
    print("Archivos generados:")
    for p in outs:
        print(" -", p)

if __name__ == "__main__":
    main()
