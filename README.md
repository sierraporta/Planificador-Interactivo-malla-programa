# Planificador Interactivo de Mallas Curriculares
Herramienta para generar **mallas curriculares interactivas** (un único archivo HTML, sin servidor) a partir de un **Excel**.  
Permite marcar cursos aprobados, respeta **prerrequisitos**, calcula **créditos aprobados/faltantes**, exporta el **progreso a CSV** e incluye campos para **código de estudiante** y **nombre**, además de **tema claro/oscuro** con persistencia por programa.

[![Python](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![pandas](https://img.shields.io/badge/pandas-%3E=1.0-yellowgreen.svg)](https://pandas.pydata.org/)
[![openpyxl](https://img.shields.io/badge/openpyxl-required-orange.svg)](https://openpyxl.readthedocs.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![pandas](https://img.shields.io/badge/pandas-%3E=1.0-yellowgreen.svg)](https://pandas.pydata.org/)
[![openpyxl](https://img.shields.io/badge/openpyxl-required-orange.svg)](https://openpyxl.readthedocs.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![PyPI Downloads](https://img.shields.io/pypi/dm/tu-paquete)](https://pypi.org/project/tu-paquete/)

---

## Características
- ✅ HTML **autónomo** (funciona offline).
- ✅ **Desbloqueo** automático de cursos por prerrequisitos.
- ✅ **Progreso** persistente por programa (usa `localStorage`).
- ✅ **Exportación a CSV** con detalle + resumen + datos del estudiante.
- ✅ **Tema** claro/oscuro, también persistente.
- ✅ Colores por **área** con paleta automática (expansión si hay > 15 áreas y opción de aleatoriedad reproducible).
- ✅ Soporta **una o varias** mallas (hojas) en el mismo Excel.

---

## Estructura del repositorio
.
├── generar_mallas.py # Script: Excel → HTML interactivo
├── ejemplos/ # (opcional) Excels de ejemplo
├── dist/ # Salida (HTMLs generados)
└── README.md


---

## Requisitos
- Python 3.8+
- Paquetes: `pandas`, `openpyxl`
Instalación:
```bash
pip install pandas openpyxl
```

# Formato del Excel

### 1) Nombre de la hoja

Usa el patrón:

TITULO DEL PROGRAMA - CODIGO


- **TITULO DEL PROGRAMA** → se muestra como “Planificador Interactivo – TITULO…” en el HTML.  
- **CODIGO** → nombre del archivo generado: `CODIGO.html`.  

> Si la hoja no tiene guion, el script usa el nombre del archivo Excel como respaldo para inferir título/código.

---

### 2) Columnas requeridas

| Columna  | Tipo   | Descripción                                                    |
|----------|--------|----------------------------------------------------------------|
| `LEVEL`  | int/str| Nivel del curso (1,2,3… o I,II,III…).                           |
| `ID`     | str    | Identificador completo `AREA-COD` (ej. `CBAS-M01A`).            |
| `NAME`   | str    | Nombre del curso.                                               |
| `CREDITS`| int    | Créditos del curso.                                             |
| `AREA`   | str    | Área del curso (ej. `CBAS`, `ISCO`, `CDAT`…).                   |

---

### 3) Prerrequisitos (opcionales)

- `PRE1`, `PRE2`, `PRE3`, …  
- Recomendado: usar **ID completo** (mismo formato que `ID`, ej. `CBAS-M01A`).  
- Se admiten múltiples separados por `,`, `;` o `/`.

**Ejemplo mínimo:**

LEVEL,AREA,ID,NAME,CREDITS,PRE1,PRE2
1,CBAS,CBAS-M01A,Cálculo Diferencial,4,,
1,ECON,ECON-U01A,Desarrollo Universitario,0,,
2,CBAS,CBAS-M02A,Cálculo Integral,4,CBAS-M01A,
3,CBAS,CBAS-M05A,Cálculo Vectorial,4,CBAS-M02A,CBAS-M03A


---

# Uso

### Generar HTMLs desde Excel

```bash
python3 generar_mallas.py pensum.xlsx --outdir dist
```

### Colores aleatorios por área (reproducibles con semilla):
```bash
python3 generar_mallas.py pensum.xlsx --randomize-colors --seed 123
```

### Pruebas internas de parseo/colores:
```bash
python3 generar_mallas.py --selftest
```
**Salida:** se crearán uno o varios HTML en dist/ con nombre CODIGO.html, donde CODIGO viene del nombre de la hoja.

### Cómo usar el HTML generado**
- Ábrelo con tu navegador (doble clic).
- Los cursos sin prerrequisitos aparecen disponibles.
- Al marcarlos como aprobados, se van desbloqueando los dependientes.

**En el encabezado:**
- Campos Código estudiante y Nombre (persisten por programa).
- Reiniciar (borra el progreso local del programa).
- Exportar CSV (incluye detalle por curso + resumen + meta del estudiante).
- Tema: Claro/Oscuro (persistente por programa).

**Qué incluye el CSV**
- Meta inicial:
- Programa (código)
- Código estudiante
- Nombre
- Fecha de export
- Detalle por curso:
  - ID, Nombre, Área, Nivel, Créditos, Prerrequisitos, Estado
- Resumen:
  - Créditos aprobados / totales / faltantes
  - Cursos aprobados / pendientes
- Colores por área
  - Se asignan automáticamente a cada AREA.
  - Si hay más de 15 áreas, la paleta se expande (HSL → HEX) para mantener buena separación visual.
  - Con --randomize-colors --seed N obtendrás colores aleatorios reproducibles (misma semilla → mismos colores).
**Persistencia**
- Progreso, tema y datos del estudiante se guardan en localStorage por programa (clave basada en el código del programa).
- Cambiar de CODIGO crea un “espacio” nuevo de datos.

## Solución de problemas
- Cursos que no se desbloquean: asegúrate de que PRE* use exactamente los mismos IDs que la columna ID (formato AREA-COD, mayúsculas y guiones coherentes).
- Faltan columnas: confirma que existan al menos LEVEL, ID, NAME, CREDITS, AREA.
- Título o nombre de archivo inesperado: revisa que el nombre de la hoja tenga el formato TITULO - CODIGO.

## Créditos

Desarrollado por David Sierra Porta para facilitar la planificación de mallas en Ciencia de Datos y otros programas académicos.
¡Se agradecen issues y pull requests!
