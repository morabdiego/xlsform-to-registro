# XLSForm to Registro (Data Dictionary Generator)

[![PyPI version](https://img.shields.io/pypi/v/xlsform-to-registro.svg)](https://pypi.org/project/xlsform-to-registro/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Herramienta autónoma en Python para convertir **cualquier formulario XLSForm estándar** (`.xlsx` o `.xls`) de KoboToolbox, ODK o Enketo en un **Diseño de Registro institucional (Diccionario de Datos / Codebook)** en formato Excel, con formato editorial profesional, estilos visuales estandarizados y hojas complementarias para catálogos extensos.

---

## 🚀 Características principales

- **100% basado en el estándar XLSForm**: No requiere acceso a bases de datos ni conexión a internet. Lee directamente las hojas `survey`, `choices` y `settings`.
- **Detección automática de entidades y repeats**:
  - Las preguntas del nivel principal se agrupan en una hoja principal (ej. `base_hogar` o `base_principal`).
  - Cada bloque `begin_repeat` genera automáticamente una **hoja independiente** (ej. `base_individual`, `gastos`, etc.).
- **Jerarquía temática**: Cada `begin_group` genera una fila separadora de sección con estilo institucional.
- **Traductor de reglas de salto a Universos de referencia**: Analiza la sintaxis de la columna `relevant` (ej. `${edad} >= 18 and ${ocup} = 1`) y la traduce a descripciones comprensibles de población objetivo (ej. `edad >= 18 y ocup = 1`), heredando condiciones de grupos contenedores.
- **Tratamiento inteligente de catálogos extensos**: Si una lista de opciones supera un umbral configurable (por defecto 30 opciones, como barrios, clasificadores o localidades), genera automáticamente una **hoja técnica complementaria** (`catalogo_[nombre]`) y en la hoja principal coloca una referencia clara: `Ver hoja 'catalogo_[nombre]'`.
- **Limpieza de variables dinámicas**: Reemplaza placeholders de ODK tipo `${nombre}` o `${integrante}` por expresiones naturales en español como *"la persona entrevistada"* o *"el/la integrante"*.
- **Diseño visual profesional (openpyxl)**:
  - Encabezados institucionales con texto blanco centrado.
  - Celdas combinadas verticalmente para variables policotómicas (`Nombre`, `Etiqueta` y `Universo`).
  - Bordes delgados y ajuste calibrado de anchos de columna.
  - Cuadrícula de Excel siempre visible activa.

---

## 📦 Instalación

Instálalo directamente desde [PyPI](https://pypi.org/project/xlsform-to-registro/):

```bash
pip install xlsform-to-registro
```

Requiere Python 3.10 o superior y las dependencias `pandas` y `openpyxl`.

---

## 💻 Uso desde la Línea de Comandos (CLI)

Al instalarlo mediante `pip`, dispones del comando directo `xlsform-to-registro` en tu terminal:

```bash
# Uso básico: genera 'diseño_registro_<nombre>.xlsx'
xlsform-to-registro ruta/a/tu_formulario.xlsx

# Especificar archivo de salida
xlsform-to-registro ruta/a/tu_formulario.xlsx -o diseño_registro.xlsx

# Filtrar por idioma específico (si el XLSForm es multilingüe)
xlsform-to-registro formulario.xlsx --lang es

# Cambiar el color de los encabezados (código hexadecimal sin '#')
xlsform-to-registro formulario.xlsx --header-color 1B365D

# Ajustar el umbral para extraer catálogos extensos a hojas separadas (ej. más de 20 opciones)
xlsform-to-registro formulario.xlsx --catalogue-threshold 20

# Incluir variables de tipo calculate o note
xlsform-to-registro formulario.xlsx --include-calculate --include-notes
```

*(También puedes invocarlo como módulo: `python -m xlsform_to_registro ruta/a/tu_formulario.xlsx`)*.

### Opciones disponibles en CLI:

| Argumento | Descripción | Valor por defecto |
| :--- | :--- | :--- |
| `xlsform` | Ruta al archivo XLSForm de entrada (`.xlsx` o `.xls`). | *(Obligatorio)* |
| `-o`, `--output` | Ruta del archivo Excel de salida. | `diseño_registro_<nombre>.xlsx` |
| `--lang` | Idioma o sufijo de etiqueta a utilizar (ej. `es`, `Español`, `en`). | Primer label disponible |
| `--header-color` | Color hexadecimal de fondo para encabezados y secciones. | `64b692` |
| `--catalogue-threshold` | Cantidad de opciones a partir de la cual una lista pasa a hoja separada. | `30` |
| `--main-sheet-name` | Nombre para la hoja principal fuera de repeats. | `base_hogar` |
| `--include-calculate` | Si se incluye, documenta variables de cálculo interno (`calculate`). | `False` |
| `--include-notes` | Si se incluye, documenta notas (`note`). | `False` |
| `--no-na-code` | Desactiva la adición automática de la opción `0: No corresponde` en preguntas con salto. | `False` |

---

## 🐍 Uso como Librería en Python

Puedes integrar el generador en tus propios scripts o pipelines de datos:

```python
from xlsform_to_registro import XLSFormToRegistro

# Inicializar convertidor
converter = XLSFormToRegistro(
    xlsform_path="tu_formulario.xlsx",
    lang="es",
    header_color="333399",
    catalogue_threshold=30,
    main_sheet_name="base_hogar",
    add_not_applicable_code=True,
    custom_replacements={
        "nombre": "la persona entrevistada",
        "organizacion": "la organización comunitaria"
    }
)

# Exportar a Excel
output_file = converter.to_excel("diseño_registro_modelo.xlsx")
print(f"Generado exitosamente en: {output_file}")
```

---

## 📁 Estructura del Proyecto

```
xlsform-to-registro/
├── .github/
│   └── workflows/
│       ├── ci.yml                 # Integración continua con tests en varias versiones de Python
│       └── publish.yml            # Publicación automática a PyPI con Trusted Publishing
├── src/
│   └── xlsform_to_registro/
│       ├── __init__.py            # Exporta XLSFormToRegistro y __version__
│       ├── __main__.py            # Permite ejecutar con python -m xlsform_to_registro
│       ├── cli.py                 # Punto de entrada para la consola (CLI)
│       ├── generator.py           # Orquestador del flujo de parsing y renderizado
│       ├── core/
│       │   ├── parser.py          # Lógica de lectura de survey, choices y settings
│       │   ├── schema.py          # Dataclasses (Variable, Section, Sheet, Choice)
│       │   ├── relevant_parser.py # Traductor de expresiones relevant a universos
│       │   └── text_cleaner.py    # Limpieza de HTML y variables interpoladas
│       └── exporters/
│           └── excel_styler.py    # Motor de maquetado en openpyxl con estilos institucionales
├── tests/
│   ├── test_package.py            # Tests de integridad de importaciones
│   └── test_cli.py                # Tests del comando de consola
├── pyproject.toml                 # Configuración de empaquetado (Hatchling) y metadatos
├── LICENSE                        # Licencia MIT
└── README.md                      # Documentación del proyecto
```

---

## 📋 Mapeo de tipos de datos de XLSForm

| Tipo en XLSForm | Tratamiento en el Diseño de Registro |
| :--- | :--- |
| `select_one [lista]` | Lista cada opción de `choices` con su código y etiqueta. Si tiene salto, añade `0: No corresponde`. |
| `select_multiple [lista]` | Idem `select_one`, aclarando que admite selección múltiple. |
| `select_one ... or_other` | Incorpora automáticamente la opción `other: Otro/a (especificar)`. |
| Listas con $\ge 30$ opciones | Genera una hoja técnica `catalogo_[lista]` con todas las columnas adicionales (ej. provincia, localidad) y coloca referencia en la hoja principal. |
| `integer` | Valor: `...` \| Etiqueta de valor: `Valor numérico entero`. |
| `decimal` | Valor: `...` \| Etiqueta de valor: `Valor numérico decimal`. |
| `text` | Valor: `...` \| Etiqueta de valor: `Texto alfanumérico libre`. |
| `date`, `today` | Valor: `...` \| Etiqueta de valor: `Fecha en formato AAAA-MM-DD`. |
| `time` | Valor: `...` \| Etiqueta de valor: `Hora en formato HH:MM`. |
| `geopoint` | Valor: `...` \| Etiqueta de valor: `Coordenadas GPS (Latitud, Longitud, Altitud, Precisión)`. |
| `calculate` | Valor: `...` \| Etiqueta de valor: `Variable calculada internamente: [fórmula]`. |

---

## 📄 Licencia

Este proyecto se distribuye bajo la licencia [MIT](LICENSE).
