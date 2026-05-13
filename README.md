# Md2Word

Conversor de Markdown (`.md`) a documento Word (`.docx`) usando una plantilla corporativa como base.

## Requisitos

- **Python 3.7+**
- **python-docx** — manipulación de documentos Word
- **Pillow** — redimensionado de imágenes
- **lxml** — manipulación XML directa
- **pywin32** — actualización automática del índice via Word COM (solo Windows)
- **@mermaid-js/mermaid-cli** (opcional) — renderizado de diagramas Mermaid a PNG

### Instalación

```bash
pip install -r requirements.txt

# Opcional: para soporte Mermaid
npm install -g @mermaid-js/mermaid-cli
```

## Estructura de carpetas

```
Md2Word\
├── Md2Word.bat                         → Lanzador principal Windows
├── Md2Word.py                          → Script principal de conversión
├── PrepararPlantilla.py                → Analiza la plantilla y guarda mapa de estilos
├── GetTitle.ps1                        → Extrae título H1 del .md
├── Settings.json                       → Configuración general
├── requirements.txt                    → Dependencias Python
├── README.md                           → Esta guía
├── PROMPT GENERACION MD ENTRADA.md     → Ejemplo de prompt para generar el Markdown de entrada con IA
└── docs\
    ├── ejemplo_entrada.md              → Ejemplo de documento Markdown a convertir
    ├── ejemplo_plantilla.docx          → Ejemplo de plantilla corporativa
    └── img\                            → Imágenes referenciadas desde el .md
```

## Settings.json

```json
{
  "General": {
    "Author":    "Nombre Apellido",
    "AuthorRole": "Cargo",
    "ExtraInfo": "Datos complementarios",
    "WebUrl":    "empresa.com"
  },
  "Python": {
    "InterpreterPath": "C:\\Python3\\python.exe"
  },
  "Styles": {
    "H1": "Heading 1",
    "H2": "Heading 2",
    "H3": "Heading 3",
    "Body": "Normal",
    "Bullet": "List Bullet",
    "Code": "Code",
    "Caption": "Caption",
    "Footer": "Footer"
  }
}
```

| Campo | Descripción |
|-------|-------------|
| `General.Author` | Nombre del autor (portada) |
| `General.AuthorRole` | Cargo del autor (portada) |
| `General.ExtraInfo` | Datos complementarios (portada) |
| `General.WebUrl` | URL que aparece en el footer de la portada |
| `Python.InterpreterPath` | Ruta al intérprete Python |
| `Styles.*` | Mapeo de roles a nombres de estilo Word de la plantilla |

## Flujo de trabajo

### Formato del Markdown de entrada

Se puede usar el fichero `PROMPT GENERACION MD ENTRADA.md` en la raíz del proyecto para que una IA genere un MD con la referencia de sintaxis correcta. Ver también `docs\ejemplo_entrada.md`.

### Opción 1: Flujo Automático (recomendado)

1. **Ejecutar el script principal**:
   ```bash
   Md2Word.bat
   ```
   Este comando realiza los siguientes pasos automáticamente:
   - Ejecuta `PrepararPlantilla.py` para analizar la plantilla y actualizar `Settings.json`.
   - Convierte el archivo Markdown más reciente en `docs\` a un documento Word usando la plantilla más reciente.
   - Genera el archivo de salida en `docs\` con un nombre basado en el título del Markdown.

### Opción 2: Flujo Manual

1. **Preparar la plantilla** (una sola vez por plantilla):
   ```bash
   python PrepararPlantilla.py docs\plantilla.docx
   ```
   Esto analiza los estilos de la plantilla y actualiza `Settings.json`.

2. **Revisar Settings.json** — ajustar datos de autor/empresa y confirmar el mapeo de estilos.

3. **Ejecutar la conversión**:
   ```bash
   python Md2Word.py docs\mi_documento.md docs\plantilla.docx docs\salida.docx
   ```
   Esto convierte el archivo Markdown especificado en un documento Word usando la plantilla indicada.

## Comportamiento con la plantilla

El script espera una plantilla `.docx` con esta estructura de secciones:

| Sección | Contenido |
|---------|-----------|
| 0 | Portada (con shapes flotantes de título y autor) |
| 1 | Índice (SDT TOC) + inicio del contenido |
| 2 | Página de cierre |

El contenido generado se inserta **entre el índice y la página de cierre**, preservando ambas intactas.

### Portada
La portada se rellena automáticamente con datos de `Settings.json`:
- **Título del documento** → extraído del `# H1` del Markdown
- **Autor - Cargo** → `General.Author` + `General.AuthorRole`
- **Datos complementarios** → `General.ExtraInfo`
- **URL footer de portada** → `General.WebUrl`

### Índice
Se genera un índice automático (TOC) con:
- Numeración de secciones (1, 2, 2.1, 2.2…)
- Hipervínculos internos (Ctrl+clic navega a la sección)
- Números de página actualizados automáticamente via Word COM

El título "Índice" usa el formato visual del Heading 1 de la plantilla pero con estilo `Normal`, para **no aparecer dentro del propio índice**.

## Sintaxis Markdown soportada

| Sintaxis | Resultado |
|----------|-----------|
| `# Título` | Título del documento (portada), no se repite en el cuerpo |
| `## Sección` | Heading 1 numerado en el índice |
| `### Subsección` | Heading 2 numerado en el índice |
| `**texto**` / `__texto__` | Negrita |
| `*texto*` / `_texto_` | Cursiva |
| `~~texto~~` | Tachado |
| `` `código` `` | Código inline |
| `[texto](url)` | Enlace (texto sin hipervínculo externo) |
| `![alt](ruta)` | Imagen insertada y escalada |
| `- texto` / `  - texto` | Viñeta / subviñeta |
| `1. texto` | Lista numerada |
| ` ```lang … ``` ` | Bloque de código |
| ` ```mermaid … ``` ` | Diagrama Mermaid renderizado (caption "Diagrama Mermaid N") |
| `\| col \| col \|` | Tabla Word |
| `---` | Separador |
| Sección `# Índice` / `# Index` | Marca dónde insertar el TOC |

## Diagramas Mermaid

Si `mmdc` está instalado, los bloques ` ```mermaid ` se renderizan como imagen PNG e insertan en el documento con un caption numerado automáticamente:

```
Diagrama Mermaid 1
Diagrama Mermaid 2
…
```

Si `mmdc` no está disponible o falla, el bloque se inserta como código monoespaciado.

## Modo --debug

```bash
python Md2Word.py entrada.md plantilla.docx salida.docx --debug
```

Muestra estilos de la plantilla, punto de inserción, elementos parseados del Markdown y resultado del TOC.

## Nombre del fichero de salida

Cuando se usa `Md2Word.bat`:
- El nombre se extrae del primer `# Título` del Markdown (via `GetTitle.ps1`).
- Si el archivo ya existe, se añade sufijo `_1`, `_2`, etc.