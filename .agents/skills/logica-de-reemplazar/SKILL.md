---
name: logica-de-reemplazar
description: >-
  Logica estandar del proyecto App_Informes_Offline para rellenar/sobrescribir
  plantillas Word (.docx) con datos extraidos de PDFs y el padron Excel.
  Usar este skill cuando se necesite: agregar nuevas variables a la plantilla,
  modificar el mecanismo de reemplazo, solucionar que datos no se esten
  sobrescribiendo, o replicar este patron en nuevas plantillas Word del proyecto.
---

# Logica de Reemplazar - App Informes Offline

Patron estandar probado y validado para rellenar plantillas Word con datos
dinamicos en el proyecto App_Informes_Offline.

Archivos de referencia:
- generador_word.py  (Motor principal de generacion)
- procesador.py      (Extrae el contexto y llama al generador)

---

## Arquitectura: Dos Pasos en Cascada

El reemplazo opera en DOS PASOS CONSECUTIVOS para garantizar maxima cobertura:

  PLANTILLA .docx
        |
        v
  [PASO 1] docxtpl (Jinja2)
        |  Reemplaza etiquetas {{ VAR }} usando el contexto expandido.
        |  Si falla -> copia la plantilla sin modificar (no aborta).
        |
        v
  [PASO 2] python-docx (reemplazo directo)
        |  Reemplaza etiquetas residuales en TODAS las sintaxis posibles.
        |  + Reemplazos fallback de texto plano estatico (sin etiquetas).
        |  Cubre: parrafos, tablas, encabezados, pies de pagina.
        |
        v
  ARCHIVO .docx FINAL (sobrescrito y guardado)

---

## PASO 1 - docxtpl (Jinja2)

La plantilla .docx debe tener etiquetas con sintaxis Jinja2: {{ NOMBRE_VARIABLE }}

```python
from docxtpl import DocxTemplate
tpl = DocxTemplate(ruta_plantilla)
tpl.render(ctx_limpio)
tpl.save(ruta_salida)
```

### Preparacion del Contexto (ctx_limpio)

Antes de renderizar, cada clave del contexto se expande en 3 variantes de casing:

```python
ctx_limpio[k]         = val_str   # original
ctx_limpio[k.lower()] = val_str   # minusculas
ctx_limpio[k.upper()] = val_str   # MAYUSCULAS
ctx_limpio[k.title()] = val_str   # Title Case
```

Para listas (ej: TABLA_CICLOS), cada item dict tambien se expande en sus 3 casings.

---

## PASO 2 - python-docx (Reemplazo Profundo)

### 2a. Patrones de etiquetas soportados

Por cada clave se generan los siguientes patrones (UPPER, lower y Title):

  {{ VAR }}    {{VAR}}    [ VAR ]    [VAR]
  <VAR>        <<VAR>>    ${VAR}     {VAR}

### 2b. Fallback: reemplazo de texto plano estatico

Para plantillas con texto hardcoded (sin etiquetas), se define un listado
de sustituciones directas. IMPORTANTE: actualizar cuando cambie el caso base:

```python
reemplazos_fallback = [
    ("LAIDY SCARLE PANTOJA CUSI",       ctx["NOMBRES_Y_APELLIDOS_VALIDADOS"]),
    ("75551078",                         ctx["DNI_VALIDADO"]),
    ("Universidad Peruana Cayetano Heredia", ctx["INSTITUCION"]),
    ("CAR.OUB-UPCH-1565-2026",          ctx["CODIGO_DOC_IES"]),
    ("17 de julio de 2026",             ctx["FECHA_SOLICITUD_TEXTO"]),
    ("23-08-2021",                       ctx["FECHA_INICIO_SIBEC"]),
    ("14-09-2026",                       ctx["FECHA_FIN_SIBEC"]),
]
```

### 2c. Funcion procesar_parrafo

Aplica PRIMERO etiquetas y LUEGO fallbacks de texto plano:

```python
def procesar_parrafo(p):
    if not p.text:
        return
    texto_nuevo = p.text
    for tag, val in patrones_reemplazo:       # etiquetas {{ }} [ ] < >
        if tag in texto_nuevo:
            texto_nuevo = texto_nuevo.replace(tag, val)
    for old_val, new_val in reemplazos_fallback:  # texto plano hardcoded
        if old_val in texto_nuevo:
            texto_nuevo = texto_nuevo.replace(old_val, new_val)
    if texto_nuevo != p.text:
        p.text = texto_nuevo
```

### 2d. Cobertura de zonas del documento

```python
for p in doc.paragraphs:                           # cuerpo
    procesar_parrafo(p)
for table in doc.tables:                           # tablas del cuerpo
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                procesar_parrafo(p)
for section in doc.sections:
    for p in section.header.paragraphs:            # encabezados
        procesar_parrafo(p)
    for p in section.footer.paragraphs:            # pies de pagina
        procesar_parrafo(p)
    # + tablas dentro de header y footer
```

---

## Localizacion flexible de la plantilla

Para funcionar en desarrollo Y en ejecutable .exe (PyInstaller):

```python
def obtener_ruta_plantilla() -> Path:
    rutas_candidatas = [
        BASE_DIR / "plantillas" / "plantilla_informe.docx",
        BASE_DIR / "dist" / "main" / "plantillas" / "plantilla_informe.docx",
        BASE_DIR / "dist" / "plantillas" / "plantilla_informe.docx",
        Path(__file__).resolve().parent / "plantillas" / "plantilla_informe.docx",
    ]
    for ruta in rutas_candidatas:
        if ruta.is_file():
            return ruta
    return BASE_DIR / "plantillas" / "plantilla_informe.docx"
```

REGLA: Usar siempre obtener_ruta_plantilla() en lugar de una ruta fija.

---

## Alias de Variables

El sistema mapea claves principales a multiples sinonimos para ser resiliente:

```python
alias_map = {
    "NOMBRES_Y_APELLIDOS_VALIDADOS": ["NOMBRES_Y_APELLIDOS", "NOMBRE_BECARIO", ...],
    "DNI_VALIDADO":    ["DNI", "DNI_BECARIO", "NUMERO_DNI", ...],
    "INSTITUCION":     ["IES", "UNIVERSIDAD", "INSTITUTO", ...],
    "EL_LA_BECARIO_A": ["SEXO_ARTICULO_2", "SEXO_BECARIO_A"],
    "DEL_BECARIO_A":   ["DEL_BECARIO", "DE_LA_BECARIA"],
    # Ver generador_word.py para la lista completa
}
```

---

## Variables de Concordancia de Genero

Generadas en procesador.py -> generar_concordancias_genero(sexo):

  Variable            | Femenino             | Masculino
  EL_LA_BECARIO_A     | la becaria           | el becario
  DEL_BECARIO_A       | de la becaria        | del becario
  A_LA_BECARIO_A      | a la becaria         | al becario
  SEXO_CONECTOR       | presentada por una   | presentado por un
  SEXO_ARTICULO_1     | 1 becaria            | 1 becario

---

## Tabla de Ciclos (rellenado dinamico)

```python
# Borrar filas viejas (dejar solo encabezado)
while len(table.rows) > 1:
    table._tbl.remove(table.rows[-1]._tr)

# Agregar filas nuevas desde TABLA_CICLOS
for item in tabla_ciclos:
    new_row = table.add_row()
    new_row.cells[0].text = item.get("momento") or item.get("MOMENTO") or ""
    new_row.cells[1].text = item.get("ciclo")   or item.get("CICLO")   or ""
    new_row.cells[2].text = item.get("semestre") or item.get("SEMESTRE") or ""
```

---

## Como Agregar una Nueva Variable a la Plantilla

1. Definir la variable en procesador.py -> ctx["NUEVA_VARIABLE"] = valor
2. Agregar sus alias en alias_map de generador_word.py (opcional)
3. Editar la plantilla Word colocando {{ NUEVA_VARIABLE }} en el lugar correcto
4. Actualizar el fallback en _reemplazo_profundo_docx si hay texto estatico
5. Actualizar TODAS las copias de la plantilla:
   - plantillas/plantilla_informe.docx
   - dist/main/plantillas/plantilla_informe.docx
   - dist/plantillas/plantilla_informe.docx

---

## Validacion de Que el Reemplazo Funciono

```python
from docx import Document
doc = Document(ruta_archivo_generado)
texto_total = "\n".join(p.text for p in doc.paragraphs)
for t in doc.tables:
    for r in t.rows:
        texto_total += "\n" + " | ".join(c.text.strip() for c in r.cells)

assert "ROMERO SACRAMENTO LUZ MERY"  in texto_total     # nombre real OK
assert "LAIDY SCARLE PANTOJA CUSI"   not in texto_total  # sin residuos OK
```
