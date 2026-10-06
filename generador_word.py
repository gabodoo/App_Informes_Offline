from __future__ import annotations

import sys
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Callable

import docx
from docx import Document
from docx.shared import Pt
from docxtpl import DocxTemplate


LogCallback = Callable[[str], None]


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "main" and exe_dir.parent.name.lower() == "dist":
            return exe_dir.parent.parent
        if exe_dir.name.lower() == "dist":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()


def obtener_ruta_plantilla() -> Path:
    """Busca y retorna la ruta existente de plantilla_informe.docx probando múltiples ubicaciones."""
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


PLANTILLA_PATH = obtener_ruta_plantilla()
SALIDA_DIR = BASE_DIR / "Informes_Generados"


def formatear_nombre_ies(nombre: str) -> str:
    """Formatea el nombre de la IES a tipo título respetando conectores en minúsculas.
    Convierte conectores como 'de', 'del', 'y', 'e', 'la', 'las', 'el', 'los', 'en', 'para', 'por', 'con', 'a' a minúsculas.
    Ej: 'UNIVERSIDAD PERUANA DE CIENCIAS APLICADAS' -> 'Universidad Peruana de Ciencias Aplicadas'
    Ej: 'PONTIFICIA UNIVERSIDAD CATÓLICA DEL PERÚ' -> 'Pontificia Universidad Católica del Perú'"""
    if not nombre:
        return ""
    nombre_clean = str(nombre).strip()
    # Si viene incompleto como 'Pontificia Universidad' o 'PUCP', expandir al nombre oficial completo
    if re.search(r"^\s*Pontificia\s+Universidad(?:\s+Cat[oó]lica)?\s*$", nombre_clean, re.IGNORECASE) or re.search(r"\bPUCP\b", nombre_clean, re.IGNORECASE):
        return "Pontificia Universidad Católica del Perú"
    if "Pontificia Universidad" in nombre_clean and "Católica del Perú" not in nombre_clean and "Catolica del Peru" not in nombre_clean:
        return "Pontificia Universidad Católica del Perú"

    palabras = nombre_clean.split()
    if not palabras:
        return ""

    conectores = {"de", "del", "y", "e", "la", "las", "el", "los", "en", "para", "por", "con", "a"}
    palabras_fmt = []
    for i, p in enumerate(palabras):
        p_low = p.lower()
        if i == 0:
            palabras_fmt.append(p_low.capitalize())
        elif p_low in conectores:
            palabras_fmt.append(p_low)
        else:
            palabras_fmt.append(p_low.capitalize())

    return " ".join(palabras_fmt)


class GeneradorWord:
    """Generador híbrido ultra-robusto que aplica docxtpl Y reemplazo directo en python-docx
    cubriendo todas las variantes de sintaxis: {{VAR}}, [VAR], <VAR>, <<VAR>>, ${VAR}, {VAR}."""

    def generar(self, contexto: dict, log: LogCallback | None = None) -> Path:
        _log = log or (lambda msg: None)

        ruta_plantilla = obtener_ruta_plantilla()
        if not ruta_plantilla.exists():
            raise FileNotFoundError(
                f"No se encontró la plantilla en: {ruta_plantilla.resolve()}\n"
                "Coloca 'plantilla_informe.docx' dentro de la carpeta 'plantillas'."
            )

        SALIDA_DIR.mkdir(parents=True, exist_ok=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        sigedo_val = str(contexto.get("NUMERO_SIGEDO", ""))
        sigedo_corto = sigedo_val.split("-")[0] if sigedo_val else f"informe_{ts}"
        nombre_archivo = f"{sigedo_corto}_informe.docx"
        ruta_salida = SALIDA_DIR / nombre_archivo

        _log(f"  Plantilla: {ruta_plantilla.resolve()}")
        _log(f"  Salida: {ruta_salida.resolve()}")

        # 1. Mapeo exhaustivo de claves y alias
        alias_map = {
            "NOMBRES_Y_APELLIDOS_VALIDADOS": [
                "NOMBRES_Y_APELLIDOS", "NOMBRE_Y_APELLIDOS", "APELLIDOS_Y_NOMBRES",
                "NOMBRE_BECARIO", "NOMBRE_BECARIA", "NOMBRE_ESTUDIANTE", "BECARIO",
                "BECARIA", "ESTUDIANTE", "ALUMNO", "ALUMNA", "NOMBRES", "APELLIDOS",
                "NOMBRE_COMPLETO", "DATOS_BECARIO", "BENEFICIARIO", "POSTULANTE"
            ],
            "DNI_VALIDADO": [
                "DNI", "DNI_BECARIO", "NUMERO_DNI", "NRO_DNI", "DOCUMENTO",
                "DOC_IDENTIDAD", "NUM_DNI", "NRO_DOC", "NUM_DOC"
            ],
            "BECA_Y_CONVOCATORIA_VALIDADA": [
                "BECA_Y_CONVOCATORIA", "PROGRAMA_Y_CONVOCATORIA", "BECA",
                "PROGRAMA_BECA", "PROGRAMA", "CONVOCATORIA", "MODALIDAD", "BECA_CONVOCATORIA"
            ],
            "FECHA_INICIO_SIBEC": [
                "INICIO_SIBEC", "FECHA_INICIO", "F_INICIO", "INICIO_BECA",
                "PERIODO_INICIO", "FECHA_INICIO_ESTUDIOS", "INICIO"
            ],
            "FECHA_FIN_SIBEC": [
                "FIN_SIBEC", "FECHA_FIN", "F_FIN", "FIN_BECA",
                "PERIODO_FIN", "FECHA_FIN_ESTUDIOS", "FIN"
            ],
            "FECHA_SOLICITUD_TEXTO": [
                "FECHA_SOLICITUD", "FECHA_SOL", "FECHA_INGRESO", "FECHA_REGISTRO"
            ],
            "NOMBRE_INFORME_SUCCOR": [
                "INFORME_SUCCOR", "INFORME", "NUMERO_INFORME", "NUM_INFORME", "NRO_INFORME"
            ],
            "NUMERO_SIGEDO": [
                "NUM_SIGEDO", "SIGEDO", "EXPEDIENTE_SIGEDO", "EXPEDIENTE"
            ],
            "SEMESTRE_SOLICITADO": [
                "SEMESTRE", "SEMESTRE_SOLICITUD", "PERIODO_SOLICITADO"
            ],
            "SEMESTRE_ANTERIOR": [
                "SEMESTRE_PREVIO", "PERIODO_ANTERIOR"
            ],
            "RJD_ADJUDICACION": [
                "RJD", "RESOLUCION", "RESOLUCION_RJD"
            ],
            "INSTITUCION": [
                "IES", "UNIVERSIDAD", "INSTITUTO", "INSTITUCION_EDUCATIVA"
            ],
            "CARRERA": [
                "ESPECIALIDAD", "PROGRAMA_ESTUDIOS", "CARRERA_PROFESIONAL"
            ],
            "FECHA_MATRICULA": [
                "MATRICULA", "FECHA_DE_MATRICULA"
            ],
            "FECHA_INICIO_ESTUDIOS": [
                "INICIO_ESTUDIOS", "FECHA_INICIO_CLASES"
            ],
            "ESTADO_PROCEDENCIA": [
                "PROCEDENCIA", "EVALUACION_PROCEDENCIA", "PROCEDENTE_OBSERVADO"
            ],
            "CODIGO_DOC_IES": [
                "CODIGO_IES", "CARTA_IES", "OFICIO_IES", "DOCUMENTO_IES"
            ],
            "FECHA_DOC_IES_TEXTO": [
                "FECHA_DOC_IES", "FECHA_CARTA_IES"
            ],
            "CURSOS_PENDIENTES": [
                "CURSOS", "CURSOS_PENDIENTES_TEXTO"
            ],
            "FECHA_ACTUAL_TEXTO": [
                "FECHA_ACTUAL", "FECHA_HOY", "HOY", "FECHA"
            ],
            "TABLA_CICLOS": [
                "CICLOS", "tabla_ciclos", "ciclos", "TABLA_DE_CICLOS"
            ],
            "EL_LA_BECARIO_A": [
                "SEXO_ARTICULO_2", "SEXO_BECARIO_A"
            ],
            "DEL_BECARIO_A": [
                "DEL_BECARIO", "DE_LA_BECARIA"
            ],
            "A_LA_BECARIO_A": [
                "AL_BECARIO", "A_LA_BECARIA"
            ],
            "BECA_TITULO": [
                "BECA_TITLE_CASE", "BECA_NOMBRE"
            ],
            "REFERENCIA_A": [
                "REF_A", "REFERENCIA_SOLICITUD"
            ],
            "REFERENCIA_B": [
                "REF_B", "REFERENCIA_INFORME"
            ],
            "NOMBRE_PRIMERO_NOMBRES": [
                "NOMBRE_TITULO", "NOMBRE_NOMBRES_PRIMERO"
            ],
        }

        # 2. Expandir contexto con claves y alias
        ctx_expandido = dict(contexto)
        for clave_principal, lista_alias in alias_map.items():
            if clave_principal in contexto:
                val = contexto[clave_principal]
                for alias in lista_alias:
                    if alias not in ctx_expandido:
                        ctx_expandido[alias] = val

        # 3. Preparar diccionario de reemplazos plano para cadenas (ignorando listas/dicts)
        # NOTA: CURSOS_PENDIENTES y sus alias se excluyen del mapa de texto plano porque
        # contienen saltos de línea y deben expandirse en párrafos reales por _expandir_cursos_pendientes (PASO 3).
        CLAVES_EXCLUIDAS_REEMPLAZO = {
            "CURSOS_PENDIENTES", "CURSOS", "CURSOS_PENDIENTES_TEXTO",
            "cursos_pendientes", "cursos", "cursos_pendientes_texto",
            "Cursos_Pendientes", "Cursos", "Cursos_Pendientes_Texto",
        }

        mapa_reemplazos = {}
        ctx_limpio = {}

        for k, v in ctx_expandido.items():
            if v is None:
                val_str = ""
                ctx_limpio[k] = ""
            elif isinstance(v, list):
                val_list = []
                for item in v:
                    if isinstance(item, dict):
                        item_var = dict(item)
                        for ik, iv in list(item.items()):
                            item_var[ik.lower()] = iv
                            item_var[ik.upper()] = iv
                            item_var[ik.title()] = iv
                        val_list.append(item_var)
                    else:
                        val_list.append(item)
                ctx_limpio[k] = val_list
                continue
            elif isinstance(v, dict):
                ctx_limpio[k] = v
                continue
            else:
                val_str = str(v)
                ctx_limpio[k] = val_str

            ctx_limpio[k.lower()] = val_str
            ctx_limpio[k.upper()] = val_str
            ctx_limpio[k.title()] = val_str

            # Excluir claves de cursos del mapa de reemplazo de texto plano
            if k.upper() not in CLAVES_EXCLUIDAS_REEMPLAZO and k not in CLAVES_EXCLUIDAS_REEMPLAZO:
                mapa_reemplazos[k.upper()] = val_str

        # PASO 1: Renderizado mediante docxtpl (Jinja2)
        # CURSOS_PENDIENTES se sustituye por un placeholder unico detectable.
        # Asi Jinja2 no lo borra ni lo convierte en texto plano con \n embebidos.
        # El PASO 3 (_expandir_cursos_pendientes) se encarga de expandirlo en parrafos reales.
        _PLACEHOLDER_CURSOS = "__CURSOS_PH__"
        ctx_tpl = {k: v for k, v in ctx_limpio.items() if k not in CLAVES_EXCLUIDAS_REEMPLAZO}
        ctx_tpl["CURSOS_PENDIENTES"] = _PLACEHOLDER_CURSOS
        ctx_tpl["cursos_pendientes"] = _PLACEHOLDER_CURSOS
        ctx_tpl["Cursos_Pendientes"] = _PLACEHOLDER_CURSOS
        try:
            tpl = DocxTemplate(ruta_plantilla)
            tpl.render(ctx_tpl)
            tpl.save(ruta_salida)
        except Exception as err:
            _log(f"  [AVISO docxtpl]: {err}. Se aplicará reemplazo directo con python-docx.")
            shutil.copy(ruta_plantilla, ruta_salida)

        # PASO 2: Reemplazo profundo directo con python-docx para TODAS las sintaxis y fallback de cadenas
        self._reemplazo_profundo_docx(ruta_salida, mapa_reemplazos, contexto)

        # PASO 3: Expansión especial de CURSOS_PENDIENTES en párrafos numerados individuales
        cursos_lista = contexto.get("CURSOS_PENDIENTES", "")
        if not cursos_lista:
            cursos_lista = "(No se detectaron cursos pendientes)"
        self._expandir_cursos_pendientes(ruta_salida, cursos_lista)

        _log("  Documento Word guardado correctamente.")
        return ruta_salida

    def _reemplazo_profundo_docx(self, ruta_docx: Path, mapa_reemplazos: dict[str, str], contexto: dict) -> None:
        """Abre el archivo Word y reemplaza cualquier etiqueta residual o cadena estática en párrafos, tablas, encabezados y pies."""
        doc = Document(ruta_docx)
        tabla_ciclos = contexto.get("TABLA_CICLOS", [])

        patrones_reemplazo = []
        for clave, valor in mapa_reemplazos.items():
            if not clave:
                continue
            nombres = {clave, clave.lower(), clave.title()}
            for n in nombres:
                patrones_reemplazo.extend([
                    (f"{{{{{n}}}}}", valor),
                    (f"{{{{ {n} }}}}", valor),
                    (f"[{n}]", valor),
                    (f"[ {n} ]", valor),
                    (f"<{n}>", valor),
                    (f"<<{n}>>", valor),
                    (f"${{{n}}}", valor),
                    (f"{{{n}}}", valor),
                ])

        # Mapeo de reemplazos estáticos de resguardo (para plantillas antiguas sin etiquetas)
        reemplazos_fallback = []
        if contexto.get("NOMBRES_Y_APELLIDOS_VALIDADOS"):
            val_nom = str(contexto["NOMBRES_Y_APELLIDOS_VALIDADOS"])
            reemplazos_fallback.extend([
                ("LAIDY SCARLE PANTOJA CUSI", val_nom),
                ("Laidy Scarle Pantoja Cusi", val_nom),
            ])
        if contexto.get("DNI_VALIDADO"):
            val_dni = str(contexto["DNI_VALIDADO"])
            reemplazos_fallback.append(("75551078", val_dni))
        if contexto.get("INSTITUCION"):
            val_ies_raw = str(contexto["INSTITUCION"])
            val_ies_fmt = formatear_nombre_ies(val_ies_raw)
            reemplazos_fallback.extend([
                ("Universidad Peruana Cayetano Heredia", val_ies_fmt),
                ("UNIVERSIDAD PERUANA CAYETANO HEREDIA", val_ies_raw.upper()),
                ("Universidad Peruana de Ciencias Aplicadas", val_ies_fmt),
                ("UNIVERSIDAD PERUANA DE CIENCIAS APLICADAS", val_ies_raw.upper()),
                ("Universidad Peruana De Ciencias Aplicadas", val_ies_fmt),
            ])
        if contexto.get("CODIGO_DOC_IES"):
            val_doc_ies = str(contexto["CODIGO_DOC_IES"])
            reemplazos_fallback.append(("CAR.OUB-UPCH-1565-2026", val_doc_ies))
        val_succor = str(contexto.get("NOMBRE_INFORME_SUCCOR") or "").strip()
        val_succor_clean = ""
        if val_succor:
            val_succor_clean = re.sub(
                r"^\s*(?:INFORME|Informe)(?:\s+(?:T[EÉ]CNICO|LEGAL|FINAL))?\s*(?:N[°ºo\.\s]*|NRO\.?|N[UÚ]MERO|NUMERO|N\.º|N\.°|N°:|N°\s*:)?\s*[:\s]*",
                "Informe N° ",
                val_succor,
                flags=re.IGNORECASE
            )
            reemplazos_fallback.extend([
                ("Informe Nº 4187-2026-MINEDU/VMGI-PRONABEC-DICONCI-SUCCOR-LIMA", val_succor_clean),
                ("Informe N° 4187-2026-MINEDU/VMGI-PRONABEC-DICONCI-SUCCOR-LIMA", val_succor_clean),
                ("INFORME N° 4187-2026-MINEDU/VMGI-PRONABEC-DICONCI-SUCCOR-LIMA", val_succor_clean),
                ("Informe Nº 4164-2026-MINEDU/VMGI-PRONABEC-DICONCI-SUCCOR-LIMA", val_succor_clean),
                ("Informe N° 4164-2026-MINEDU/VMGI-PRONABEC-DICONCI-SUCCOR-LIMA", val_succor_clean),
                ("INFORME N° 4164-2026-MINEDU/VMGI-PRONABEC-DICONCI-SUCCOR-LIMA", val_succor_clean),
            ])
        if contexto.get("CARRERA"):
            val_carr = str(contexto["CARRERA"])
            reemplazos_fallback.append(("CARRERA PROFESIONAL DE ENFERMERIA", val_carr))
        if contexto.get("BECA_Y_CONVOCATORIA_VALIDADA"):
            val_beca = str(contexto["BECA_Y_CONVOCATORIA_VALIDADA"])
            reemplazos_fallback.extend([
                ("BECA 18 - 2021", val_beca),
                ("Beca 18 - 2021", val_beca.title()),
                ("Beca 18 - Convocatoria 2021", val_beca.title()),
                ("BECA 18 - CONVOCATORIA 2021", val_beca.upper()),
            ])
            
            val_sexo = str(contexto.get("SEXO_ARTICULO_1", "1 becario/a"))
            val_beca_titulo = str(contexto.get("BECA_TITULO", val_beca))
            
            reemplazos_fallback.extend([
                ("1 becaria de la Beca 18 - Convocatoria 2021", f"{val_sexo} de la {val_beca_titulo}"),
                ("1 becario de la Beca 18 - Convocatoria 2021", f"{val_sexo} de la {val_beca_titulo}"),
            ])
        if contexto.get("FECHA_SOLICITUD_TEXTO"):
            val_fsol = str(contexto["FECHA_SOLICITUD_TEXTO"])
            reemplazos_fallback.extend([
                ("17 DE JULIO DE 2026", val_fsol),
                ("17 de julio de 2026", val_fsol),
            ])
        if contexto.get("FECHA_DOC_IES_TEXTO"):
            val_fdoc = str(contexto["FECHA_DOC_IES_TEXTO"])
            reemplazos_fallback.append(("16 de julio del 2026", val_fdoc))
        if contexto.get("FECHA_INICIO_SIBEC"):
            val_fini = str(contexto["FECHA_INICIO_SIBEC"])
            reemplazos_fallback.extend([
                ("23-08-2021", val_fini),
                ("23/08/2021", val_fini),
            ])
        if contexto.get("FECHA_FIN_SIBEC"):
            val_ffin = str(contexto["FECHA_FIN_SIBEC"])
            reemplazos_fallback.extend([
                ("14-09-2026", val_ffin),
                ("14/09/2026", val_ffin),
            ])

        def _reemplazar_texto_run(run, patrones_reemplazo, reemplazos_fallback, nro_inf_gen):
            """Reemplaza texto dentro de un run preservando el formato original."""
            texto_nuevo = run.text
            for tag, val in patrones_reemplazo:
                if tag in texto_nuevo:
                    texto_nuevo = texto_nuevo.replace(tag, val)
            for old_val, new_val in reemplazos_fallback:
                if old_val in texto_nuevo:
                    texto_nuevo = texto_nuevo.replace(old_val, new_val)
            if val_succor_clean:
                patron_succor_generico = re.compile(
                    r"(?:INFORME|Informe)(?:\s+(?:T[EÉ]CNICO|LEGAL|FINAL))?\s*(?:N[°ºo\.\s]*|NRO\.?|N[UÚ]MERO|NUMERO|N\.º|N\.°|N°:|N°\s*:)?\s*[\d]+[-\s\w/\.]*(?:SUCCOR[-\s\w]*LIMA|SUCCOR[-\s\w]*|DICONCI[-\s\w]*SUCCOR)",
                    re.IGNORECASE
                )
                if patron_succor_generico.search(texto_nuevo):
                    texto_nuevo = patron_succor_generico.sub(val_succor_clean, texto_nuevo)
            if nro_inf_gen:
                texto_nuevo = re.sub(r"(INFORME\s+N[º°o]?\s*)\d+(-\d{4}-MINEDU/VMGI-PRONABEC-DIBEC-SUS)", rf"\g<1>{nro_inf_gen}\g<2>", texto_nuevo, flags=re.IGNORECASE)
            if texto_nuevo != run.text:
                run.text = texto_nuevo

        def procesar_parrafo(p):
            if not p.text:
                return
            nro_inf_gen = contexto.get("NUMERO_INFORME_GENERAR", "")
            
            # Primero intentar reemplazo a nivel de run (preserva formato de cada run)
            texto_pre = p.text
            for run in p.runs:
                _reemplazar_texto_run(run, patrones_reemplazo, reemplazos_fallback, nro_inf_gen)

            # Reemplazos y rescates a nivel de párrafo si quedaron huecos o etiquetas
            texto_orig = p.text
            texto_nuevo = texto_orig
            for tag, val in patrones_reemplazo:
                if tag in texto_nuevo:
                    texto_nuevo = texto_nuevo.replace(tag, val)
            for old_val, new_val in reemplazos_fallback:
                if old_val in texto_nuevo:
                    texto_nuevo = texto_nuevo.replace(old_val, new_val)
            if val_succor_clean:
                patron_succor_generico = re.compile(
                    r"(?:INFORME|Informe)(?:\s+(?:T[EÉ]CNICO|LEGAL|FINAL))?\s*(?:N[°ºo\.\s]*|NRO\.?|N[UÚ]MERO|NUMERO|N\.º|N\.°|N°:|N°\s*:)?\s*[\d]+[-\s\w/\.]*(?:SUCCOR[-\s\w]*LIMA|SUCCOR[-\s\w]*|DICONCI[-\s\w]*SUCCOR)",
                    re.IGNORECASE
                )
                if patron_succor_generico.search(texto_nuevo):
                    texto_nuevo = patron_succor_generico.sub(val_succor_clean, texto_nuevo)
                # Rescate contextual para párrafos 18 y 77 si quedaron vacíos por renderizado
                texto_nuevo = re.sub(
                    r"Mediante\s+el\s*(?:,\s*|\s+)la\s+Subdirecci[oó]n",
                    f"Mediante el {val_succor_clean}, la Subdirección",
                    texto_nuevo,
                    flags=re.IGNORECASE
                )
                texto_nuevo = re.sub(
                    r"mediante\s*(?:,\s*|\s+)se\s+concluye",
                    f"mediante {val_succor_clean} se concluye",
                    texto_nuevo,
                    flags=re.IGNORECASE
                )
            if nro_inf_gen:
                texto_nuevo = re.sub(r"(INFORME\s+N[º°o]?\s*)\d+(-\d{4}-MINEDU/VMGI-PRONABEC-DIBEC-SUS)", rf"\g<1>{nro_inf_gen}\g<2>", texto_nuevo, flags=re.IGNORECASE)
            
            if texto_nuevo != texto_orig:
                # Guardar formato del primer run antes de sobrescribir
                fmt_guardado = {}
                if p.runs:
                    r0 = p.runs[0]
                    fmt_guardado = {
                        "name": r0.font.name,
                        "size": r0.font.size,
                        "bold": r0.bold,
                        "italic": r0.italic,
                        "underline": r0.underline,
                    }
                p.text = texto_nuevo
                # Restaurar formato al nuevo run para mantener Arial 11
                if p.runs and fmt_guardado:
                    r = p.runs[0]
                    if fmt_guardado.get("name"):
                        r.font.name = fmt_guardado["name"]
                    if fmt_guardado.get("size"):
                        r.font.size = fmt_guardado["size"]
                    if fmt_guardado.get("bold") is not None:
                        r.bold = fmt_guardado["bold"]
                    if fmt_guardado.get("italic") is not None:
                        r.italic = fmt_guardado["italic"]
                    if fmt_guardado.get("underline") is not None:
                        r.underline = fmt_guardado["underline"]

        # 1. Párrafos principales
        for p in doc.paragraphs:
            procesar_parrafo(p)

        # 2. Tablas principales
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        procesar_parrafo(p)

        # 3. Encabezados y pies de página
        for section in doc.sections:
            for p in section.header.paragraphs:
                procesar_parrafo(p)
            for table in section.header.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            procesar_parrafo(p)

            for p in section.footer.paragraphs:
                procesar_parrafo(p)
            for table in section.footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            procesar_parrafo(p)

        # 4. Verificación explícita de la sección REFERENCIA en la cabecera (literal b)
        if val_succor_clean and len(doc.tables) > 0:
            for table in doc.tables[:2]:
                for row in table.rows:
                    celdas_txt = [c.text.strip().upper() for c in row.cells]
                    if any("REFERENCIA" in c for c in celdas_txt):
                        for cell in row.cells:
                            if len(cell.paragraphs) >= 2:
                                p_b = cell.paragraphs[1]
                                t_b = p_b.text.strip()
                                if not t_b or t_b in ("b)", "b).", "b.-", "b-", "b") or "NOMBRE_INFORME_SUCCOR" in t_b or "{{" in t_b:
                                    p_b.text = val_succor_clean
                                    if p_b.runs:
                                        p_b.runs[0].font.name = "Arial"
                                        p_b.runs[0].font.size = Pt(11)

        # 4. Rellenado dinámico de Tabla de Ciclos
        # Detectar la tabla por columnas Momento/Ciclo/Semestre y reconstruirla
        # con 'Duración de estudios' fusionado en columna 0 y ordinales en columna 1.
        if tabla_ciclos and isinstance(tabla_ciclos, list) and len(tabla_ciclos) > 0:
            from docx.shared import Pt as _Pt2
            from docx.enum.text import WD_ALIGN_PARAGRAPH as _WAP
            from docx.enum.table import WD_ALIGN_VERTICAL as _WAV

            for table in doc.tables:
                if len(table.rows) == 0:
                    continue
                first_row_texts = [cell.text.strip().upper() for cell in table.rows[0].cells]
                tiene_momento = any("MOMENTO" in t for t in first_row_texts)
                tiene_ciclo = any("CICLO" in t for t in first_row_texts)
                tiene_semestre = any("SEMESTRE" in t for t in first_row_texts)
                if not (tiene_momento and tiene_ciclo and tiene_semestre):
                    continue

                # Limpiar filas de datos (mantener solo encabezado)
                while len(table.rows) > 1:
                    tr = table.rows[-1]._tr
                    table._tbl.remove(tr)

                # Agregar filas de ciclos (col 0 vacía, col 1 = ordinal, col 2 = semestre)
                for item in tabla_ciclos:
                    new_row = table.add_row()
                    nc0 = new_row.cells[0] if len(new_row.cells) > 0 else None
                    nc1 = new_row.cells[1] if len(new_row.cells) > 1 else None
                    nc2 = new_row.cells[2] if len(new_row.cells) > 2 else None

                    val_ciclo = str(
                        item.get("ciclo") or item.get("CICLO") or item.get("Ciclo") or ""
                    )
                    val_semestre = str(
                        item.get("semestre") or item.get("SEMESTRE") or item.get("Semestre") or ""
                    )

                    if nc0:
                        nc0.text = ""
                        nc0.vertical_alignment = _WAV.CENTER
                    if nc1:
                        nc1.text = val_ciclo
                        nc1.vertical_alignment = _WAV.CENTER
                        for para in nc1.paragraphs:
                            para.alignment = _WAP.CENTER
                            for run in para.runs:
                                run.font.name = "Arial"
                                run.font.size = _Pt2(11)
                    if nc2:
                        nc2.text = val_semestre
                        nc2.vertical_alignment = _WAV.CENTER
                        for para in nc2.paragraphs:
                            para.alignment = _WAP.CENTER
                            for run in para.runs:
                                run.font.name = "Arial"
                                run.font.size = _Pt2(11)

                # Fusionar celdas de columna 0 en todas las filas de datos
                n_datos = len(tabla_ciclos)
                if n_datos >= 1:
                    from docx.oxml.ns import qn as _qn2
                    celda_inicio = table.cell(1, 0)
                    if n_datos > 1:
                        celda_fin = table.cell(n_datos, 0)
                        celda_inicio.merge(celda_fin)
                    # Escribir 'Duración de estudios' en la celda fusionada
                    celda_inicio.text = "Duración de estudios"
                    celda_inicio.vertical_alignment = _WAV.CENTER
                    for para in celda_inicio.paragraphs:
                        para.alignment = _WAP.CENTER
                        for run in para.runs:
                            run.font.name = "Arial"
                            run.font.size = _Pt2(11)
                    if not celda_inicio.paragraphs[0].runs:
                        run = celda_inicio.paragraphs[0].add_run("Duración de estudios")
                        run.font.name = "Arial"
                        run.font.size = _Pt2(11)

                break  # Solo procesar la primera tabla de ciclos encontrada

        # 5. Validación y llenado garantizado de Cuadro N° 1 (datos del becario)
        # Asegurar: DNI, NOMBRES Y APELLIDOS (sin comas), RJD, BECA Y CONVOCATORIA, INSTITUCIÓN COMPLETA, CARRERA
        from docx.oxml.ns import qn as _qn_b
        from docx.shared import Pt as _Pt_b
        from docx.enum.table import WD_ALIGN_VERTICAL as _WAV_b
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WAP_b

        PALABRAS_TABLA_BECARIO = ("DNI", "BECARIO", "BECA Y CONVOCATORIA", "INSTITUCION", "CARRERA", "CONVOCATORIA", "RJD", "RJ ")
        for table in doc.tables:
            encabezado = " ".join(cell.text for row in table.rows[:2] for cell in row.cells).upper()
            if sum(1 for w in PALABRAS_TABLA_BECARIO if w in encabezado) >= 2:
                # Identificar qué columna corresponde a cada campo
                col_map = {}
                header_row = table.rows[0]
                for c_idx, cell in enumerate(header_row.cells):
                    c_txt = cell.text.upper().strip()
                    if "DNI" in c_txt or "DOC" in c_txt:
                        col_map["dni"] = c_idx
                    elif any(w in c_txt for w in ("BECARIO", "APELLIDOS", "NOMBRES", "ESTUDIANTE", "ALUMNO")):
                        col_map["nombre"] = c_idx
                    elif any(w in c_txt for w in ("RJD", "RJ", "RESOLUCION", "RESOLUCIÓN", "ADJUDICACION", "ADJUDICACIÓN")):
                        col_map["rjd"] = c_idx
                    elif any(w in c_txt for w in ("BECA", "CONVOCATORIA", "MODALIDAD")):
                        col_map["beca"] = c_idx
                    elif any(w in c_txt for w in ("INSTITUCION", "INSTITUCIÓN", "UNIVERSIDAD", "IES")):
                        col_map["institucion"] = c_idx
                    elif any(w in c_txt for w in ("CARRERA", "ESPECIALIDAD", "PROGRAMA")):
                        col_map["carrera"] = c_idx

                if len(table.rows) > 2:
                    header_row_2 = table.rows[1]
                    for c_idx, cell in enumerate(header_row_2.cells):
                        c_txt = cell.text.upper().strip()
                        if "dni" not in col_map and ("DNI" in c_txt or "DOC" in c_txt):
                            col_map["dni"] = c_idx
                        elif "nombre" not in col_map and any(w in c_txt for w in ("BECARIO", "APELLIDOS", "NOMBRES")):
                            col_map["nombre"] = c_idx
                        elif "rjd" not in col_map and any(w in c_txt for w in ("RJD", "RJ", "RESOLUCION", "RESOLUCIÓN")):
                            col_map["rjd"] = c_idx
                        elif "beca" not in col_map and any(w in c_txt for w in ("BECA", "CONVOCATORIA")):
                            col_map["beca"] = c_idx
                        elif "institucion" not in col_map and any(w in c_txt for w in ("INSTITUCION", "INSTITUCIÓN", "UNIVERSIDAD", "IES")):
                            col_map["institucion"] = c_idx
                        elif "carrera" not in col_map and any(w in c_txt for w in ("CARRERA", "ESPECIALIDAD", "PROGRAMA")):
                            col_map["carrera"] = c_idx

                start_row_idx = 1
                if len(table.rows) > 2 and any(w in table.rows[1].cells[0].text.upper() for w in ("DNI", "N°", "ITEM", "BECARIO", "APELLIDOS")):
                    start_row_idx = 2

                val_nom_ord = str(contexto.get("NOMBRES_Y_APELLIDOS_VALIDADOS") or contexto.get("NOMBRE_PRIMERO_NOMBRES") or "").strip().upper()
                if "," in val_nom_ord:
                    p = [x.strip() for x in val_nom_ord.split(",", 1)]
                    val_nom_ord = f"{p[1]} {p[0]}".strip().upper()

                val_dni = str(contexto.get("DNI_VALIDADO") or "").strip()
                val_rjd = str(contexto.get("RJD_ADJUDICACION") or contexto.get("RJ_ADJUDICACION") or "").strip()
                val_beca = str(contexto.get("BECA_Y_CONVOCATORIA_VALIDADA") or contexto.get("BECA_TITULO") or "Beca 18 - Convocatoria 2023").strip()
                val_ies = str(contexto.get("INSTITUCION") or "Pontificia Universidad Católica del Perú").strip()
                if not val_ies or re.search(r"^\s*Pontificia\s+Universidad(?:\s+Cat[oó]lica)?\s*$", val_ies, re.I) or ("Pontificia Universidad" in val_ies and "Católica del Perú" not in val_ies and "Catolica del Peru" not in val_ies):
                    val_ies = "Pontificia Universidad Católica del Perú"
                val_carr = str(contexto.get("CARRERA") or "").strip()

                for r_idx in range(start_row_idx, len(table.rows)):
                    row = table.rows[r_idx]
                    if "dni" in col_map and col_map["dni"] < len(row.cells):
                        c_cell = row.cells[col_map["dni"]]
                        if not c_cell.text.strip() or "{{" in c_cell.text:
                            c_cell.text = val_dni
                    
                    if "nombre" in col_map and col_map["nombre"] < len(row.cells):
                        c_cell = row.cells[col_map["nombre"]]
                        if not c_cell.text.strip() or "{{" in c_cell.text:
                            c_cell.text = val_nom_ord
                        elif "," in c_cell.text:
                            p = [x.strip() for x in c_cell.text.split(",", 1)]
                            c_cell.text = f"{p[1]} {p[0]}".strip().upper()

                    if "rjd" in col_map and col_map["rjd"] < len(row.cells):
                        c_cell = row.cells[col_map["rjd"]]
                        if (not c_cell.text.strip() or "{{" in c_cell.text) and val_rjd:
                            c_cell.text = val_rjd

                    if "beca" in col_map and col_map["beca"] < len(row.cells):
                        c_cell = row.cells[col_map["beca"]]
                        if not c_cell.text.strip() or "{{" in c_cell.text or c_cell.text.strip() in ("-", "--"):
                            c_cell.text = val_beca

                    if "institucion" in col_map and col_map["institucion"] < len(row.cells):
                        c_cell = row.cells[col_map["institucion"]]
                        if not c_cell.text.strip() or "{{" in c_cell.text or ("PONTIFICIA UNIVERSIDAD" in c_cell.text.upper() and "CATÓLICA" not in c_cell.text.upper() and "CATOLICA" not in c_cell.text.upper()):
                            c_cell.text = val_ies

                    if "carrera" in col_map and col_map["carrera"] < len(row.cells):
                        c_cell = row.cells[col_map["carrera"]]
                        if (not c_cell.text.strip() or "{{" in c_cell.text) and val_carr:
                            c_cell.text = val_carr

                    # Posiciones estándar de 7 columnas si no hubo col_map
                    if len(row.cells) >= 7 and not col_map:
                        if not row.cells[1].text.strip() or "{{" in row.cells[1].text:
                            row.cells[1].text = val_nom_ord
                        if not row.cells[2].text.strip() or "{{" in row.cells[2].text:
                            row.cells[2].text = val_dni
                        if (not row.cells[3].text.strip() or "{{" in row.cells[3].text) and val_rjd:
                            row.cells[3].text = val_rjd
                        if not row.cells[4].text.strip() or "{{" in row.cells[4].text or row.cells[4].text.strip() in ("-", "--"):
                            row.cells[4].text = val_beca
                        if not row.cells[5].text.strip() or "{{" in row.cells[5].text or ("PONTIFICIA UNIVERSIDAD" in row.cells[5].text.upper() and "CATÓLICA" not in row.cells[5].text.upper()):
                            row.cells[5].text = val_ies
                        if (not row.cells[6].text.strip() or "{{" in row.cells[6].text) and val_carr:
                            row.cells[6].text = val_carr

                # Formato final para toda la tabla del becario
                for row in table.rows:
                    for cell in row.cells:
                        cell.vertical_alignment = _WAV_b.CENTER
                        for para in cell.paragraphs:
                            para.alignment = _WAP_b.CENTER
                            for run in para.runs:
                                run.font.name = "Arial"
                                run.font.size = _Pt_b(8)
                                run.text = run.text.upper()
                            if not para.runs and para.text.strip():
                                run = para.add_run(para.text.upper())
                                run.font.name = "Arial"
                                run.font.size = _Pt_b(8)
                                for child in list(para._p):
                                    if child.tag != _qn_b("w:r"):
                                        continue
                                    if child is not run._r:
                                        para._p.remove(child)
                break

        # 6. Cuadro de Periodo de Estudios (Inicio y Fin) si existe
        for t in doc.tables:
            encabezado_t = " ".join(cell.text for row in t.rows[:2] for cell in row.cells).upper()
            if ("INICIO" in encabezado_t and "FIN" in encabezado_t) or ("PERIODO" in encabezado_t and "ESTUDIOS" in encabezado_t):
                if sum(1 for w in PALABRAS_TABLA_BECARIO if w in encabezado_t) < 2:
                    for r_idx in range(1, len(t.rows)):
                        row = t.rows[r_idx]
                        for c_idx, cell in enumerate(row.cells):
                            hdr = t.rows[0].cells[c_idx].text.upper() if c_idx < len(t.rows[0].cells) else ""
                            if "INICIO" in hdr and (not cell.text.strip() or "{{" in cell.text):
                                cell.text = str(contexto.get("FECHA_INICIO_SIBEC") or contexto.get("FECHA_INICIO") or "17/08/2026")
                            elif "FIN" in hdr and (not cell.text.strip() or "{{" in cell.text):
                                cell.text = str(contexto.get("FECHA_FIN_SIBEC") or contexto.get("FECHA_FIN") or "15/12/2026")

        # 7. Asegurar que nunca quede 'Pontificia Universidad' sin 'Católica del Perú' en ningún párrafo o tabla
        patron_pucp_incompleto = re.compile(r'\bPontificia\s+Universidad\b(?!\s+Cat[oó]lica\s+del\s+Per[uú])', re.IGNORECASE)
        for p in doc.paragraphs:
            if patron_pucp_incompleto.search(p.text):
                p.text = patron_pucp_incompleto.sub("Pontificia Universidad Católica del Perú", p.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if patron_pucp_incompleto.search(p.text):
                            p.text = patron_pucp_incompleto.sub("Pontificia Universidad Católica del Perú", p.text)

        # 8. Corrección global: Asegurar año en SIGEDO (ej. SIGEDO: 59288 -> 59288-2026)
        for p in doc.paragraphs:
            if "SIGEDO" in p.text.upper():
                if not re.search(r"SIGEDO\s*(?:INTEGRADO)?\s*:\s*\d+-\d{4}", p.text, re.I):
                    p.text = re.sub(r'(SIGEDO\s*(?:INTEGRADO)?\s*:\s*)(\d{4,8})\b', r'\g<1>\g<2>-2026', p.text, flags=re.I)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if "SIGEDO" in p.text.upper():
                            if not re.search(r"SIGEDO\s*(?:INTEGRADO)?\s*:\s*\d+-\d{4}", p.text, re.I):
                                p.text = re.sub(r'(SIGEDO\s*(?:INTEGRADO)?\s*:\s*)(\d{4,8})\b', r'\g<1>\g<2>-2026', p.text, flags=re.I)

        # 9. Reemplazo de Beca en ASUNTO y párrafos si quedó texto plantilla estático
        beca_real_indiv = contexto.get("BECA_TITULO", "") or contexto.get("BECA_Y_CONVOCATORIA_VALIDADA", "")
        if beca_real_indiv:
            for p in doc.paragraphs:
                if "Beca 18 - Convocatoria 2021" in p.text:
                    p.text = p.text.replace("Beca 18 - Convocatoria 2021", beca_real_indiv)
                if "Beca 18 - 2021" in p.text:
                    p.text = p.text.replace("Beca 18 - 2021", beca_real_indiv)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if "Beca 18 - Convocatoria 2021" in p.text:
                                p.text = p.text.replace("Beca 18 - Convocatoria 2021", beca_real_indiv)
                            if "Beca 18 - 2021" in p.text:
                                p.text = p.text.replace("Beca 18 - 2021", beca_real_indiv)

        doc.save(ruta_docx)

    def _expandir_cursos_pendientes(self, ruta_docx: Path, cursos_texto: str) -> None:
        """Detecta en el documento Word el párrafo que contiene el listado de cursos
        pendientes (marcador o texto ya renderizado con saltos de línea) y lo expande
        en párrafos individuales numerados, clonando el formato XML del párrafo original."""
        import copy
        from lxml import etree
        from docx.oxml.ns import qn

        doc = Document(ruta_docx)

        if isinstance(cursos_texto, list):
            cursos_texto = "\n".join(str(c) for c in cursos_texto)
        elif not isinstance(cursos_texto, str):
            cursos_texto = str(cursos_texto or "")

        # Construir lista de líneas de cursos (quitar vacíos)
        lineas = [l.strip() for l in cursos_texto.splitlines() if l.strip()]
        if not lineas:
            doc.save(ruta_docx)
            return

        # Asegurarse de que cada línea tenga su numeración
        lineas_numeradas = []
        for i, linea in enumerate(lineas):
            # Si la línea ya empieza con "N." (ej. "1. Curso"), usarla tal cual
            if re.match(r"^\d+\.\s+", linea):
                lineas_numeradas.append(linea)
            else:
                lineas_numeradas.append(f"{i+1}. {linea}")

        # Buscar el párrafo en el cuerpo que contiene el marcador residual o
        # que ya tiene el texto completo de cursos (todos juntos en un solo párrafo)
        body = doc.element.body
        parrafos_body = [p for p in body.iterchildren() if p.tag.endswith('}p')]

        idx_objetivo = None
        parrafo_objetivo = None

        MARCADORES = (
            "__CURSOS_PH__",          # placeholder inyectado por PASO 1
            "{{ CURSOS_PENDIENTES }}",
            "{{CURSOS_PENDIENTES}}",
            "[ CURSOS_PENDIENTES ]",
            "[CURSOS_PENDIENTES]",
            "<CURSOS_PENDIENTES>",
            "${CURSOS_PENDIENTES}",
            "{CURSOS_PENDIENTES}",
            "{{ cursos_pendientes }}",
            "{{ Cursos_Pendientes }}",
        )

        for idx, p_elem in enumerate(parrafos_body):
            # Obtener texto completo del párrafo
            texto_p = "".join(
                t.text or "" for t in p_elem.iter(qn("w:t"))
            ).strip()
            # Detectar marcador residual O placeholder O texto ya renderizado con el primer curso
            es_marcador = any(m in texto_p for m in MARCADORES)
            # Detectar texto ya renderizado: contiene la primera línea del listado
            primera_linea = lineas_numeradas[0] if lineas_numeradas else ""
            es_renderizado = primera_linea and primera_linea in texto_p and len(texto_p) > len(primera_linea)
            es_solo_primera = primera_linea and texto_p == primera_linea and len(lineas_numeradas) > 1

            if es_marcador or es_renderizado or es_solo_primera:
                idx_objetivo = idx
                parrafo_objetivo = p_elem
                break

        # Eliminar parrafos vacios con numId (lista) que sigan al parrafo objetivo.
        # La plantilla tiene un P_extra vacio [numId=32] despues de {{ CURSOS_PENDIENTES }}
        # que generaria un numero extra en blanco.
        if parrafo_objetivo is not None:
            siguiente = parrafo_objetivo.getnext()
            while siguiente is not None and siguiente.tag.endswith('}p'):
                # Solo eliminar si es un parrafo de lista numerada (tiene numPr) y esta vacio
                p_texto = "".join(t.text or "" for t in siguiente.iter(qn("w:t"))).strip()
                p_pPr = siguiente.find(qn("w:pPr"))
                p_numPr = p_pPr.find(qn("w:numPr")) if p_pPr is not None else None
                if p_numPr is not None and not p_texto:
                    siguiente_sig = siguiente.getnext()
                    siguiente.getparent().remove(siguiente)
                    siguiente = siguiente_sig
                else:
                    break

        if parrafo_objetivo is None or idx_objetivo is None:
            # No se encontró el párrafo objetivo; no modificar
            doc.save(ruta_docx)
            return

        # Clonar el párrafo plantilla (formato: estilo, numeración, indentación, fuente)
        parrafo_base = copy.deepcopy(parrafo_objetivo)

        # Limpiar todos los runs del párrafo base clonado para usarlo como molde
        for r_elem in parrafo_base.findall(qn("w:r")):
            parrafo_base.remove(r_elem)
        # También limpiar w:hyperlink y otros hijos que no sean w:pPr
        for child in list(parrafo_base):
            if child.tag not in (qn("w:pPr"),):
                parrafo_base.remove(child)

        # Obtener el rPr (formato de caracteres) del run original si existe
        rPr_original = None
        runs_orig = parrafo_objetivo.findall(qn("w:r"))
        if runs_orig:
            rPr_original = runs_orig[0].find(qn("w:rPr"))

        def crear_parrafo_curso(texto_curso: str) -> etree._Element:
            """Crea un elemento <w:p> clonado del párrafo base con el texto del curso."""
            p_nuevo = copy.deepcopy(parrafo_base)
            # Crear el run con el texto
            ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            r_elem = etree.SubElement(p_nuevo, qn("w:r"))
            if rPr_original is not None:
                r_elem.append(copy.deepcopy(rPr_original))
            t_elem = etree.SubElement(r_elem, qn("w:t"))
            t_elem.text = texto_curso
            t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            return p_nuevo

        # Reemplazar el párrafo objetivo por los N párrafos de cursos
        # Primero: reemplazar el contenido del párrafo objetivo con el primer curso
        # Luego: insertar los demás inmediatamente después

        # Limpiar runs del párrafo objetivo
        for r_elem in list(parrafo_objetivo.findall(qn("w:r"))):
            parrafo_objetivo.remove(r_elem)
        for child in list(parrafo_objetivo):
            if child.tag not in (qn("w:pPr"),):
                parrafo_objetivo.remove(child)

        # Insertar texto del primer curso en el párrafo objetivo
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        r0 = etree.SubElement(parrafo_objetivo, qn("w:r"))
        if rPr_original is not None:
            r0.append(copy.deepcopy(rPr_original))
        t0 = etree.SubElement(r0, qn("w:t"))
        t0.text = lineas_numeradas[0]
        t0.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        # Insertar los párrafos restantes después del párrafo objetivo
        insert_after = parrafo_objetivo
        for linea in lineas_numeradas[1:]:
            p_nuevo = crear_parrafo_curso(linea)
            insert_after.addnext(p_nuevo)
            insert_after = p_nuevo

        doc.save(ruta_docx)


    def generar_informe_multiple(self, super_contexto: dict, log_callback=None) -> Path:
        def log(msg):
            if log_callback:
                log_callback(msg)
                
        self.ruta_plantilla_informe = get_base_dir() / "plantillas" / "plantilla_informe_multiple.docx"
        if not self.ruta_plantilla_informe.exists():
            log(f"No se encontró la plantilla múltiple en {self.ruta_plantilla_informe}")
            raise FileNotFoundError(f"Falta plantilla: {self.ruta_plantilla_informe}")
            
        log("Cargando plantilla de informe múltiple (docxtpl)...")
        from docxtpl import DocxTemplate
        import copy

        # Normalizar y expandir INSTITUCION_GLOBAL
        inst_g = super_contexto.get("INSTITUCION_GLOBAL", "")
        if not inst_g or re.search(r"^\s*Pontificia\s+Universidad(?:\s+Cat[oó]lica)?\s*$", inst_g, re.IGNORECASE) or ("Pontificia Universidad" in inst_g and "Católica del Perú" not in inst_g and "Catolica del Peru" not in inst_g):
            super_contexto["INSTITUCION_GLOBAL"] = "Pontificia Universidad Católica del Perú"
            
        beca_g = super_contexto.get("BECA_TITULO_GLOBAL", "")
        if not beca_g:
            super_contexto["BECA_TITULO_GLOBAL"] = "Beca 18 - Convocatoria 2023"

        # Expandir cada becario en super_contexto con aliases, fallbacks y variantes de casing
        becarios_limpios = []
        for b in super_contexto.get("becarios", []):
            b_exp = dict(b)
            # Asegurar institución completa
            inst_b = b_exp.get("INSTITUCION", "")
            if not inst_b or re.search(r"^\s*Pontificia\s+Universidad(?:\s+Cat[oó]lica)?\s*$", inst_b, re.IGNORECASE) or ("Pontificia Universidad" in inst_b and "Católica del Perú" not in inst_b and "Catolica del Peru" not in inst_b):
                b_exp["INSTITUCION"] = "Pontificia Universidad Católica del Perú"
            
            # Asegurar beca y convocatoria
            if not b_exp.get("BECA_Y_CONVOCATORIA_VALIDADA"):
                b_exp["BECA_Y_CONVOCATORIA_VALIDADA"] = super_contexto.get("BECA_TITULO_GLOBAL", "Beca 18 - Convocatoria 2023")
                
            # Fechas de inicio y fin (SIBEC / Periodo actual)
            f_ini = b_exp.get("FECHA_INICIO_SIBEC") or b_exp.get("FECHA_INICIO") or b_exp.get("INICIO") or "17/08/2026"
            f_fin = b_exp.get("FECHA_FIN_SIBEC") or b_exp.get("FECHA_FIN") or b_exp.get("FIN") or "15/12/2026"
            b_exp["FECHA_INICIO_SIBEC"] = f_ini
            b_exp["FECHA_INICIO"] = f_ini
            b_exp["INICIO"] = f_ini
            b_exp["PERIODO_INICIO"] = f_ini
            b_exp["FECHA_INICIO_ESTUDIOS"] = f_ini
            b_exp["FECHA_FIN_SIBEC"] = f_fin
            b_exp["FECHA_FIN"] = f_fin
            b_exp["FIN"] = f_fin
            b_exp["PERIODO_FIN"] = f_fin
            b_exp["FECHA_FIN_ESTUDIOS"] = f_fin

            b_exp["BECA_Y_CONVOCATORIA"] = b_exp["BECA_Y_CONVOCATORIA_VALIDADA"]
            b_exp["BECA"] = b_exp["BECA_Y_CONVOCATORIA_VALIDADA"]
            b_exp["PROGRAMA_BECA"] = b_exp["BECA_Y_CONVOCATORIA_VALIDADA"]
            
            carr = b_exp.get("CARRERA", "")
            b_exp["PROGRAMA_ESTUDIOS"] = carr
            b_exp["ESPECIALIDAD"] = carr
            b_exp["IES"] = b_exp["INSTITUCION"]
            b_exp["UNIVERSIDAD"] = b_exp["INSTITUCION"]

            # Multi-casing para cada becario
            for k, val in list(b_exp.items()):
                if isinstance(val, (str, int, float)):
                    val_str = str(val)
                    b_exp[k.lower()] = val_str
                    b_exp[k.upper()] = val_str
                    b_exp[k.title()] = val_str
            becarios_limpios.append(b_exp)

        super_contexto["becarios"] = becarios_limpios

        # Multi-casing a nivel de super_contexto
        for k, val in list(super_contexto.items()):
            if isinstance(val, (str, int, float)):
                val_str = str(val)
                super_contexto[k.lower()] = val_str
                super_contexto[k.upper()] = val_str
                super_contexto[k.title()] = val_str

        doc = DocxTemplate(str(self.ruta_plantilla_informe))
        doc.render(super_contexto)
        
        num_inf = super_contexto.get("NUMERO_INFORME_GENERAR", "S-N")
        nombre_salida = f"Informe_{num_inf}_Multiple.docx"
        ruta_salida = get_base_dir() / "Informes_Generados" / nombre_salida
        ruta_salida.parent.mkdir(parents=True, exist_ok=True)
        
        log(f"Guardando Informe Múltiple en: {ruta_salida.name}")
        doc.save(ruta_salida)

        # Post-procesamiento: Reemplazo profundo y formato de referencias
        from docx import Document
        from docx.shared import Inches, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        import re

        doc_final = Document(ruta_salida)
        modificado = False

        # Corrección global: Asegurar año en SIGEDO INTEGRADO (ej. SIGEDO INTEGRADO: 59288 -> 59288-2026)
        for p in doc_final.paragraphs:
            if "SIGEDO INTEGRADO" in p.text:
                if not re.search(r"SIGEDO\s+INTEGRADO\s*:\s*\d+-\d{4}", p.text):
                    p.text = re.sub(r'(SIGEDO\s+INTEGRADO\s*:\s*)(\d{4,8})\b', r'\g<1>\g<2>-2026', p.text)
                    modificado = True
        for t in doc_final.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if "SIGEDO INTEGRADO" in p.text:
                            if not re.search(r"SIGEDO\s+INTEGRADO\s*:\s*\d+-\d{4}", p.text):
                                p.text = re.sub(r'(SIGEDO\s+INTEGRADO\s*:\s*)(\d{4,8})\b', r'\g<1>\g<2>-2026', p.text)
                                modificado = True

        # Corrección global: Beca y Convocatoria real en ASUNTO y numeral 2.1
        beca_real = super_contexto.get("BECA_TITULO_GLOBAL", "")
        if beca_real:
            for p in doc_final.paragraphs:
                if "Beca 18 - Convocatoria 2021" in p.text:
                    p.text = p.text.replace("Beca 18 - Convocatoria 2021", beca_real)
                    modificado = True
            for t in doc_final.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if "Beca 18 - Convocatoria 2021" in p.text:
                                p.text = p.text.replace("Beca 18 - Convocatoria 2021", beca_real)
                                modificado = True

        # 0. Reemplazar número de informe en el título (INFORME Nº XXXX-2026...)
        # Requerimiento: Arial 11, todo en mayúsculas, negrita y subrayado
        num_inf = super_contexto.get("NUMERO_INFORME_GENERAR", "")
        patron_num_inf = re.compile(r'(INFORME\s+N[ºo\.]?\s*)\d*(-\d{4}-MINEDU/VMGI-PRONABEC-DIBEC-SUS)', re.IGNORECASE)
        for p in doc_final.paragraphs[:5]:
            if patron_num_inf.search(p.text) or "MINEDU/VMGI-PRONABEC-DIBEC-SUS" in p.text:
                if num_inf and num_inf != "S-N":
                    nuevo_txt = patron_num_inf.sub(rf"\g<1>{num_inf}\g<2>", p.text)
                else:
                    nuevo_txt = p.text
                nuevo_txt = nuevo_txt.upper()
                p.text = nuevo_txt
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(11)
                    r.bold = True
                    r.underline = True
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                modificado = True
                break

        # A. Actualizar nombre de la IES a Title Case respetando conectores en minúsculas ('de', 'del', 'y', etc.)
        inst_real = super_contexto.get("INSTITUCION_GLOBAL", "")
        if inst_real:
            inst_upper = inst_real.upper().strip()
            inst_title = formatear_nombre_ies(inst_real)
            inst_wrong = " ".join(w.capitalize() for w in inst_real.split())
            texto_ies_ant = "Universidad Peruana de Ciencias Aplicadas"
            texto_ies_ant_wrong = "Universidad Peruana De Ciencias Aplicadas"
            
            # A1. Párrafos narrativos (numeral 2.6, etc.)
            for p in doc_final.paragraphs:
                if inst_upper in p.text:
                    p.text = p.text.replace(inst_upper, inst_title)
                    modificado = True
                if inst_wrong in p.text:
                    p.text = p.text.replace(inst_wrong, inst_title)
                    modificado = True
                if texto_ies_ant in p.text:
                    p.text = p.text.replace(texto_ies_ant, inst_title)
                    modificado = True
                if texto_ies_ant_wrong in p.text:
                    p.text = p.text.replace(texto_ies_ant_wrong, inst_title)
                    modificado = True

            # A2. Tabla 0 (Encabezado: ASUNTO)
            if len(doc_final.tables) > 0:
                for row in doc_final.tables[0].rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if inst_upper in p.text:
                                p.text = p.text.replace(inst_upper, inst_title)
                                modificado = True
                            if inst_wrong in p.text:
                                p.text = p.text.replace(inst_wrong, inst_title)
                                modificado = True
                            if texto_ies_ant in p.text:
                                p.text = p.text.replace(texto_ies_ant, inst_title)
                                modificado = True
                            if texto_ies_ant_wrong in p.text:
                                p.text = p.text.replace(texto_ies_ant_wrong, inst_title)
                                modificado = True

        # Asegurar que nunca quede 'Pontificia Universidad' sin 'Católica del Perú' en ningún párrafo o tabla
        patron_pucp_incompleto = re.compile(r'\bPontificia\s+Universidad\b(?!\s+Cat[oó]lica\s+del\s+Per[uú])', re.IGNORECASE)
        for p in doc_final.paragraphs:
            if patron_pucp_incompleto.search(p.text):
                p.text = patron_pucp_incompleto.sub("Pontificia Universidad Católica del Perú", p.text)
                modificado = True
        for t in doc_final.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if patron_pucp_incompleto.search(p.text):
                            p.text = patron_pucp_incompleto.sub("Pontificia Universidad Católica del Perú", p.text)
                            modificado = True

        # B. Actualizar informe SUCCOR: cambiar a 'Informe N° ' en todas sus menciones
        succor_real = super_contexto.get("NOMBRE_INFORME_SUCCOR", "")
        if succor_real:
            succor_real = re.sub(
                r"^\s*(?:INFORME|Informe)(?:\s+(?:T[EÉ]CNICO|LEGAL|FINAL))?\s*(?:N[°ºo\.\s]*|NRO\.?|N[UÚ]MERO|NUMERO|N\.º|N\.°|N°:|N°\s*:)?\s*[:\s]*",
                "Informe N° ",
                succor_real,
                flags=re.IGNORECASE
            )
            patron_succor_gen = re.compile(
                r"(?:INFORME|Informe)(?:\s+(?:T[EÉ]CNICO|LEGAL|FINAL))?\s*(?:N[°ºo\.\s]*|NRO\.?|N[UÚ]MERO|NUMERO|N\.º|N\.°|N°:|N°\s*:)?\s*[\d]+[-\s\w/\.]*(?:SUCCOR[-\s\w]*LIMA|SUCCOR[-\s\w]*|DICONCI[-\s\w]*SUCCOR)",
                re.IGNORECASE
            )
            for p in doc_final.paragraphs:
                if patron_succor_gen.search(p.text):
                    p.text = patron_succor_gen.sub(succor_real, p.text)
                    modificado = True
                p_text_sub = re.sub(r"Mediante\s+el\s*(?:,\s*|\s+)la\s+Subdirecci[oó]n", f"Mediante el {succor_real}, la Subdirección", p.text, flags=re.IGNORECASE)
                if p_text_sub != p.text:
                    p.text = p_text_sub
                    modificado = True
                p_text_concl = re.sub(r"mediante\s*(?:,\s*|\s+)se\s+concluye", f"mediante {succor_real} se concluye", p.text, flags=re.IGNORECASE)
                if p_text_concl != p.text:
                    p.text = p_text_concl
                    modificado = True
            for t in doc_final.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if patron_succor_gen.search(p.text):
                                p.text = patron_succor_gen.sub(succor_real, p.text)
                                modificado = True

        # C. Reemplazo del documento IES en numeral 2.6 (obviando fecha en documentos múltiples)
        ref_ies_real = super_contexto.get("REFERENCIA_DOC_IES", "")
        if ref_ies_real:
            patron_26 = re.compile(r'Carta\s+N[º°o]\s*º(?:\s*,\s*de\s+fecha\s+[^,\.\n]+)?', re.IGNORECASE)
            for p in doc_final.paragraphs:
                if patron_26.search(p.text):
                    p.text = patron_26.sub(ref_ies_real, p.text)
                    modificado = True
                elif "Carta Nº º" in p.text or "Carta Nº º," in p.text:
                    p.text = p.text.replace("Carta Nº º,", f"{ref_ies_real},").replace("Carta Nº º", ref_ies_real)
                    modificado = True

        # D. Formato de REFERENCIAS (Tabla 0, Fila 3, Celda 2): Párrafos independientes con sangría francesa a Arial 11 pt
        # y adición del año '-2026' al Expediente SIGEDO
        referencias_lista = super_contexto.get("REFERENCIAS", [])
        if referencias_lista and len(doc_final.tables) > 0 and len(doc_final.tables[0].rows) > 3:
            row_ref = doc_final.tables[0].rows[3]
            for c_idx in (0, 1):
                for p in row_ref.cells[c_idx].paragraphs:
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(11)

            cell_ref = row_ref.cells[2]
            cell_ref.text = ""
            for i, ref_texto in enumerate(referencias_lista):
                ref_texto_limpio = re.sub(r'\b(?:INFORME|Informe)\s+(?:N[º°o\.]*|N\.o|No)\s*(\d+-\d{4}-MINEDU/VMGI-PRONABEC)', r'Informe N° \1', ref_texto, flags=re.IGNORECASE)
                # Asegurar año -2026 en SIGEDO
                ref_texto_limpio = re.sub(r'\(Expediente SIGEDO\s+(\d{4,8})\)(?!-2026)', r'(Expediente SIGEDO \1-2026)', ref_texto_limpio)
                p = cell_ref.paragraphs[0] if i == 0 else cell_ref.add_paragraph()
                p.text = ref_texto_limpio
                p.paragraph_format.left_indent = Inches(0.25)
                p.paragraph_format.first_line_indent = Inches(-0.25)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.0
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(11)
            modificado = True

        # E. Numeral 2.5: Reemplazo de fechas concatenadas ('con fecha' vs 'con fechas')
        fechas_sol_texto = super_contexto.get("FECHAS_SOLICITUD_TEXTO", "")
        if fechas_sol_texto:
            prefijo_f = "con fechas" if (" y " in fechas_sol_texto or "," in fechas_sol_texto) else "con fecha"
            patron_num25 = re.compile(r'(ingresadas por mesa de partes\s+)(?:con\s+fechas?|con\s+fecha:?)\s*(.*?)([\.,]\s*por lo tanto)', re.IGNORECASE)
            for p in doc_final.paragraphs:
                if "ingresadas por mesa de partes" in p.text and "plazo de presentación" in p.text:
                    if patron_num25.search(p.text):
                        p.text = patron_num25.sub(rf"\g<1>{prefijo_f} {fechas_sol_texto}\g<3>", p.text)
                        for r in p.runs:
                            r.font.name = "Arial"
                            r.font.size = Pt(11)
                        modificado = True

        # F. Numeral 3.1: Formatear a Arial 11 pt y asegurar 'Informe N° ' SUCCOR
        for p in doc_final.paragraphs:
            if "De conformidad a lo informado por la Subdirección" in p.text:
                if patron_inf_succor.search(p.text):
                    p.text = patron_inf_succor.sub(r"Informe N° \1", p.text)
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(11)
                modificado = True

        # G. Concordancia de Género Femenino (si todos los becarios son de sexo femenino)
        todos_femeninos = super_contexto.get("TODOS_FEMENINOS", False)
        if todos_femeninos:
            # 1. En Tabla 0 (ASUNTO)
            if len(doc_final.tables) > 0:
                for row in doc_final.tables[0].rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if "becarios" in p.text.lower():
                                p.text = re.sub(r'\b(por\s+\d+\s+)becarios\b', r'\g<1>becarias', p.text, flags=re.IGNORECASE)
                                p.text = re.sub(r'\bde\s+los\s+becarios\b', 'de las becarias', p.text, flags=re.IGNORECASE)
                                p.text = re.sub(r'\blos\s+(\d+\s+)?becarios\b', r'las \g<1>becarias', p.text, flags=re.IGNORECASE)
                                modificado = True

            # 2. En párrafos narrativos
            for p in doc_final.paragraphs:
                txt_ant = p.text
                if "Becarios que solicitan" in p.text:
                    p.text = p.text.replace("Becarios que solicitan", "Becarias que solicitan")
                elif "Becarios y relación de cursos" in p.text:
                    p.text = p.text.replace("Becarios y relación de cursos", "Becarias y relación de cursos")
                elif any(k in p.text for k in ["traslada la solicitud", "ingresadas por mesa de partes", "remite la relación de cursos", "solicitudes de los becarios", "De conformidad a lo informado", "revisión de las solicitudes"]):
                    p.text = re.sub(r'\b(de\s+\d+\s+)becarios\b', r'\g<1>becarias', p.text, flags=re.IGNORECASE)
                    p.text = re.sub(r'\b(Las solicitudes de\s+)los(\s+\d+\s+)becarios\b', r'\g<1>las\g<2>becarias', p.text, flags=re.IGNORECASE)
                    p.text = re.sub(r'\blos(\s+\d+\s+)becarios\s+señalados\b', r'las\g<1>becarias señaladas', p.text, flags=re.IGNORECASE)
                    p.text = re.sub(r'\blos\s+citados\s+becarios\b', 'las citadas becarias', p.text, flags=re.IGNORECASE)
                    p.text = re.sub(r'\b(solicitudes?\s+de\s+)los\s+becarios\b', r'\g<1>las becarias', p.text, flags=re.IGNORECASE)
                    p.text = re.sub(r'\blos(\s+\d+\s+)becarios\s+detallados\b', r'las\g<1>becarias detalladas', p.text, flags=re.IGNORECASE)
                    p.text = re.sub(r'\blos(\s+\d+\s+)becarios\s+cumplen\b', r'las\g<1>becarias cumplen', p.text, flags=re.IGNORECASE)
                if p.text != txt_ant:
                    modificado = True

        # H. Cuadro N° 2: Formatear Cursos Pendientes con numeración (1., 2., 3...) y sangría francesa
        if len(doc_final.tables) > 2:
            t2 = doc_final.tables[2]
            for r_idx in range(2, len(t2.rows)):
                cell_cursos = t2.rows[r_idx].cells[2]
                raw_txt = cell_cursos.text.strip()
                if raw_txt:
                    lineas_raw = [l.strip() for l in raw_txt.splitlines() if l.strip()]
                    cursos_limpios = []
                    for l in lineas_raw:
                        l_prot = re.sub(r'\bECOLOGY\s*,\s*ENVIRONMENT', 'ECOLOGY###COMMA###ENVIRONMENT', l, flags=re.IGNORECASE)
                        if ";" in l_prot:
                            subpartes = re.split(r";+", l_prot)
                        elif len(re.findall(r'\b\d+[\.\)]', l_prot)) > 1:
                            subpartes = re.split(r"(?<=\w)\s+(?=\d+[\.\)]\s+)", l_prot)
                        elif "," in l_prot and not re.match(r'^\d+[\.\)]\s+', l):
                            subpartes = re.split(r",\s*(?![^()]*\))", l_prot)
                        else:
                            subpartes = [l_prot]

                        for sp in subpartes:
                            sp_clean = sp.replace("###COMMA###", ", ").strip()
                            sp_clean = re.sub(r"^\s*\d+[\.\)]\s*", "", sp_clean).strip()
                            if sp_clean:
                                cursos_limpios.append(sp_clean)
                    
                    if cursos_limpios:
                        cell_cursos.text = ""
                        for c_i, c_nom in enumerate(cursos_limpios):
                            p_c = cell_cursos.paragraphs[0] if c_i == 0 else cell_cursos.add_paragraph()
                            p_c.text = f"{c_i+1}. {c_nom.upper()}"
                            p_c.paragraph_format.left_indent = Inches(0.22)
                            p_c.paragraph_format.first_line_indent = Inches(-0.22)
                            p_c.paragraph_format.space_after = Pt(2)
                            p_c.paragraph_format.space_before = Pt(0)
                            p_c.paragraph_format.line_spacing = 1.0
                            p_c.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            for r in p_c.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8.5)
                        modificado = True

        # I. Cuadro N° 1: Validación y llenado garantizado de datos
        if len(doc_final.tables) > 1:
            t1 = doc_final.tables[1]
            becarios_list = super_contexto.get("becarios", [])
            for r_idx in range(1, len(t1.rows)):
                row = t1.rows[r_idx]
                b_idx = r_idx - 1
                b_data = becarios_list[b_idx] if b_idx < len(becarios_list) else {}
                
                # Col 2: RJ Adjudicación
                if len(row.cells) > 2:
                    c2_txt = row.cells[2].text.strip()
                    if not c2_txt or "{{" in c2_txt:
                        rjd_v = b_data.get("RJD_ADJUDICACION") or super_contexto.get("RJD_ADJUDICACION_GLOBAL", "")
                        if rjd_v:
                            row.cells[2].text = rjd_v
                            modificado = True

                # Col 3: BECA Y CONVOCATORIA
                if len(row.cells) > 3:
                    c3_txt = row.cells[3].text.strip()
                    if not c3_txt or "{{" in c3_txt:
                        b_c_val = b_data.get("BECA_Y_CONVOCATORIA_VALIDADA") or super_contexto.get("BECA_TITULO_GLOBAL", "Beca 18 - Convocatoria 2023")
                        row.cells[3].text = b_c_val
                        for p in row.cells[3].paragraphs:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for r in p.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8)
                        modificado = True

                # Col 4: INSTITUCION
                if len(row.cells) > 4:
                    c4_txt = row.cells[4].text.strip()
                    if not c4_txt or "{{" in c4_txt or ("PONTIFICIA UNIVERSIDAD" in c4_txt.upper() and "CATÓLICA" not in c4_txt.upper() and "CATOLICA" not in c4_txt.upper()):
                        row.cells[4].text = "Pontificia Universidad Católica del Perú"
                        for p in row.cells[4].paragraphs:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for r in p.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8)
                        modificado = True

                # Col 5: CARRERA
                if len(row.cells) > 5:
                    c5_txt = row.cells[5].text.strip()
                    if not c5_txt or "{{" in c5_txt:
                        carr_v = b_data.get("CARRERA") or ""
                        if carr_v:
                            row.cells[5].text = carr_v
                            for p in row.cells[5].paragraphs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for r in p.runs:
                                    r.font.name = "Arial"
                                    r.font.size = Pt(8)
                            modificado = True

        # J. Cuadro N° 2: Validación y llenado de Periodo de Estudios (Inicio y Fin)
        if len(doc_final.tables) > 2:
            t2 = doc_final.tables[2]
            becarios_list = super_contexto.get("becarios", [])
            for r_idx in range(2, len(t2.rows)):
                row = t2.rows[r_idx]
                b_idx = r_idx - 2
                b_data = becarios_list[b_idx] if b_idx < len(becarios_list) else {}
                
                # Col 3: Periodo Inicio
                if len(row.cells) > 3:
                    c3_txt = row.cells[3].text.strip()
                    if not c3_txt or "{{" in c3_txt:
                        f_ini = b_data.get("FECHA_INICIO_SIBEC") or b_data.get("FECHA_INICIO") or b_data.get("INICIO") or "17/08/2026"
                        row.cells[3].text = f_ini
                        for p in row.cells[3].paragraphs:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for r in p.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8.5)
                        modificado = True
                        
                # Col 4: Periodo Fin
                if len(row.cells) > 4:
                    c4_txt = row.cells[4].text.strip()
                    if not c4_txt or "{{" in c4_txt:
                        f_fin = b_data.get("FECHA_FIN_SIBEC") or b_data.get("FECHA_FIN") or b_data.get("FIN") or "15/12/2026"
                        row.cells[4].text = f_fin
                        for p in row.cells[4].paragraphs:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for r in p.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8.5)
                        modificado = True

        if modificado:
            doc_final.save(ruta_salida)
            log("Post-procesamiento de formato y variables aplicado en documento final.")

        return ruta_salida
