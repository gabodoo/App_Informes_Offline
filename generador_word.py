from __future__ import annotations

import sys
import re
import shutil
from datetime import datetime, date
from pathlib import Path
from typing import Callable, Any

import copy
import docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
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


def obtener_ruta_plantilla_multiple() -> Path:
    """Busca y retorna la ruta existente de plantilla_informe_multiple.docx probando múltiples ubicaciones."""
    rutas_candidatas = [
        BASE_DIR / "plantillas" / "plantilla_informe_multiple.docx",
        BASE_DIR / "dist" / "main" / "plantillas" / "plantilla_informe_multiple.docx",
        BASE_DIR / "dist" / "plantillas" / "plantilla_informe_multiple.docx",
        Path(__file__).resolve().parent / "plantillas" / "plantilla_informe_multiple.docx",
    ]
    for ruta in rutas_candidatas:
        if ruta.is_file():
            return ruta
    return BASE_DIR / "plantillas" / "plantilla_informe_multiple.docx"


PLANTILLA_PATH = obtener_ruta_plantilla()
PLANTILLA_MULTIPLE_PATH = obtener_ruta_plantilla_multiple()
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
    if "Pontificia Universidad" in nombre_clean and ("Católica" in nombre_clean or "Catolica" in nombre_clean):
        return "Pontificia Universidad Católica del Perú"
    if "Pontificia Universidad" in nombre_clean:
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


def limpiar_nombre_ies_sin_sede(nombre: str) -> str:
    """Elimina sufijos de sede, filial, campus u otra información que no sea el nombre oficial de la IES."""
    if not nombre:
        return ""
    n = str(nombre).strip()
    n = re.sub(r"[\s\.]*[/\-–—]\s*(?:Sede|Filial|Campus|Local|Sucursal)\b.*$", "", n, flags=re.IGNORECASE)
    n = re.sub(r"\s*,\s*(?:Sede|Filial|Campus|Local|Sucursal)\b.*$", "", n, flags=re.IGNORECASE)
    n = re.sub(r"\s+\b(?:Sede|Filial|Campus|Local)\s+.*$", "", n, flags=re.IGNORECASE)
    n = re.sub(r"[\s\/\-–—\.]+$", "", n).strip()
    return formatear_nombre_ies(n)


def corregir_nombre_beca_excelencia(texto: str) -> str:
    """Renombra 'Beca Excelencia' por 'Beca de Excelencia Académica para Hijos de Docentes'."""
    if not texto:
        return ""
    s = str(texto)
    # Mayúsculas completas
    s = re.sub(
        r'\bBECA\s+(?:DE\s+)?EXCELENCIA\b(?!\s+ACAD[EÉ]MICA)',
        'BECA DE EXCELENCIA ACADÉMICA PARA HIJOS DE DOCENTES',
        s
    )
    # Title Case o mixto
    s = re.sub(
        r'\b[Bb]eca\s+(?:de\s+)?[Ee]xcelencia\b(?!\s+[Aa]cad[eé]mica)',
        'Beca de Excelencia Académica para Hijos de Docentes',
        s
    )
    return s


def parsear_fecha_solicitud(fecha_val: Any) -> date | None:
    """Convierte un valor de fecha (date, datetime, string) a un objeto date."""
    if not fecha_val:
        return None
    if isinstance(fecha_val, datetime):
        return fecha_val.date()
    if isinstance(fecha_val, date):
        return fecha_val
    s = str(fecha_val).strip()
    if not s or s.lower() in ("(no detectada)", "none", ""):
        return None
    m_num = re.search(r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b', s)
    if m_num:
        try:
            return date(int(m_num.group(3)), int(m_num.group(2)), int(m_num.group(1)))
        except Exception:
            pass
    m_txt = re.search(r'\b(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚ]+)\s+de(?:l)?\s+(\d{4})\b', s)
    if m_txt:
        meses = {
            'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
            'julio': 7, 'agosto': 8, 'setiembre': 9, 'septiembre': 9, 'octubre': 10,
            'noviembre': 11, 'diciembre': 12
        }
        mes = meses.get(m_txt.group(2).lower(), 0)
        if mes > 0:
            try:
                return date(int(m_txt.group(3)), mes, int(m_txt.group(1)))
            except Exception:
                pass
    return None


def es_fecha_hasta_24_julio(fecha_val: Any) -> bool:
    """Determina si la fecha de solicitud es igual o menor al 24 de julio de 2026."""
    d = parsear_fecha_solicitud(fecha_val)
    if d:
        return d <= date(2026, 7, 24)
    return False


def analizar_fechas_solicitud_multiple(lista_fechas: list) -> dict:
    """Evalúa las fechas de los formatos autogenerados en un informe múltiple."""
    limite = date(2026, 7, 24)
    fechas_parsed = []
    for f in lista_fechas:
        d = parsear_fecha_solicitud(f)
        if d:
            fechas_parsed.append(d)
        elif isinstance(f, str) and f.strip():
            for m in re.finditer(r'(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚ]+)\s+de(?:l)?\s+(\d{4})', f):
                d_sub = parsear_fecha_solicitud(m.group(0))
                if d_sub:
                    fechas_parsed.append(d_sub)
    if not fechas_parsed:
        return {"hay_antes": False, "hay_despues": False, "coexisten": False, "todas_antes": False}
    hay_antes = any(d <= limite for d in fechas_parsed)
    hay_despues = any(d > limite for d in fechas_parsed)
    todas_antes = all(d <= limite for d in fechas_parsed)
    return {
        "hay_antes": hay_antes,
        "hay_despues": hay_despues,
        "coexisten": (hay_antes and hay_despues),
        "todas_antes": todas_antes
    }


TEXTO_NUMERAL_15_067_2025 = (
    "1.5.\tResolución Directoral Ejecutiva N° 067-2025-MINEDU-VMGI-PRONABEC, y modificatorias "
    "que aprueba la norma técnica denominada “Disposiciones para el becario del PRONABEC”."
)
TEXTO_NUMERAL_15_067_2025_SIN_NUM = (
    "Resolución Directoral Ejecutiva N° 067-2025-MINEDU-VMGI-PRONABEC, y modificatorias "
    "que aprueba la norma técnica denominada “Disposiciones para el becario del PRONABEC”."
)

TEXTO_NUMERAL_23_067_2025_ENCABEZADO = (
    "2.3.\tAsimismo, el artículo 2 de la norma técnica denominada “Disposiciones para el becario del PRONABEC”, "
    "aprobada mediante Resolución Directoral Ejecutiva N° 067-2025-MINEDU-VMGI-PRONABEC y modificatorias, "
    "señala en relación al financiamiento de la ampliación excepcional dispuesto en el literal a) del artículo 30 "
    "del Reglamento, lo siguiente:"
)
TEXTO_NUMERAL_23_067_2025_ENCABEZADO_SIN_NUM = (
    "Asimismo, el artículo 2 de la norma técnica denominada “Disposiciones para el becario del PRONABEC”, "
    "aprobada mediante Resolución Directoral Ejecutiva N° 067-2025-MINEDU-VMGI-PRONABEC y modificatorias, "
    "señala en relación al financiamiento de la ampliación excepcional dispuesto en el literal a) del artículo 30 "
    "del Reglamento, lo siguiente:"
)

PARRAFOS_CITA_ARTICULO_2_067_2025 = [
    ("(…)", False, 0.5, 0.0),
    ("Artículo 2.- Ampliación excepcional de la duración del periodo de estudios", True, 0.5, 0.0),
    ("(…)", False, 0.5, 0.0),
    ("El financiamiento de la ampliación excepcional está sujeto al cumplimiento de las siguientes condiciones:", False, 0.5, 0.0),
    (
        "a)\tEl becario presenta la solicitud de ampliación excepcional de un (01) periodo de estudios, por mesa de partes "
        "del PRONABEC al culminar su último periodo programado en el SIBEC y antes de la matrícula e inicio del periodo "
        "académico a ampliar, adjuntando la documentación que acredite que aún no ha culminado la carrera o programa de estudios "
        "y que precise los cursos pendientes que tiene el becario. La ampliación excepcional se otorga por un (01) periodo "
        "académico adicional, está sujeta a la disponibilidad presupuestal, se debe efectuar en el periodo académico inmediato "
        "posterior al término del periodo de estudios programado y debe ser suficiente para culminar la carrera o programa de estudios.",
        False, 0.75, -0.25
    ),
    (
        "b)\tLa OAGD traslada la solicitud a la SUCCOR o la que haga sus veces, quien se encarga de verificar la información "
        "presentada por el becario y remite, mediante informe, el expediente a la DIBEC.",
        False, 0.75, -0.25
    ),
    (
        "c)\tLa DIBEC evalúa cada caso y emite opinión, mediante informe, y con oficio emite pronunciamiento al becario, "
        "con copia a la SUCCOR o la que haga sus veces, quien tramita su notificación ante la OAGD. La opinión favorable "
        "es también comunicada a la OAF para los fines correspondientes.",
        False, 0.75, -0.25
    ),
    (
        "d)\tLa SUS de la DIBEC actualiza la fecha de duración de la beca en los casos que otorgue opinión favorable a la "
        "solicitud del becario. Asimismo, actualiza el estado y condición del becario en el SIBEC.",
        False, 0.75, -0.25
    ),
    ("(…)”.", False, 0.5, 0.0),
]

TEXTO_NUMERAL_25_INDIVIDUAL_067_2025 = (
    "2.5.\tLa norma técnica denominada “Disposiciones para el becario del PRONABEC”, establece para el otorgamiento "
    "de la ampliación excepcional de la duración del periodo de estudios, lo siguiente: “El becario presenta la solicitud "
    "de ampliación excepcional de un (01) periodo de estudios, por mesa de partes del PRONABEC al culminar su último periodo "
    "programado en el SIBEC y antes de la matrícula e inicio del periodo académico a ampliar, adjuntando la documentación que "
    "acredite que aún no ha culminado la carrera o programa de estudios y que precise los cursos pendientes que tiene el becario. "
    "La ampliación excepcional se otorga por un (01) periodo académico adicional, está sujeta a la disponibilidad presupuestal, "
    "se debe efectuar en el periodo académico inmediato posterior al término del periodo de estudios programado y debe ser "
    "suficiente para culminar la carrera o programa de estudios”."
)
TEXTO_NUMERAL_25_INDIVIDUAL_067_2025_SIN_NUM = (
    "La norma técnica denominada “Disposiciones para el becario del PRONABEC”, establece para el otorgamiento "
    "de la ampliación excepcional de la duración del periodo de estudios, lo siguiente: “El becario presenta la solicitud "
    "de ampliación excepcional de un (01) periodo de estudios, por mesa de partes del PRONABEC al culminar su último periodo "
    "programado en el SIBEC y antes de la matrícula e inicio del periodo académico a ampliar, adjuntando la documentación que "
    "acredite que aún no ha culminado la carrera o programa de estudios y que precise los cursos pendientes que tiene el becario. "
    "La ampliación excepcional se otorga por un (01) periodo académico adicional, está sujeta a la disponibilidad presupuestal, "
    "se debe efectuar en el periodo académico inmediato posterior al término del periodo de estudios programado y debe ser "
    "suficiente para culminar la carrera o programa de estudios”."
)

TEXTO_NUMERAL_26_MULTIPLE_COMBINADO = (
    "2.6.\tLa norma técnica denominada “Disposiciones para el becario del PRONABEC” aprobadas por la "
    "Resolución Directoral Ejecutiva N° 067-2025-MINEDU-VMGI-PRONABEC y modificatorias, y "
    "Resolución Directoral Ejecutiva N° 119-2026-MINEDU-VMGI-PRONABEC, establecen para el otorgamiento "
    "de la ampliación excepcional de la duración del periodo, requisitos y condiciones similares."
)
TEXTO_NUMERAL_26_MULTIPLE_COMBINADO_SIN_NUM = (
    "La norma técnica denominada “Disposiciones para el becario del PRONABEC” aprobadas por la "
    "Resolución Directoral Ejecutiva N° 067-2025-MINEDU-VMGI-PRONABEC y modificatorias, y "
    "Resolución Directoral Ejecutiva N° 119-2026-MINEDU-VMGI-PRONABEC, establecen para el otorgamiento "
    "de la ampliación excepcional de la duración del periodo, requisitos y condiciones similares."
)


def tiene_num_pr(p) -> bool:
    """Verifica si el párrafo tiene numeración automática de Word (w:numPr)."""
    p_pPr = p._p.find(qn("w:pPr"))
    return p_pPr is not None and p_pPr.find(qn("w:numPr")) is not None


def aplicar_texto_numeral(p, texto_con_numeral: str, texto_sin_numeral: str, negrita: bool = False):
    """Establece el texto en un párrafo respetando si su numeración es automática de Word o tipeada."""
    if tiene_num_pr(p):
        texto_final = texto_sin_numeral
    else:
        texto_final = texto_con_numeral

    p.text = texto_final
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.25)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for r in p.runs:
        r.font.name = "Arial"
        r.font.size = Pt(11)
        r.bold = negrita


def es_parrafo_15(p) -> bool:
    """Detecta de forma ultra-robusta el párrafo 1.5 en la sección de antecedentes."""
    txt = p.text.strip()
    if re.search(r"^\s*1\.5[\.\t\s]", txt):
        return True
    txt_l = txt.lower()
    if ("119-2026" in txt or "067-2025" in txt) and (
        "disposiciones para el becario" in txt_l or "norma técnica" in txt_l or "norma tecnica" in txt_l
    ):
        return True
    return False


def es_parrafo_23(p) -> bool:
    """Detecta de forma ultra-robusta el encabezado del numeral 2.3 en análisis."""
    txt = p.text.strip()
    if re.search(r"^\s*2\.3[\.\t\s]", txt):
        return True
    txt_l = txt.lower()
    if ("artículo 2" in txt_l or "articulo 2" in txt_l) and "disposiciones para el becario" in txt_l:
        if "119-2026" in txt or "067-2025" in txt or "artículo 30" in txt_l or "articulo 30" in txt_l or "financiamiento" in txt_l:
            return True
    return False


def es_parrafo_24(p) -> bool:
    """Detecta el numeral 2.4 (cierre de la cita de las condiciones de ampliación)."""
    txt = p.text.strip()
    if re.search(r"^\s*2\.4[\.\t\s]", txt):
        return True
    txt_l = txt.lower()
    if "de lo expuesto, debemos indicar" in txt_l and "ampliación excepcional" in txt_l:
        return True
    return False


def es_parrafo_25(p) -> bool:
    """Detecta de forma ultra-robusta el numeral 2.5 (análisis de la solicitud - norma técnica)."""
    txt = p.text.strip()
    if re.search(r"^\s*2\.5[\.\t\s]", txt):
        return True
    txt_l = txt.lower()
    if "disposiciones para el becario" in txt_l:
        if "establece para el otorgamiento" in txt_l or "establecen para el otorgamiento" in txt_l:
            return True
        if ("versión 02" in txt_l or "version 02" in txt_l) and ("cursos pendientes" in txt_l or "ampliación excepcional" in txt_l):
            return True
        if "requisitos y condiciones similares" in txt_l:
            return True
    return False


def renumera_parrafo(p, num_viejo: int, num_nuevo: int, prefijo: str = "2.") -> bool:
    """Renombra el numeral de un párrafo con número tipeado (ej: '2.5.\t' -> '2.6.\t')."""
    patron = re.compile(rf"^(\s*{re.escape(prefijo)}){num_viejo}([\.\t\s])")
    if not patron.search(p.text):
        return False

    if p.runs and patron.search(p.runs[0].text):
        p.runs[0].text = patron.sub(rf"\g<1>{num_nuevo}\g<2>", p.runs[0].text, count=1)
        return True

    nuevo_texto = patron.sub(rf"\g<1>{num_nuevo}\g<2>", p.text, count=1)
    r0_name = p.runs[0].font.name if p.runs and p.runs[0].font.name else "Arial"
    r0_size = p.runs[0].font.size if p.runs and p.runs[0].font.size else Pt(11)
    r0_bold = p.runs[0].bold if p.runs else False
    p.text = nuevo_texto
    if p.runs:
        p.runs[0].font.name = r0_name
        p.runs[0].font.size = r0_size
        p.runs[0].bold = r0_bold
    return True


def ajustar_normativa_informe_individual(doc: Document, contexto: dict, log: LogCallback | None = None) -> bool:
    """Si la fecha del formato autogenerado es <= 24 de julio de 2026:
    - Numeral 1.5: se reemplaza por la RDE N° 067-2025.
    - Numeral 2.3: se reemplaza por la RDE N° 067-2025 (artículo 2, condiciones y literales a, b, c, d).
    - Numeral 2.5: se reemplaza por el texto de la RDE N° 067-2025 sobre ampliación excepcional."""
    _log = log or (lambda msg: None)

    # 1. Probar todas las posibles fuentes de la fecha en el contexto
    candidatos_fecha = [
        contexto.get("FECHA_SOLICITUD_OBJ"),
        contexto.get("FECHA_SOLICITUD_TEXTO"),
        contexto.get("FECHA_SOLICITUD"),
        contexto.get("FECHA_SOL"),
        contexto.get("REFERENCIA_A"),
        contexto.get("SOLICITUD_MESA_PARTES"),
    ]

    fecha_detectada = None
    for cand in candidatos_fecha:
        if cand and es_fecha_hasta_24_julio(cand):
            fecha_detectada = cand
            break

    # 2. Si no se encontró en el contexto, escanear párrafos del documento (ej. en REFERENCIA o numeral 2.6)
    if not fecha_detectada:
        for p in doc.paragraphs:
            txt = p.text
            if "ingresada por mesa de partes el" in txt.lower() or "mesa de partes" in txt.lower():
                if es_fecha_hasta_24_julio(txt):
                    fecha_detectada = txt
                    break

    if not fecha_detectada:
        return False

    _log(f"  [Normativa Individual] Solicitud con fecha <= 24/07/2026 ({fecha_detectada}). Aplicando RDE N° 067-2025 en 1.5, 2.3 y 2.5...")
    modificado = False

    # 1. Reemplazo del numeral 1.5
    for p in doc.paragraphs:
        txt_u = p.text.strip().upper()
        if txt_u.startswith("II.") or "ANÁLISIS" in txt_u or "ANALISIS" in txt_u:
            break
        if es_parrafo_15(p):
            aplicar_texto_numeral(
                p,
                texto_con_numeral=TEXTO_NUMERAL_15_067_2025,
                texto_sin_numeral=TEXTO_NUMERAL_15_067_2025_SIN_NUM,
                negrita=False,
            )
            modificado = True
            break

    # 2. Reemplazo del numeral 2.3 y sus párrafos de cita
    idx_23 = None
    idx_24 = None
    for i, p in enumerate(doc.paragraphs):
        if idx_23 is None and es_parrafo_23(p):
            idx_23 = i
        elif idx_23 is not None and idx_24 is None and es_parrafo_24(p):
            idx_24 = i
            break

    # Fallback para idx_24 si no se encontró 2.4 exacto
    if idx_23 is not None and idx_24 is None:
        for i in range(idx_23 + 1, len(doc.paragraphs)):
            p_cand = doc.paragraphs[i]
            t_cand = p_cand.text.strip()
            if not t_cand:
                continue
            t_low = t_cand.lower()
            if t_low.startswith("de lo expuesto") or "análisis de la solicitud" in t_low or "analisis de la solicitud" in t_low:
                idx_24 = i
                break
            if re.match(r"^\s*2\.\d+", t_cand) and not any(t_cand.startswith(x) for x in ("a)", "b)", "c)", "d)")):
                idx_24 = i
                break

    if idx_23 is not None:
        p_23 = doc.paragraphs[idx_23]
        aplicar_texto_numeral(
            p_23,
            texto_con_numeral=TEXTO_NUMERAL_23_067_2025_ENCABEZADO,
            texto_sin_numeral=TEXTO_NUMERAL_23_067_2025_ENCABEZADO_SIN_NUM,
            negrita=False,
        )

        if idx_24 is not None:
            p_24 = doc.paragraphs[idx_24]
            parrafos_eliminar = doc.paragraphs[idx_23 + 1:idx_24]
            for p_del in parrafos_eliminar:
                try:
                    p_del._p.getparent().remove(p_del._p)
                except Exception:
                    pass

            for q_txt, es_b, ind_l, ind_f in PARRAFOS_CITA_ARTICULO_2_067_2025:
                p_nuevo = p_24.insert_paragraph_before(q_txt)
                pPr = p_nuevo._p.find(qn("w:pPr"))
                if pPr is not None:
                    numPr = pPr.find(qn("w:numPr"))
                    if numPr is not None:
                        pPr.remove(numPr)
                if ind_l:
                    p_nuevo.paragraph_format.left_indent = Inches(ind_l)
                if ind_f:
                    p_nuevo.paragraph_format.first_line_indent = Inches(ind_f)
                p_nuevo.paragraph_format.space_after = Pt(2)
                p_nuevo.paragraph_format.line_spacing = 1.0
                p_nuevo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if any(q_txt.startswith(x) for x in ("a)", "b)", "c)", "d)")) else WD_ALIGN_PARAGRAPH.LEFT
                for r in p_nuevo.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(11)
                    r.bold = es_b
        modificado = True

    # 3. Reemplazo del numeral 2.5
    for p in doc.paragraphs:
        if es_parrafo_25(p):
            aplicar_texto_numeral(
                p,
                texto_con_numeral=TEXTO_NUMERAL_25_INDIVIDUAL_067_2025,
                texto_sin_numeral=TEXTO_NUMERAL_25_INDIVIDUAL_067_2025_SIN_NUM,
                negrita=False,
            )
            modificado = True
            break

    return modificado


def ajustar_normativa_informe_multiple(doc: Document, super_contexto: dict, log: LogCallback | None = None) -> bool:
    """Ajusta la normativa en informes múltiples según las fechas de las solicitudes:
    Caso A: Coexisten fechas <= 24/07/2026 y >= 25/07/2026:
      - Adiciona RDE 067-2025 antes de 1.5 (como nuevo 1.5), corriendo 1.5+ a 1.6+ en sección 1.
      - Adiciona la cita de RDE 067-2025 después de 2.2 como numeral 2.3.
      - Corre numeral 2.3 existente a 2.4, y 2.4 a 2.5.
      - El numeral que continúa tras 'Análisis de la solicitud' (renumerado como 2.6) reemplaza su contenido
        indicando que ambas normas (067-2025 y 119-2026) establecen condiciones similares.
      - Corre todos los numerales subsiguientes en sección 2 (+1) para evitar duplicidades.
    Caso B: Todas las fechas son <= 24/07/2026:
      - Aplica reemplazo 1-a-1 de 1.5, 2.3 y 2.5 con RDE 067-2025.
    Caso C: Todas las fechas son >= 25/07/2026:
      - Mantiene la plantilla intacta (RDE 119-2026)."""
    _log = log or (lambda msg: None)

    fechas = super_contexto.get("FECHAS_SOLICITUD_TODAS_OBJS", [])
    if not fechas and "becarios" in super_contexto:
        fechas = [b.get("FECHA_SOLICITUD_OBJ") or b.get("FECHA_SOLICITUD_TEXTO") or b.get("FECHA_SOLICITUD") for b in super_contexto["becarios"]]
    if not fechas:
        fechas = [super_contexto.get("FECHAS_SOLICITUD_TEXTO", "")]

    analisis_f = analizar_fechas_solicitud_multiple(fechas)
    if not analisis_f["hay_antes"]:
        return False

    # CASO B: Todas las solicitudes son <= 24/07/2026 (sin coexistencia)
    if analisis_f["todas_antes"] and not analisis_f["hay_despues"]:
        _log("  [Normativa Múltiple] Todas las solicitudes <= 24/07/2026. Aplicando RDE N° 067-2025.")
        return ajustar_normativa_informe_individual(doc, super_contexto, log=_log)

    # CASO A: Coexistencia (solicitudes <= 24/07/2026 Y solicitudes >= 25/07/2026)
    _log("  [Normativa Múltiple] Coexistencia de fechas detectada (<= 24/07 y >= 25/07). Aplicando adición y corrimiento de numerales...")
    modificado = False

    # -------------------------------------------------------------
    # PASO 1: SECCIÓN 1 (BASE LEGAL)
    # Adicionar 067-2025 antes del numeral 1.5 y correr 1.5+ a 1.6+
    # -------------------------------------------------------------
    p_15_orig = None
    idx_15_orig = None
    for i, p in enumerate(doc.paragraphs):
        txt_u = p.text.strip().upper()
        if txt_u.startswith("II.") or "ANÁLISIS" in txt_u or "ANALISIS" in txt_u:
            break
        if es_parrafo_15(p):
            p_15_orig = p
            idx_15_orig = i
            break

    if p_15_orig is not None:
        ya_insertado_15 = False
        if idx_15_orig > 0 and "067-2025" in doc.paragraphs[idx_15_orig - 1].text:
            ya_insertado_15 = True

        if not ya_insertado_15:
            fin_seccion_1 = len(doc.paragraphs)
            for j in range(idx_15_orig, len(doc.paragraphs)):
                txt_j = doc.paragraphs[j].text.strip().upper()
                if txt_j.startswith("II.") or "ANÁLISIS" in txt_j or "ANALISIS" in txt_j or re.match(r"^\s*2\.\d", txt_j):
                    fin_seccion_1 = j
                    break

            ref_pPr = p_15_orig._p.find(qn("w:pPr"))
            ref_numPr = ref_pPr.find(qn("w:numPr")) if ref_pPr is not None else None
            tiene_num = ref_numPr is not None

            parrafos_sec1_correr = []
            for j in range(idx_15_orig, fin_seccion_1):
                p_item = doc.paragraphs[j]
                m_num = re.match(r"^\s*1\.(\d+)[\.\t\s]", p_item.text)
                if m_num:
                    num_val = int(m_num.group(1))
                    if num_val >= 5:
                        parrafos_sec1_correr.append((p_item, num_val))

            for p_item, num_val in sorted(parrafos_sec1_correr, key=lambda x: x[1], reverse=True):
                renumera_parrafo(p_item, num_val, num_val + 1, prefijo="1.")

            p_nuevo_15 = p_15_orig.insert_paragraph_before()
            if tiene_num:
                p_nuevo_pPr = p_nuevo_15._p.get_or_add_pPr()
                p_nuevo_pPr.append(copy.deepcopy(ref_numPr))
                p_nuevo_15.text = TEXTO_NUMERAL_15_067_2025_SIN_NUM
            else:
                p_nuevo_15.text = TEXTO_NUMERAL_15_067_2025

            p_nuevo_15.paragraph_format.left_indent = Inches(0.25)
            p_nuevo_15.paragraph_format.first_line_indent = Inches(-0.25)
            p_nuevo_15.paragraph_format.space_after = Pt(2)
            p_nuevo_15.paragraph_format.line_spacing = 1.0
            p_nuevo_15.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p_nuevo_15.runs:
                r.font.name = "Arial"
                r.font.size = Pt(11)
            modificado = True

    # -------------------------------------------------------------
    # PASO 2: SECCIÓN 2 (ANÁLISIS)
    # Adicionar 067-2025 después de 2.2 como nuevo 2.3,
    # correr 2.3 -> 2.4, 2.4 -> 2.5,
    # reemplazar numeral tras 'Análisis de la solicitud' por 2.6 (ambas normas),
    # y correr 2.6 -> 2.7, 2.7 -> 2.8, etc.
    # -------------------------------------------------------------
    p_23_orig = None
    p_24_orig = None
    p_25_orig = None
    idx_sub_analisis = None
    idx_fin_seccion_2 = len(doc.paragraphs)

    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        txt_u = txt.upper()
        if p_23_orig is None and es_parrafo_23(p):
            p_23_orig = p
        elif p_24_orig is None and es_parrafo_24(p):
            p_24_orig = p
        elif idx_sub_analisis is None and ("ANÁLISIS DE LA" in txt_u or "ANALISIS DE LA" in txt_u or "ANÁLISIS DE LAS" in txt_u or "ANALISIS DE LAS" in txt_u):
            idx_sub_analisis = i
        elif idx_sub_analisis is not None and p_25_orig is None and es_parrafo_25(p):
            p_25_orig = p
        elif txt_u.startswith("III.") or "CONCLUSI" in txt_u or re.match(r"^\s*3\.\d", txt_u):
            if idx_fin_seccion_2 == len(doc.paragraphs):
                idx_fin_seccion_2 = i

    if p_25_orig is None:
        for i, p in enumerate(doc.paragraphs):
            if p_23_orig is not None and i > doc.paragraphs.index(p_23_orig):
                if es_parrafo_25(p):
                    p_25_orig = p
                    break

    if p_23_orig is not None:
        ya_insertado_23 = False
        idx_cur_23 = doc.paragraphs.index(p_23_orig)
        if idx_cur_23 > 0 and "067-2025" in doc.paragraphs[idx_cur_23 - 1].text:
            ya_insertado_23 = True

        if not ya_insertado_23:
            parrafos_despues_25 = []
            inicio_busqueda = doc.paragraphs.index(p_25_orig) if p_25_orig else (idx_sub_analisis or doc.paragraphs.index(p_23_orig))
            for j in range(inicio_busqueda + 1, idx_fin_seccion_2):
                p_item = doc.paragraphs[j]
                m_num = re.match(r"^\s*2\.(\d+)[\.\t\s]", p_item.text)
                if m_num:
                    num_val = int(m_num.group(1))
                    if num_val >= 6:
                        parrafos_despues_25.append((p_item, num_val))

            for p_item, num_val in sorted(parrafos_despues_25, key=lambda x: x[1], reverse=True):
                renumera_parrafo(p_item, num_val, num_val + 1, prefijo="2.")

            if p_25_orig is not None:
                aplicar_texto_numeral(
                    p_25_orig,
                    texto_con_numeral=TEXTO_NUMERAL_26_MULTIPLE_COMBINADO,
                    texto_sin_numeral=TEXTO_NUMERAL_26_MULTIPLE_COMBINADO_SIN_NUM,
                    negrita=False,
                )

            if p_24_orig is not None:
                renumera_parrafo(p_24_orig, 4, 5, prefijo="2.")

            renumera_parrafo(p_23_orig, 3, 4, prefijo="2.")

            ref_pPr_23 = p_23_orig._p.find(qn("w:pPr"))
            ref_numPr_23 = ref_pPr_23.find(qn("w:numPr")) if ref_pPr_23 is not None else None
            tiene_num_23 = ref_numPr_23 is not None

            p_nuevo_23 = p_23_orig.insert_paragraph_before()
            if tiene_num_23:
                p_nuevo_23_pPr = p_nuevo_23._p.get_or_add_pPr()
                p_nuevo_23_pPr.append(copy.deepcopy(ref_numPr_23))
                p_nuevo_23.text = TEXTO_NUMERAL_23_067_2025_ENCABEZADO_SIN_NUM
            else:
                p_nuevo_23.text = TEXTO_NUMERAL_23_067_2025_ENCABEZADO

            p_nuevo_23.paragraph_format.left_indent = Inches(0.25)
            p_nuevo_23.paragraph_format.first_line_indent = Inches(-0.25)
            p_nuevo_23.paragraph_format.space_after = Pt(2)
            p_nuevo_23.paragraph_format.line_spacing = 1.0
            p_nuevo_23.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p_nuevo_23.runs:
                r.font.name = "Arial"
                r.font.size = Pt(11)

            for q_txt, es_b, ind_l, ind_f in PARRAFOS_CITA_ARTICULO_2_067_2025:
                p_q = p_23_orig.insert_paragraph_before(q_txt)
                pPr = p_q._p.find(qn("w:pPr"))
                if pPr is not None:
                    numPr = pPr.find(qn("w:numPr"))
                    if numPr is not None:
                        pPr.remove(numPr)
                if ind_l:
                    p_q.paragraph_format.left_indent = Inches(ind_l)
                if ind_f:
                    p_q.paragraph_format.first_line_indent = Inches(ind_f)
                p_q.paragraph_format.space_after = Pt(2)
                p_q.paragraph_format.line_spacing = 1.0
                p_q.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if any(q_txt.startswith(x) for x in ("a)", "b)", "c)", "d)")) else WD_ALIGN_PARAGRAPH.LEFT
                for r in p_q.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(11)
                    r.bold = es_b

            modificado = True

    return modificado


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
            "EXPEDIENTE_BECARIO": [
                "EXPEDIENTE", "NEXPEDIENTE", "NUMERO_EXPEDIENTE", "N_EXPEDIENTE",
                "NRO_EXPEDIENTE", "NUM_EXPEDIENTE", "EXPEDIENTE_PADRON"
            ],
            "NUMERO_SIGEDO": [
                "NUM_SIGEDO", "SIGEDO", "EXPEDIENTE_SIGEDO", "SIGEDO_CORTO"
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
        self._reemplazo_profundo_docx(ruta_salida, mapa_reemplazos, contexto, log=_log)

        # PASO 3: Expansión especial de CURSOS_PENDIENTES en párrafos numerados individuales
        cursos_lista = contexto.get("CURSOS_PENDIENTES", "")
        if not cursos_lista:
            cursos_lista = "(No se detectaron cursos pendientes)"
        self._expandir_cursos_pendientes(ruta_salida, cursos_lista)

        _log("  Documento Word guardado correctamente.")
        return ruta_salida

    def _reemplazo_profundo_docx(self, ruta_docx: Path, mapa_reemplazos: dict[str, str], contexto: dict, log: LogCallback | None = None) -> None:
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
        val_fsol = str(contexto.get("FECHA_SOLICITUD_TEXTO") or "").strip()
        if val_fsol and val_fsol != "(no detectada)":
            reemplazos_fallback.extend([
                ("17 DE JULIO DE 2026", val_fsol),
                ("17 de julio de 2026", val_fsol),
                ("16 DE JULIO DE 2026", val_fsol),
                ("16 de julio de 2026", val_fsol),
                ("16 de julio del 2026", val_fsol),
                ("17 de julio del 2026", val_fsol),
                ("18 DE JULIO DE 2026", val_fsol),
                ("18 de julio de 2026", val_fsol),
                ("18 de julio del 2026", val_fsol),
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

            # Reemplazo dinámico de fecha de presentación en párrafo 2.6 y otros párrafos
            if val_fsol and val_fsol != "(no detectada)":
                patron_mesa = re.compile(
                    r'(ingresada por mesa de partes\s+el\s+)\d{1,2}\s+de\s+[a-zA-ZáéíóúÁÉÍÓÚ]+\s+de(?:l)?\s+\d{4}',
                    re.IGNORECASE
                )
                if patron_mesa.search(texto_nuevo):
                    texto_nuevo = patron_mesa.sub(rf'\g<1>{val_fsol}', texto_nuevo)

                patron_ref_a = re.compile(
                    r'(a\)\s+Solicitud ingresada por mesa de partes\s+el\s+)(?:\d{1,2}\s+de\s+[a-zA-ZáéíóúÁÉÍÓÚ]+\s+de(?:l)?\s+\d{4}|\(no detectada\))',
                    re.IGNORECASE
                )
                if patron_ref_a.search(texto_nuevo):
                    texto_nuevo = patron_ref_a.sub(rf'\g<1>{val_fsol}', texto_nuevo)

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

        # 10. Corrección global de Beca Excelencia a Beca de Excelencia Académica para Hijos de Docentes
        for p in doc.paragraphs:
            txt_corr = corregir_nombre_beca_excelencia(p.text)
            if txt_corr != p.text:
                p.text = txt_corr
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        txt_corr = corregir_nombre_beca_excelencia(p.text)
                        if txt_corr != p.text:
                            p.text = txt_corr

        # 11. Ajuste normativo por fecha de solicitud (RDE N° 067-2025 para <= 24 de julio de 2026)
        ajustar_normativa_informe_individual(doc, contexto, log=log)

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
                
        ruta_plantilla = obtener_ruta_plantilla_multiple()
        if not ruta_plantilla.exists():
            log(f"No se encontró la plantilla múltiple en {ruta_plantilla}")
            raise FileNotFoundError(f"Falta plantilla: {ruta_plantilla}")
            
        log("Cargando plantilla de informe múltiple...")
        from docxtpl import DocxTemplate
        import copy
        import shutil

        # Normalizar y expandir INSTITUCION_GLOBAL (sin sede)
        inst_g = limpiar_nombre_ies_sin_sede(super_contexto.get("INSTITUCION_GLOBAL", ""))
        if not inst_g or re.search(r"^\s*Pontificia\s+Universidad(?:\s+Cat[oó]lica)?\s*$", inst_g, re.IGNORECASE) or ("Pontificia Universidad" in inst_g and "Católica del Perú" not in inst_g and "Catolica del Peru" not in inst_g):
            inst_g = "Pontificia Universidad Católica del Perú"
        super_contexto["INSTITUCION_GLOBAL"] = inst_g
        super_contexto["IES_GLOBAL"] = inst_g
        super_contexto["UNIVERSIDAD_GLOBAL"] = inst_g
            
        beca_g = super_contexto.get("BECA_TITULO_GLOBAL", "")
        if not beca_g:
            super_contexto["BECA_TITULO_GLOBAL"] = "Beca 18 - Convocatoria 2023"

        # Expandir cada becario en super_contexto con aliases, fallbacks y variantes de casing
        becarios_limpios = []
        for b in super_contexto.get("becarios", []):
            b_exp = dict(b)
            # Asegurar institución completa SIN INFORMACIÓN DE SEDE
            inst_b = b_exp.get("INSTITUCION", "")
            inst_limpia = limpiar_nombre_ies_sin_sede(inst_b) or inst_g
            if not inst_limpia or re.search(r"^\s*Pontificia\s+Universidad(?:\s+Cat[oó]lica)?\s*$", inst_limpia, re.IGNORECASE) or ("Pontificia Universidad" in inst_limpia and "Católica del Perú" not in inst_limpia and "Catolica del Peru" not in inst_limpia):
                inst_limpia = "Pontificia Universidad Católica del Perú"
            b_exp["INSTITUCION"] = inst_limpia
            b_exp["IES"] = inst_limpia
            b_exp["UNIVERSIDAD"] = inst_limpia

            # Asegurar número de expediente del becario consistente
            exp_val = str(b_exp.get("EXPEDIENTE_PADRON") or b_exp.get("EXPEDIENTE_BECARIO") or b_exp.get("EXPEDIENTE") or b_exp.get("NUMERO_EXPEDIENTE") or "").strip()
            exp_val = re.sub(r"\.0+$", "", exp_val).strip()
            b_exp["EXPEDIENTE"] = exp_val
            b_exp["NUMERO_EXPEDIENTE"] = exp_val
            b_exp["EXPEDIENTE_BECARIO"] = exp_val
            b_exp["EXPEDIENTE_PADRON"] = exp_val
            b_exp["NEXPEDIENTE"] = exp_val
            b_exp["N_EXPEDIENTE"] = exp_val
            b_exp["NRO_EXPEDIENTE"] = exp_val
            b_exp["NUM_EXPEDIENTE"] = exp_val
            
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

            # Multi-casing para cada becario (evitar bool)
            for k, val in list(b_exp.items()):
                if not isinstance(val, bool) and isinstance(val, (str, int, float)):
                    val_str = str(val)
                    b_exp[k.lower()] = val_str
                    b_exp[k.upper()] = val_str
                    b_exp[k.title()] = val_str
            becarios_limpios.append(b_exp)

        super_contexto["becarios"] = becarios_limpios

        # Multi-casing a nivel de super_contexto (evitar bool)
        for k, val in list(super_contexto.items()):
            if not isinstance(val, bool) and isinstance(val, (str, int, float)):
                val_str = str(val)
                super_contexto[k.lower()] = val_str
                super_contexto[k.upper()] = val_str
                super_contexto[k.title()] = val_str

        # Sanitizar y definir nombre de salida de forma 100% segura
        num_inf = str(super_contexto.get("NUMERO_INFORME_GENERAR", "")).strip()
        sigedo_val = str(super_contexto.get("NUMERO_SIGEDO_GLOBAL", "")).strip()
        sigedo_corto = sigedo_val.split("-")[0] if sigedo_val else ""
        num_limpio = re.sub(r'[\\/*?:"<>|]', '_', num_inf) if num_inf else ""
        
        if num_limpio and num_limpio != "S-N":
            nombre_salida = f"Informe_{num_limpio}_Multiple.docx"
        elif sigedo_corto:
            nombre_salida = f"Informe_{sigedo_corto}_Multiple.docx"
        else:
            nombre_salida = "Informe_Multiple.docx"

        ruta_salida = SALIDA_DIR / nombre_salida
        ruta_salida.parent.mkdir(parents=True, exist_ok=True)

        # PASO 1: Renderizado mediante docxtpl (Jinja2) con protección ante fallas
        try:
            doc = DocxTemplate(str(ruta_plantilla))
            doc.render(super_contexto)
            doc.save(ruta_salida)
            log(f"Renderizado docxtpl completado en: {ruta_salida.name}")
        except Exception as err:
            log(f"  [AVISO docxtpl múltiple]: {err}. Se aplicará copia de plantilla y post-procesamiento python-docx.")
            shutil.copy(ruta_plantilla, ruta_salida)

        # PASO 2: Post-procesamiento profundo y validación con python-docx
        from docx import Document
        from docx.shared import Inches, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc_final = Document(ruta_salida)
        modificado = False

        # Corrección global: Asegurar año en SIGEDO INTEGRADO (ej. SIGEDO INTEGRADO: 59288 -> 59288-2026)
        try:
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
        except Exception as e_sig:
            log(f"  [Aviso SIGEDO post-proc]: {e_sig}")

        # Corrección global: Beca y Convocatoria real en ASUNTO y numeral 2.1
        try:
            beca_real = corregir_nombre_beca_excelencia(super_contexto.get("BECA_TITULO_GLOBAL", ""))
            if beca_real:
                for p in doc_final.paragraphs:
                    if "Beca 18 - Convocatoria 2021" in p.text:
                        p.text = p.text.replace("Beca 18 - Convocatoria 2021", beca_real)
                        modificado = True
                    if "Beca 18 - 2021" in p.text:
                        p.text = p.text.replace("Beca 18 - 2021", beca_real)
                        modificado = True
                for t in doc_final.tables:
                    for row in t.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                if "Beca 18 - Convocatoria 2021" in p.text:
                                    p.text = p.text.replace("Beca 18 - Convocatoria 2021", beca_real)
                                    modificado = True
                                if "Beca 18 - 2021" in p.text:
                                    p.text = p.text.replace("Beca 18 - 2021", beca_real)
                                    modificado = True
        except Exception as e_beca:
            log(f"  [Aviso Beca post-proc]: {e_beca}")

        # 0. Reemplazar número de informe en el título (INFORME Nº XXXX-2026...)
        try:
            num_inf_tit = super_contexto.get("NUMERO_INFORME_GENERAR", "")
            patron_num_inf = re.compile(r'(INFORME\s+N[ºo\.]?\s*)\d*(-\d{4}-MINEDU/VMGI-PRONABEC-DIBEC-SUS)', re.IGNORECASE)
            for p in doc_final.paragraphs[:10]:
                if patron_num_inf.search(p.text) or "MINEDU/VMGI-PRONABEC-DIBEC-SUS" in p.text:
                    if num_inf_tit and num_inf_tit != "S-N":
                        nuevo_txt = patron_num_inf.sub(rf"\g<1>{num_inf_tit}\g<2>", p.text)
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
        except Exception as e_tit:
            log(f"  [Aviso Título post-proc]: {e_tit}")

        # A. Actualizar nombre de la IES a Title Case
        try:
            inst_real = super_contexto.get("INSTITUCION_GLOBAL", "")
            if inst_real:
                inst_upper = inst_real.upper().strip()
                inst_title = formatear_nombre_ies(inst_real)
                inst_wrong = " ".join(w.capitalize() for w in inst_real.split())
                texto_ies_ant = "Universidad Peruana de Ciencias Aplicadas"
                texto_ies_ant_wrong = "Universidad Peruana De Ciencias Aplicadas"
                
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
        except Exception as e_ies:
            log(f"  [Aviso IES post-proc]: {e_ies}")

        # Asegurar que nunca quede 'Pontificia Universidad' sin 'Católica del Perú'
        try:
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
        except Exception as e_pucp:
            log(f"  [Aviso PUCP post-proc]: {e_pucp}")

        # B. Actualizar informe SUCCOR: cambiar a 'Informe N° ' en todas sus menciones
        try:
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
        except Exception as e_succ:
            log(f"  [Aviso SUCCOR post-proc]: {e_succ}")

        # C. Reemplazo del documento IES en numeral 2.6
        try:
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
        except Exception as e_ies26:
            log(f"  [Aviso Doc IES post-proc]: {e_ies26}")

        # D. Formato de REFERENCIAS (Tabla 0, Fila 3, Celda 2)
        try:
            referencias_lista = super_contexto.get("REFERENCIAS", [])
            if referencias_lista and len(doc_final.tables) > 0 and len(doc_final.tables[0].rows) > 3:
                row_ref = doc_final.tables[0].rows[3]
                for c_idx in (0, 1):
                    if len(row_ref.cells) > c_idx:
                        for p in row_ref.cells[c_idx].paragraphs:
                            for r in p.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(11)

                if len(row_ref.cells) > 2:
                    cell_ref = row_ref.cells[2]
                    cell_ref.text = ""
                    for i, ref_texto in enumerate(referencias_lista):
                        ref_texto_limpio = re.sub(r'\b(?:INFORME|Informe)\s+(?:N[º°o\.]*|N\.o|No)\s*(\d+-\d{4}-MINEDU/VMGI-PRONABEC)', r'Informe N° \1', ref_texto, flags=re.IGNORECASE)
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
        except Exception as e_ref:
            log(f"  [Aviso Referencias post-proc]: {e_ref}")

        # E. Numeral 2.5: Reemplazo de fechas concatenadas
        try:
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
        except Exception as e_num25:
            log(f"  [Aviso Numeral 2.5 post-proc]: {e_num25}")

        # F. Numeral 3.1: Formatear a Arial 11 pt y asegurar 'Informe N° ' SUCCOR
        try:
            patron_inf_succor = re.compile(
                r"(?:INFORME|Informe)(?:\s+(?:T[EÉ]CNICO|LEGAL|FINAL))?\s*(?:N[°ºo\.\s]*|NRO\.?|N[UÚ]MERO|NUMERO|N\.º|N\.°|N°:|N°\s*:)?\s*([\d]+[-\s\w/\.]*(?:SUCCOR[-\s\w]*LIMA|SUCCOR[-\s\w]*|DICONCI[-\s\w]*SUCCOR))",
                re.IGNORECASE
            )
            for p in doc_final.paragraphs:
                if "De conformidad a lo informado por la Subdirección" in p.text:
                    if patron_inf_succor.search(p.text):
                        p.text = patron_inf_succor.sub(r"Informe N° \1", p.text)
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(11)
                    modificado = True
        except Exception as e_num31:
            log(f"  [Aviso Numeral 3.1 post-proc]: {e_num31}")

        # G. Concordancia de Género Femenino
        try:
            todos_femeninos = bool(super_contexto.get("TODOS_FEMENINOS") is True or str(super_contexto.get("TODOS_FEMENINOS")).lower() == "true")
            if todos_femeninos:
                if len(doc_final.tables) > 0:
                    for row in doc_final.tables[0].rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                if "becarios" in p.text.lower():
                                    p.text = re.sub(r'\b(por\s+\d+\s+)becarios\b', r'\g<1>becarias', p.text, flags=re.IGNORECASE)
                                    p.text = re.sub(r'\bde\s+los\s+becarios\b', 'de las becarias', p.text, flags=re.IGNORECASE)
                                    p.text = re.sub(r'\blos\s+(\d+\s+)?becarios\b', r'las \g<1>becarias', p.text, flags=re.IGNORECASE)
                                    modificado = True

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
        except Exception as e_gen:
            log(f"  [Aviso Género post-proc]: {e_gen}")

        # H. Cuadro N° 2: Formatear Cursos Pendientes
        try:
            if len(doc_final.tables) > 2:
                t2 = doc_final.tables[2]
                for r_idx in range(2, len(t2.rows)):
                    if len(t2.rows[r_idx].cells) > 2:
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
        except Exception as e_cur:
            log(f"  [Aviso Cursos post-proc]: {e_cur}")

        # I. Cuadro N° 1: Validación y llenado garantizado de datos
        try:
            if len(doc_final.tables) > 1:
                t1 = doc_final.tables[1]
                becarios_list = super_contexto.get("becarios", [])
                
                # Si la tabla tiene menos filas que becarios, expandir filas dinámicamente
                while len(t1.rows) - 1 < len(becarios_list):
                    t1.add_row()
                    modificado = True

                for r_idx in range(1, len(t1.rows)):
                    row = t1.rows[r_idx]
                    b_idx = r_idx - 1
                    b_data = becarios_list[b_idx] if b_idx < len(becarios_list) else {}
                    
                    # Col 0: N°
                    if len(row.cells) > 0:
                        c0_txt = row.cells[0].text.strip()
                        if not c0_txt or "{{" in c0_txt:
                            row.cells[0].text = str(r_idx)
                            for p in row.cells[0].paragraphs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for r in p.runs:
                                    r.font.name = "Arial"
                                    r.font.size = Pt(8)
                            modificado = True

                    # Col 1: N° EXPEDIENTE DEL BECARIO/A
                    if len(row.cells) > 1:
                        exp_v = str(b_data.get("EXPEDIENTE_PADRON") or b_data.get("EXPEDIENTE_BECARIO") or b_data.get("EXPEDIENTE") or b_data.get("NUMERO_EXPEDIENTE") or "").strip()
                        exp_v = re.sub(r"\.0+$", "", exp_v).strip()
                        if exp_v:
                            row.cells[1].text = exp_v
                            for p in row.cells[1].paragraphs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for r in p.runs:
                                    r.font.name = "Arial"
                                    r.font.size = Pt(8)
                            modificado = True

                    # Col 2: RJ Adjudicación
                    if len(row.cells) > 2:
                        c2_txt = row.cells[2].text.strip()
                        if not c2_txt or "{{" in c2_txt:
                            rjd_v = b_data.get("RJD_ADJUDICACION") or super_contexto.get("RJD_ADJUDICACION_GLOBAL", "")
                            if rjd_v:
                                row.cells[2].text = rjd_v
                                for p in row.cells[2].paragraphs:
                                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    for r in p.runs:
                                        r.font.name = "Arial"
                                        r.font.size = Pt(8)
                                modificado = True

                    # Col 3: BECA Y CONVOCATORIA
                    if len(row.cells) > 3:
                        c3_txt = row.cells[3].text.strip()
                        if not c3_txt or "{{" in c3_txt:
                            b_c_val = b_data.get("BECA_Y_CONVOCATORIA_VALIDADA") or super_contexto.get("BECA_TITULO_GLOBAL", "Beca 18 - Convocatoria 2023")
                            b_c_val = corregir_nombre_beca_excelencia(b_c_val)
                            row.cells[3].text = b_c_val
                            for p in row.cells[3].paragraphs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for r in p.runs:
                                    r.font.name = "Arial"
                                    r.font.size = Pt(8)
                            modificado = True

                    # Col 4: INSTITUCION (SOLO NOMBRE DE LA IES, SIN SEDE)
                    if len(row.cells) > 4:
                        inst_v = b_data.get("INSTITUCION") or super_contexto.get("INSTITUCION_GLOBAL", "")
                        inst_limpia = limpiar_nombre_ies_sin_sede(inst_v)
                        if not inst_limpia or "{{" in inst_limpia:
                            inst_limpia = "Pontificia Universidad Católica del Perú"
                        row.cells[4].text = inst_limpia
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
        except Exception as e_c1:
            log(f"  [Aviso Cuadro 1 post-proc]: {e_c1}")

        # J. Cuadro N° 2: Validación y llenado de Expediente y Periodo de Estudios (Inicio y Fin)
        try:
            if len(doc_final.tables) > 2:
                t2 = doc_final.tables[2]
                becarios_list = super_contexto.get("becarios", [])
                
                # Si la tabla tiene menos filas que becarios, expandir filas dinámicamente
                while len(t2.rows) - 2 < len(becarios_list):
                    t2.add_row()
                    modificado = True

                for r_idx in range(2, len(t2.rows)):
                    row = t2.rows[r_idx]
                    b_idx = r_idx - 2
                    b_data = becarios_list[b_idx] if b_idx < len(becarios_list) else {}
                    
                    # Col 0: N°
                    if len(row.cells) > 0:
                        c0_txt = row.cells[0].text.strip()
                        if not c0_txt or "{{" in c0_txt:
                            row.cells[0].text = str(r_idx - 1)
                            for p in row.cells[0].paragraphs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for r in p.runs:
                                    r.font.name = "Arial"
                                    r.font.size = Pt(8.5)
                            modificado = True

                    # Col 1: N° EXPEDIENTE DEL BECARIO/A
                    if len(row.cells) > 1:
                        exp_v = str(b_data.get("EXPEDIENTE_PADRON") or b_data.get("EXPEDIENTE_BECARIO") or b_data.get("EXPEDIENTE") or b_data.get("NUMERO_EXPEDIENTE") or "").strip()
                        exp_v = re.sub(r"\.0+$", "", exp_v).strip()
                        if exp_v:
                            row.cells[1].text = exp_v
                            for p in row.cells[1].paragraphs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for r in p.runs:
                                    r.font.name = "Arial"
                                    r.font.size = Pt(8.5)
                            modificado = True

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
        except Exception as e_c2:
            log(f"  [Aviso Cuadro 2 post-proc]: {e_c2}")

        # Limpieza global de sede en tablas (garantizar que la institución nunca muestre sede)
        try:
            for t in doc_final.tables:
                for row in t.rows:
                    for cell in row.cells:
                        txt_c = cell.text.strip()
                        if any(w in txt_c.lower() for w in ["/ sede", "- sede", "/ filial", "/ campus"]):
                            cell.text = limpiar_nombre_ies_sin_sede(txt_c)
                            for p in cell.paragraphs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for r in p.runs:
                                    r.font.name = "Arial"
                                    r.font.size = Pt(8)
                            modificado = True
        except Exception as e_clean_tables:
            log(f"  [Aviso Limpieza Tablas Sede]: {e_clean_tables}")

        # K. Corrección global de Beca Excelencia a Beca de Excelencia Académica para Hijos de Docentes
        try:
            for p in doc_final.paragraphs:
                txt_corr = corregir_nombre_beca_excelencia(p.text)
                if txt_corr != p.text:
                    p.text = txt_corr
                    modificado = True
            for t in doc_final.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            txt_corr = corregir_nombre_beca_excelencia(p.text)
                            if txt_corr != p.text:
                                p.text = txt_corr
                                modificado = True
        except Exception as e_beca_ex:
            log(f"  [Aviso Beca Excelencia post-proc]: {e_beca_ex}")

        # L. Ajuste normativo por fechas de solicitud (RDE N° 067-2025 y RDE N° 119-2026)
        try:
            ajustar_normativa_informe_multiple(doc_final, super_contexto, log=log)
        except Exception as e_norm:
            log(f"  [Aviso Normativa Múltiple post-proc]: {e_norm}")

        # Guardado garantizado
        try:
            doc_final.save(ruta_salida)
            log(f"Informe Múltiple generado exitosamente en: {ruta_salida.name}")
        except Exception as e_save:
            log(f"  [Aviso guardado final]: {e_save}")

        return ruta_salida
