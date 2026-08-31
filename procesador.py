"""
Motor de procesamiento: extracción de datos de 4 PDFs, cruce con padrón Excel
y preparación del contexto para rellenar la plantilla Word.
"""

from __future__ import annotations

import re
from datetime import datetime, date
from pathlib import Path
from typing import Callable, Any

import pandas as pd
import pdfplumber


LogCallback = Callable[[str], None]
ProgressCallback = Callable[[float, str], None]


class BecarioNoEncontradoIESException(Exception):
    """Excepción lanzada cuando el becario no se encuentra en el Excel de la IES."""
    pass


class FechaFinInsuficienteException(Exception):
    """Excepción lanzada cuando la fecha fin del becario es insuficiente para el semestre."""
    pass


class BecarioNoCulminariaAmpliacionException(Exception):
    """Excepción lanzada cuando el Informe SUCCOR indica que el becario no culminaría estudios con la ampliación."""
    pass

MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}

MESES_NUM = {v: k for k, v in MESES_ES.items()}
MESES_NUM.update({
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
})


def fecha_a_texto(d: date) -> str:
    """Convierte un objeto date a texto español. Ej: '15 de julio de 2026'."""
    return f"{d.day} de {MESES_ES[d.month]} de {d.year}"


def fecha_a_corto(d: date) -> str:
    """Convierte un objeto date a formato dd/mm/aaaa. Ej: '23/08/2021'."""
    return f"{d.day:02d}/{d.month:02d}/{d.year}"


def numero_ordinal(n: int) -> str:
    """Retorna el ordinal español para un número de ciclo. Ej: 1->'1er', 2->'2do', 3->'3er'."""
    sufijos = {1: "er", 2: "do", 3: "er", 4: "to", 5: "to",
               6: "to", 7: "mo", 8: "vo", 9: "no"}
    sufijo = sufijos.get(n, "mo")
    return f"{n}{sufijo}"


def formatear_curso_oracion(c: str) -> str:
    """Convierte un curso a tipo oración pero mantiene números romanos al final."""
    if not c:
        return c
    c_formateado = str(c).strip().capitalize()
    return re.sub(r'\b([ivxlcdm]+)$', lambda m: m.group(1).upper(), c_formateado, flags=re.IGNORECASE)


def texto_a_fecha(texto: str) -> date | None:
    """Intenta parsear textos como '15 de julio de 2026', '15 de julio del 2026' o '15/07/2026'."""
    if not texto:
        return None
    texto = texto.strip().lower()

    # Formato largo: 15 de julio de 2026 o 15 de julio del 2026
    m = re.search(
        r"(\d{1,2})\s+de\s+([a-zñáéíóú]+)(?:\s+de(?:l)?)?\s+(\d{4})", texto, re.IGNORECASE
    )
    if m:
        dia, mes_texto, anio = int(m.group(1)), m.group(2).lower(), int(m.group(3))
        mes = MESES_NUM.get(mes_texto)
        if mes:
            try:
                return date(anio, mes, dia)
            except ValueError:
                pass

    # Formato corto: 15/07/2026, 15-07-2026, 15.07.2026
    m = re.search(r"(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})", texto)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass

    # Formato ISO: 2026-07-15
    m = re.search(r"(\d{4})[/\-\.](\d{1,2})[/\-\.](\d{1,2})", texto)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    return None


def limpiar_dni(val: str | int | float | None) -> str:
    """Limpia cadenas o números DNI eliminando sufijos .0 de Pandas y completando con ceros a 8 dígitos."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r"\D", "", s)
    if len(digits) in (7, 8):
        return digits.zfill(8)
    return digits


def calcular_semestre_anterior(semestre: str) -> str:
    """Dado '2026-II' retorna '2026-I'. Dado '2026-I' retorna '2025-II'."""
    m = re.match(r"(\d{4})-(I{1,2})", semestre.strip().upper())
    if not m:
        return semestre
    anio, periodo = int(m.group(1)), m.group(2)
    if periodo == "II":
        return f"{anio}-I"
    else:
        return f"{anio - 1}-II"


def construir_tabla_ciclos(fecha_inicio: date, fecha_fin: date) -> list[dict]:
    """
    Genera la tabla de ciclos académicos entre fecha_inicio y fecha_fin.
    Retorna lista de dicts con soporte de claves en mayúsculas y minúsculas.
    """
    if fecha_inicio.month <= 7:
        semestre_actual = f"{fecha_inicio.year}-I"
    else:
        semestre_actual = f"{fecha_inicio.year}-II"

    filas = []
    n = 1
    while True:
        m_ciclo = re.match(r"(\d{4})-(I{1,2})", semestre_actual)
        if not m_ciclo:
            break
        anio_ciclo = int(m_ciclo.group(1))
        periodo = m_ciclo.group(2)
        if periodo == "I":
            fecha_fin_ciclo = date(anio_ciclo, 7, 31)
        else:
            fecha_fin_ciclo = date(anio_ciclo, 12, 31)

        ordinal = numero_ordinal(n)
        nombre_ciclo = f"{ordinal} ciclo"
        filas.append({
            # momento vacío: se fusiona con 'Duración de estudios' en generador_word
            "momento": "",
            "MOMENTO": "",
            "Momento": "",
            "ciclo": nombre_ciclo,
            "CICLO": nombre_ciclo,
            "Ciclo": nombre_ciclo,
            "semestre": semestre_actual,
            "SEMESTRE": semestre_actual,
            "Semestre": semestre_actual,
            "nro": f"{n}",
            "NRO": f"{n}",
            "periodo": semestre_actual,
            "PERIODO": semestre_actual,
        })
        n += 1

        if fecha_fin_ciclo >= fecha_fin or semestre_actual == "2026-I":
            break
        if n > 20:
            break

        semestre_actual = (
            f"{anio_ciclo}-II" if periodo == "I" else f"{anio_ciclo + 1}-I"
        )

    return filas


# ============================================================
# EXTRACTOR DOCUMENTO 1: Formato autogenerado (fecha solicitud)
# ============================================================
class ExtractorFormatoAutogenerado:
    """Extrae la fecha de ingreso a mesa de partes del PDF autogenerado."""

    PATRON_FECHA_HORA = re.compile(
        r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})\s+(\d{1,2}):(\d{2})",
    )
    PATRON_FECHA_TEXTO = re.compile(
        r"(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})",
        re.IGNORECASE,
    )
    PALABRAS_CLAVE = (
        "mesa de partes", "ingreso", "recepcion", "recepción",
        "fecha", "registro", "solicitud",
    )
    PATRON_EXPEDIENTE = re.compile(
        r"Expediente\s*[:\s]+([\d]+)",
        re.IGNORECASE,
    )
    PATRON_CORREO = re.compile(
        r"Autorizo recibir la respuesta al correo electr[oó]nico:\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,})",
        re.IGNORECASE,
    )
    PATRON_CASILLA = re.compile(
        r"Autorizo recibir la respuesta por casilla electr[oó]nica",
        re.IGNORECASE,
    )
    PATRON_TELEFONO = re.compile(
        r"(?:Tel[eé]fono|Celular)\s*[:\s]+(\d{7,12})",
        re.IGNORECASE,
    )

    @classmethod
    def extraer(cls, ruta_pdf: str | Path) -> dict:
        ruta = Path(ruta_pdf)
        resultado = {"fecha_solicitud": None, "fecha_solicitud_texto": "", "numero_expediente": "", "correo_electronico": ""}

        with pdfplumber.open(ruta) as pdf:
            texto_completo = "\n".join(
                p.extract_text() or "" for p in pdf.pages
            )

        # Buscar fecha/hora cercana a palabras clave
        texto_lower = texto_completo.lower()
        mejor_pos = None
        for palabra in cls.PALABRAS_CLAVE:
            pos = texto_lower.find(palabra)
            if pos != -1:
                if mejor_pos is None or pos < mejor_pos:
                    mejor_pos = pos

        zona = texto_completo[max(0, (mejor_pos or 0) - 50): (mejor_pos or 0) + 300] if mejor_pos else texto_completo

        fecha = None
        m = cls.PATRON_FECHA_HORA.search(zona)
        if m:
            try:
                fecha = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                pass

        if not fecha:
            m = cls.PATRON_FECHA_HORA.search(texto_completo)
            if m:
                try:
                    fecha = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                except ValueError:
                    pass

        if not fecha:
            m = cls.PATRON_FECHA_TEXTO.search(texto_completo)
            if m:
                fecha = texto_a_fecha(m.group(0))

        if fecha:
            resultado["fecha_solicitud"] = fecha
            resultado["fecha_solicitud_texto"] = fecha_a_texto(fecha)

        # Extraer número de Expediente (SIGEDO) del formato autogenerado
        m_exp = cls.PATRON_EXPEDIENTE.search(texto_completo)
        if m_exp:
            resultado["numero_expediente"] = m_exp.group(1).strip()
            
        m_corr = cls.PATRON_CORREO.search(texto_completo)
        if m_corr:
            resultado["correo_electronico"] = m_corr.group(1).strip()

        m_casilla = cls.PATRON_CASILLA.search(texto_completo)
        if m_casilla:
            resultado["autoriza_casilla"] = True
        else:
            resultado["autoriza_casilla"] = False

        m_tel = cls.PATRON_TELEFONO.search(texto_completo)
        if m_tel:
            resultado["telefono_contacto"] = m_tel.group(1).strip()
        else:
            resultado["telefono_contacto"] = ""

        resultado["raw_text"] = texto_completo
        return resultado


# ============================================================
# EXTRACTOR DOCUMENTO 2: Informe SUCCOR
# ============================================================
class ExtractorInformeSuccor:
    """Extrae metadatos del Informe de la SUCCOR soportando tablas horizontales y verticales."""

    PATRON_NOMBRE_INFORME = re.compile(
        r"(INFORME\s+(?:N[°ºo\.]?|NRO\.?|NÚMERO|N)?\s*[\d]+[-\s\w/\-\.]+(?:SUCCOR[-\w]*)?)",
        re.IGNORECASE,
    )
    PATRON_SIGEDO = re.compile(
        r"(?:SIGEDO|Expediente|CUT|Trámite|Tramite)\s*(?:N[°ºo\.]?\s*)?([A-Z0-9\.\-]{4,18})",
        re.IGNORECASE,
    )
    PATRONES_DNI_TEXTO = [
        re.compile(r"(?:D\.?N\.?I\.?|DOC(?:UMENTO)?(?:\s+DE)?\s+IDENTIDAD|N[°ºo\.]?\s*DNI|N[°ºo\.]?\s*DOC)\s*[:/º°\s\-]*(\d{7,8})\b", re.IGNORECASE),
        re.compile(r"\b(\d{7,8})\s*(?:\(DNI\)|DNI)\b", re.IGNORECASE),
        re.compile(r"identificad[oa]\s+con\s+(?:D\.?N\.?I\.?\s*)?(\d{7,8})\b", re.IGNORECASE),
    ]
    PATRON_SEMESTRE = re.compile(
        r"\b(20\d{2})\s*-\s*(I{1,2})\b",
    )
    PATRON_RJD = re.compile(
        r"(RJD\s*N[°ºo\.]?\s*[\d\-/\w]+)",
        re.IGNORECASE,
    )
    PATRON_IES = re.compile(
        r"(?:IES|Institución|Institucion)\s*[:/]?\s*(?:Sede\s*[:/]?)?\s*(.+?)(?:\n|Carrera|Semestre|$)",
        re.IGNORECASE,
    )
    PATRON_CARRERA = re.compile(
        r"Carrera\s+(?:Profesional)?\s*[:/]?\s*(.+?)(?:\n|Semestre|IES|$)",
        re.IGNORECASE,
    )
    PATRON_NOMBRE_BECARIO = re.compile(
        r"(?:Becario\(a\)|Becario|Nombres y Apellidos|Apellidos y Nombres|Nombre del Becario|Estudiante)\s*[:/]?\s*([A-Za-zÁÉÍÓÚÑáéíóúñ\s,]{5,60})(?:\n|DNI|RUT|Código|IES|$)",
        re.IGNORECASE,
    )
    # Palabras clave para detectar la columna "Cursos Pendientes según la IES" en las tablas del SUCCOR
    PALABRAS_COLUMNA_CURSOS_IES = (
        "cursos pendientes seg",  # cubre "según" con o sin tilde
        "cursos pendientes segun",
        "cursos pendientes de la ies",
        "cursos seg",
        "pendientes ies",
        "pendientes seg",
        "segun la ies",
        "según la ies",
        "cursos ies",
    )

    @classmethod
    def extraer(cls, ruta_pdf: str | Path) -> dict:
        ruta = Path(ruta_pdf)
        resultado = {
            "nombre_informe_succor": "",
            "numero_informe": "",
            "numero_sigedo": "",
            "dni_succor": "",
            "nombre_aproximado": "",
            "semestre_solicitado": "",
            "rjd_adjudicacion": "",
            "institucion": "",
            "carrera": "",
            "cursos_pendientes_ies": [],
            "cursos_pendientes_ies_texto": "",
            "genero": "",
        }

        with pdfplumber.open(ruta) as pdf:
            textos = []
            tablas = []
            for p in pdf.pages:
                try:
                    textos.append(p.extract_text() or "")
                except Exception:
                    pass
                try:
                    t_list = p.extract_tables()
                    if t_list:
                        tablas.extend(t_list)
                except Exception:
                    pass
            texto_completo = "\n".join(textos)

        # Validar heurísticamente si el informe indica que no culminaría con la ampliación
        patron_no_culmina = re.compile(
            r"\bno\s+(?:culminar[ií]a|acabar[ií]a|terminar[ií]a|finalizar[ií]a|concluir[ií]a).{0,50}?(?:estudios|carrera|2026-ii|ampliaci[oó]n)\b",
            re.IGNORECASE
        )
        if patron_no_culmina.search(texto_completo):
            raise BecarioNoCulminariaAmpliacionException("En el Informe de la SUCCOR se señala que el becario no culminaria estudios con la ampliacion. Revisar el contenido de dicho Informe")

        # Nombre completo del informe
        m = cls.PATRON_NOMBRE_INFORME.search(texto_completo)
        if m:
            resultado["nombre_informe_succor"] = cls._limpiar(m.group(1))
            resultado["numero_informe"] = cls._limpiar(m.group(1))
        else:
            texto_lower = texto_completo.lower()
            mejor_pos = texto_lower.find("succor")
            if mejor_pos != -1:
                zona = texto_completo[max(0, mejor_pos - 400): mejor_pos + 400]
                patron_flexible = re.compile(r"(INFORME\s+(?:N(?:ro|RO|\.|°|º|\.o|o\.?|o)?|NÚMERO|NUMERO|N)?\s*[\d]+[-/\d]+[A-Z0-9/\-\.]+)", re.IGNORECASE)
                m_flex = patron_flexible.search(zona)
                if m_flex:
                    resultado["nombre_informe_succor"] = cls._limpiar(m_flex.group(1))
                    resultado["numero_informe"] = cls._limpiar(m_flex.group(1))
            
            if not resultado["nombre_informe_succor"]:
                patron_flexible = re.compile(r"(INFORME\s+(?:N(?:ro|RO|\.|°|º|\.o|o\.?|o)?|NÚMERO|NUMERO|N)?\s*[\d]+[-/\d]+[A-Z0-9/\-\.]+)", re.IGNORECASE)
                m_flex = patron_flexible.search(texto_completo)
                if m_flex:
                    resultado["nombre_informe_succor"] = cls._limpiar(m_flex.group(1))
                    resultado["numero_informe"] = cls._limpiar(m_flex.group(1))

        # SIGEDO
        m = cls.PATRON_SIGEDO.search(texto_completo)
        if m:
            resultado["numero_sigedo"] = m.group(1)

        # Semestre solicitado
        for m in cls.PATRON_SEMESTRE.finditer(texto_completo):
            sem = f"{m.group(1)}-{m.group(2)}"
            resultado["semestre_solicitado"] = sem

        # 1. Extracción de tablas (soportando tablas verticales y horizontales)
        for tabla in tablas:
            if not tabla or len(tabla) == 0:
                continue

            # A. Tablas verticales (Fila 0 = Encabezados, Filas 1..N = Datos)
            header_row = [str(c or "").lower().strip() for c in tabla[0]]
            col_indices = {}
            for j, h in enumerate(header_row):
                if any(w in h for w in ("dni", "documento", "identidad")):
                    col_indices["dni"] = j
                elif any(w in h for w in ("becario", "apellidos", "nombres", "estudiante")):
                    col_indices["nombre"] = j
                elif "carrera" in h:
                    col_indices["carrera"] = j
                elif "ies" in h or "instituci" in h:
                    col_indices["ies"] = j
                elif "rjd" in h or "adjudicaci" in h:
                    col_indices["rjd"] = j
                # Detectar columna de cursos pendientes según la IES
                elif any(kw in h for kw in cls.PALABRAS_COLUMNA_CURSOS_IES):
                    col_indices["cursos_ies"] = j

            for row in tabla[1:]:
                if not row:
                    continue
                if "dni" in col_indices and col_indices["dni"] < len(row):
                    d_digits = re.sub(r"\D", "", str(row[col_indices["dni"]] or ""))
                    if len(d_digits) in (7, 8) and not resultado["dni_succor"]:
                        resultado["dni_succor"] = d_digits.zfill(8)
                if "nombre" in col_indices and col_indices["nombre"] < len(row):
                    val_nom = str(row[col_indices["nombre"]] or "").strip()
                    if len(val_nom) > 4 and not resultado["nombre_aproximado"]:
                        resultado["nombre_aproximado"] = cls._limpiar(val_nom)
                if "carrera" in col_indices and col_indices["carrera"] < len(row):
                    val_car = str(row[col_indices["carrera"]] or "").strip()
                    if val_car and not resultado["carrera"]:
                        resultado["carrera"] = cls._limpiar(val_car)
                if "ies" in col_indices and col_indices["ies"] < len(row):
                    val_ies = str(row[col_indices["ies"]] or "").strip()
                    if val_ies and not resultado["institucion"]:
                        resultado["institucion"] = cls._limpiar(val_ies)
                if "rjd" in col_indices and col_indices["rjd"] < len(row):
                    val_rjd = str(row[col_indices["rjd"]] or "").strip()
                    if val_rjd and not resultado["rjd_adjudicacion"]:
                        resultado["rjd_adjudicacion"] = cls._limpiar(val_rjd)
                # Extraer cursos pendientes según la IES de esta columna
                if "cursos_ies" in col_indices and col_indices["cursos_ies"] < len(row):
                    val_cursos = str(row[col_indices["cursos_ies"]] or "").strip()
                    if val_cursos:
                        # La celda puede tener varios cursos separados por saltos de línea o numeración
                        cursos_celda = cls._parsear_cursos_celda(val_cursos)
                        for c in cursos_celda:
                            if c:
                                c_fmt = formatear_curso_oracion(c)
                                if c_fmt not in resultado["cursos_pendientes_ies"]:
                                    resultado["cursos_pendientes_ies"].append(c_fmt)

            # B. Tablas horizontales (Clave en celda j, Valor en celda j+1)
            for fila in tabla:
                if not fila:
                    continue
                texto_fila = " ".join(str(c or "").lower() for c in fila)
                if any(w in texto_fila for w in ("dni", "documento", "identidad")) and not resultado["dni_succor"]:
                    for i, celda in enumerate(fila):
                        c_text = str(celda or "").lower()
                        if any(w in c_text for w in ("dni", "documento", "identidad")):
                            if i + 1 < len(fila) and fila[i + 1]:
                                d_digits = re.sub(r"\D", "", str(fila[i + 1]))
                                if len(d_digits) in (7, 8):
                                    resultado["dni_succor"] = d_digits.zfill(8)
                if any(w in texto_fila for w in ("becario", "apellidos", "nombres", "estudiante")) and not resultado["nombre_aproximado"]:
                    for i, celda in enumerate(fila):
                        c_text = str(celda or "").lower()
                        if any(w in c_text for w in ("becario", "apellidos", "nombres", "estudiante")):
                            if i + 1 < len(fila) and fila[i + 1]:
                                val_nom = str(fila[i + 1]).strip()
                                if len(val_nom) > 4:
                                    resultado["nombre_aproximado"] = cls._limpiar(val_nom)

        # 2. Respaldo por patrones regex en texto explícito
        if not resultado["dni_succor"]:
            for pat in cls.PATRONES_DNI_TEXTO:
                m_dni = pat.search(texto_completo)
                if m_dni:
                    resultado["dni_succor"] = m_dni.group(1).zfill(8)
                    break

        if not resultado["nombre_aproximado"]:
            m_nom = cls.PATRON_NOMBRE_BECARIO.search(texto_completo)
            if m_nom:
                resultado["nombre_aproximado"] = cls._limpiar(m_nom.group(1))

        if not resultado["rjd_adjudicacion"]:
            m = cls.PATRON_RJD.search(texto_completo)
            if m:
                resultado["rjd_adjudicacion"] = cls._limpiar(m.group(1))

        if not resultado["institucion"]:
            m = cls.PATRON_IES.search(texto_completo)
            if m:
                resultado["institucion"] = cls._limpiar(m.group(1))

        if not resultado["carrera"]:
            m = cls.PATRON_CARRERA.search(texto_completo)
            if m:
                resultado["carrera"] = cls._limpiar(m.group(1))

        # Si no se encontraron cursos en tablas, buscar en texto plano del SUCCOR
        if not resultado["cursos_pendientes_ies"]:
            resultado["cursos_pendientes_ies"] = cls._extraer_cursos_texto_plano(texto_completo)

        # Construir texto numerado final
        if resultado["cursos_pendientes_ies"]:
            resultado["cursos_pendientes_ies_texto"] = "\n".join(
                f"{i+1}. {c}" for i, c in enumerate(resultado["cursos_pendientes_ies"])
            )
            
        # Determinar género basado en la palabra 'becaria'
        if re.search(r"\bbecaria\b", texto_completo, re.IGNORECASE):
            resultado["genero"] = "F"
        elif re.search(r"\bbecario\b", texto_completo, re.IGNORECASE):
            resultado["genero"] = "M"

        resultado["raw_text"] = texto_completo
        return resultado

    @classmethod
    def _parsear_cursos_celda(cls, texto: str) -> list[str]:
        """Divide el contenido de una celda en nombres de cursos individuales.
        Soporta saltos de línea, numeración (1., 2., -) y texto corrido."""
        cursos = []
        # Separar por saltos de línea primero
        lineas = re.split(r"[\n\r]+", texto)
        for linea in lineas:
            linea = linea.strip()
            if not linea:
                continue
            # Quitar prefijo numérico o viñeta: "1.", "1)", "-", "•"
            linea = re.sub(r"^[\d]+[\.\.\)\-]\s*", "", linea).strip()
            linea = re.sub(r"^[-•·]\s*", "", linea).strip()
            if len(linea) > 3:
                cursos.append(linea)
        # Si no hubo saltos de línea, intentar separar por numeración en el mismo texto
        if len(cursos) <= 1 and len(texto) > 10:
            partes = re.split(r"(?=\d+\.\s+[A-ZÁÉÍÓÚ])", texto)
            if len(partes) > 1:
                cursos = []
                for parte in partes:
                    parte = re.sub(r"^[\d]+\.\s*", "", parte).strip()
                    if len(parte) > 3:
                        cursos.append(formatear_curso_oracion(parte))
        return cursos

    @classmethod
    def _extraer_cursos_texto_plano(cls, texto: str) -> list[str]:
        """Busca en el texto plano del SUCCOR la sección de cursos pendientes según la IES
        y extrae los nombres de los cursos (como fallback si no se detectaron en tablas).
        Maneja el caso de PDFs con dos columnas paralelas que pdfplumber une en la misma línea,
        generando texto duplicado como: '1. Teoría de las Relaciones 1. Teoría de las Relaciones'."""
        # Buscar el bloque que menciona los cursos pendientes según la IES
        patron_bloque = re.compile(
            r"(cursos?\s+pendientes?\s+seg[uú]n\s+(?:la\s+)?ies|cursos?\s+pendientes?\s+de\s+(?:la\s+)?ies|seg[uú]n\s+(?:la\s+)?ies)",
            re.IGNORECASE,
        )
        m = patron_bloque.search(texto)
        if not m:
            return []
        # Tomar hasta 1500 caracteres desde el inicio del bloque
        bloque = texto[m.start(): m.start() + 1500]
        # Extraer líneas con numeración o viñeta
        patron_item = re.compile(r"^[\s]*(?:\d+[\.\)]\s*|[-•·]\s*)(.+)$", re.MULTILINE)
        cursos = []
        for m_item in patron_item.finditer(bloque):
            curso = m_item.group(1).strip()
            # Filtrar líneas cortas o que son encabezados de columna
            if len(curso) > 4 and not re.search(r"cursos?\s+pendientes?|seg[uú]n|ies\b", curso, re.IGNORECASE):
                # Limpiar duplicación por columnas paralelas en PDFs de dos columnas:
                # Ej: 'Teoría de las Relaciones 1. Teoría de las Relaciones' -> 'Teoría de las Relaciones'
                curso = cls._limpiar_duplicacion_columna(curso)
                if curso and len(curso) > 3:
                    curso_fmt = formatear_curso_oracion(curso)
                    if curso_fmt not in cursos:
                        cursos.append(curso_fmt)
        return cursos

    @staticmethod
    def _limpiar_duplicacion_columna(texto: str) -> str:
        """Detecta y elimina duplicación de texto causada por columnas paralelas en PDFs.
        Cuando pdfplumber extrae texto de dos columnas en la misma línea, el curso aparece dos veces.
        Ej: 'Teoría de las Relaciones 1. Teoría de las Relaciones' -> 'Teoría de las Relaciones'
        Ej: 'Seminario de 3.Seminario de' -> 'Seminario de'"""
        # Detectar marcador numérico embebido mid-texto (indica inicio de segunda columna)
        # Patron: espacio(s) + digito(s) + punto/paren + texto continuación
        m_dup = re.search(r"\s+\d+[\.)]\s*", texto)
        if m_dup:
            primera = texto[:m_dup.start()].strip()
            # Extraer la segunda parte quitando el prefijo numérico (espacios + N. + espacios)
            segunda = re.sub(r"^\s*\d+[\.)]\s*", "", texto[m_dup.start():]).strip()
            if not primera or not segunda:
                return texto
            # Solo truncar si la segunda parte es similar a la primera (columna duplicada)
            norm1 = re.sub(r"\s+", " ", primera.lower())
            norm2 = re.sub(r"\s+", " ", segunda.lower())
            cmp_len = min(len(norm1), len(norm2), 12)
            if cmp_len > 0 and (
                norm1[:cmp_len] == norm2[:cmp_len]
                or norm2.startswith(norm1[:8])
                or norm1.startswith(norm2[:8])
            ):
                return primera
        return texto

    @staticmethod
    def _limpiar(valor: str) -> str:
        v = re.sub(r"\s+", " ", valor.replace("\n", " ")).strip(" :-\t")
        # Asegurar espacio después de N° o Nº si le sigue texto pegado
        v = re.sub(r'\b(N[°º])([^\s])', r'\1 \2', v, flags=re.IGNORECASE)
        return v


# ============================================================
# EXTRACTOR DOCUMENTO 3: Calendario Académico
# ============================================================

    @classmethod
    def extraer_multiple(cls, ruta_pdf: str | Path) -> dict:
        import pdfplumber
        ruta = Path(ruta_pdf)
        resultado = {
            "nombre_informe_succor": "",
            "numero_informe": "",
            "numero_sigedo": "",
            "semestre_solicitado": "",
            "rjd_adjudicacion": "",
            "institucion": "",
            "carrera": "",
            "becarios": []
        }
        with pdfplumber.open(ruta) as pdf:
            textos = []
            tablas = []
            for p in pdf.pages:
                try:
                    textos.append(p.extract_text() or "")
                except Exception:
                    pass
                try:
                    t_list = p.extract_tables()
                    if t_list:
                        tablas.extend(t_list)
                except Exception:
                    pass
            texto_completo = "\n".join(textos)
        
        m = cls.PATRON_NOMBRE_INFORME.search(texto_completo)
        if m:
            resultado["nombre_informe_succor"] = cls._limpiar(m.group(1))
            resultado["numero_informe"] = cls._limpiar(m.group(1))
        
        m = cls.PATRON_SIGEDO.search(texto_completo)
        if m:
            resultado["numero_sigedo"] = m.group(1)
            
        for m in cls.PATRON_SEMESTRE.finditer(texto_completo):
            resultado["semestre_solicitado"] = f"{m.group(1)}-{m.group(2)}"
            
        # Extract global values
        m = cls.PATRON_RJD.search(texto_completo)
        if m: resultado["rjd_adjudicacion"] = cls._limpiar(m.group(1))
        m = cls.PATRON_IES.search(texto_completo)
        if m: resultado["institucion"] = cls._limpiar(m.group(1))
        
        # Buscar becarios en tablas
        for tabla in tablas:
            if not tabla or len(tabla) == 0: continue
            header_row = [str(c or "").lower().strip() for c in tabla[0]]
            
            col_expediente = -1
            col_dni = -1
            col_nombre = -1
            col_cursos = -1
            col_ies = -1
            
            for j, h in enumerate(header_row):
                if "expediente" in h or "sigedo" in h: col_expediente = j
                elif "dni" in h or "documento" in h: col_dni = j
                elif "becario" in h or "nombres" in h or "apellidos" in h: col_nombre = j
                elif any(kw in h for kw in cls.PALABRAS_COLUMNA_CURSOS_IES): col_cursos = j
                elif "ies" in h or "instituci" in h: col_ies = j
            
            # Asumimos que es una tabla de becarios si tiene expediente o dni o nombre
            if col_expediente != -1 or col_dni != -1 or col_nombre != -1:
                for row in tabla[1:]:
                    if not row: continue
                    bec = {"expediente": "", "dni": "", "nombre": "", "cursos_pendientes_ies": []}
                    if col_expediente != -1 and col_expediente < len(row):
                        bec["expediente"] = str(row[col_expediente] or "").strip()
                    if col_dni != -1 and col_dni < len(row):
                        d_digits = re.sub(r"\D", "", str(row[col_dni] or ""))
                        if len(d_digits) in (7, 8): bec["dni"] = d_digits.zfill(8)
                    if col_nombre != -1 and col_nombre < len(row):
                        bec["nombre"] = cls._limpiar(str(row[col_nombre] or ""))
                    if col_cursos != -1 and col_cursos < len(row):
                        c_text = str(row[col_cursos] or "").strip()
                        if c_text:
                            for c in cls._parsear_cursos_celda(c_text):
                                if c: bec["cursos_pendientes_ies"].append(formatear_curso_oracion(c))
                    
                    if bec["expediente"] or bec["dni"] or bec["nombre"]:
                        resultado["becarios"].append(bec)

        resultado["raw_text"] = texto_completo
        return resultado

class ExtractorCalendarioAcademico:
    """Extrae fechas de matrícula e inicio de estudios del Calendario Académico.
    Analiza TODAS las páginas (incluyendo la pág 1)."""

    PATRON_FECHA = re.compile(
        r"(?:del\s+)?(\d{1,2})\s+(?:al\s+\d{1,2}\s+)?de\s+([a-zñáéíóú]+)\s+de\s+(\d{4})|(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})",
        re.IGNORECASE,
    )
    PALABRAS_MATRICULA = (
        "matrícula", "matricula", "matrículas", "matriculas",
        "periodo de matrícula", "proceso de matrícula", "inscripción", "inscripcion",
        "matrcula", "matr"
    )
    PALABRAS_INICIO = (
        "inicio de clases", "inicio de estudios", "inicio del semestre",
        "inicio de actividades", "inicio lectivo", "inicio de ciclo", "inicio"
    )

    @classmethod
    def extraer(
        cls,
        ruta_pdf: str | Path,
        semestre_solicitado: str = "",
        ies_filtro: str = "",
    ) -> dict:
        ruta = Path(ruta_pdf)
        resultado = {
            "fecha_matricula": None,
            "fecha_matricula_texto": "",
            "fecha_inicio_estudios": None,
            "fecha_inicio_estudios_texto": "",
        }

        with pdfplumber.open(ruta) as pdf:
            textos = []
            tablas = []
            for p in pdf.pages:
                try:
                    textos.append(p.extract_text() or "")
                except Exception:
                    pass
                try:
                    t_list = p.extract_tables()
                    if t_list:
                        tablas.extend(t_list)
                except Exception:
                    pass
            texto_completo = "\n".join(textos)

        # 1. Búsqueda estructurada en tablas (prioridad)
        if semestre_solicitado:
            sem_lower = semestre_solicitado.lower()
            for tabla in tablas:
                if not tabla or len(tabla) < 2:
                    continue
                
                header_row = [str(c or "").lower().strip() for c in tabla[0]]
                col_mat = -1
                col_ini = -1
                col_per = -1
                
                for i, h in enumerate(header_row):
                    if col_mat == -1 and any(p in h for p in cls.PALABRAS_MATRICULA):
                        col_mat = i
                    if col_ini == -1 and any(p in h for p in cls.PALABRAS_INICIO):
                        col_ini = i
                    if col_per == -1 and ("periodo" in h or "semestre" in h):
                        col_per = i
                
                if col_per != -1 and (col_mat != -1 or col_ini != -1):
                    for fila in tabla[1:]:
                        if not fila or len(fila) <= col_per:
                            continue
                        celda_per = str(fila[col_per] or "").lower()
                        if sem_lower in celda_per:
                            if col_mat != -1 and len(fila) > col_mat and not resultado["fecha_matricula"]:
                                f_mat = cls._extraer_primera_fecha(str(fila[col_mat] or ""))
                                if f_mat:
                                    resultado["fecha_matricula"] = f_mat
                                    resultado["fecha_matricula_texto"] = fecha_a_texto(f_mat)
                            if col_ini != -1 and len(fila) > col_ini and not resultado["fecha_inicio_estudios"]:
                                f_ini = cls._extraer_primera_fecha(str(fila[col_ini] or ""))
                                if f_ini:
                                    resultado["fecha_inicio_estudios"] = f_ini
                                    resultado["fecha_inicio_estudios_texto"] = fecha_a_texto(f_ini)

        # Filtrar bloque relevante por semestre e IES si se proporcionan
        bloque = cls._filtrar_bloque(texto_completo, semestre_solicitado, ies_filtro)

        # Buscar fecha de matrícula en texto (fallback)
        if not resultado["fecha_matricula"]:
            for palabra in cls.PALABRAS_MATRICULA:
                pos = bloque.lower().find(palabra)
                if pos != -1:
                    zona = bloque[pos: pos + 250]
                    fecha = cls._extraer_primera_fecha(zona)
                    if fecha:
                        resultado["fecha_matricula"] = fecha
                        resultado["fecha_matricula_texto"] = fecha_a_texto(fecha)
                        break

        # Buscar fecha de inicio de estudios en texto (fallback)
        if not resultado["fecha_inicio_estudios"]:
            for palabra in cls.PALABRAS_INICIO:
                pos = bloque.lower().find(palabra)
                if pos != -1:
                    zona = bloque[pos: pos + 250]
                    fecha = cls._extraer_primera_fecha(zona)
                    if fecha:
                        resultado["fecha_inicio_estudios"] = fecha
                        resultado["fecha_inicio_estudios_texto"] = fecha_a_texto(fecha)
                        break

        # Fallback: buscar en tablas sin estructura clara
        if not resultado["fecha_matricula"] or not resultado["fecha_inicio_estudios"]:
            for tabla in tablas:
                for fila in tabla:
                    if not fila:
                        continue
                    texto_fila = " ".join(str(c or "") for c in fila).lower()
                    for palabra in cls.PALABRAS_MATRICULA:
                        if palabra in texto_fila and not resultado["fecha_matricula"]:
                            for celda in fila:
                                f = cls._extraer_primera_fecha(str(celda or ""))
                                if f:
                                    resultado["fecha_matricula"] = f
                                    resultado["fecha_matricula_texto"] = fecha_a_texto(f)
                    for palabra in cls.PALABRAS_INICIO:
                        if palabra in texto_fila and not resultado["fecha_inicio_estudios"]:
                            for celda in fila:
                                f = cls._extraer_primera_fecha(str(celda or ""))
                                if f:
                                    resultado["fecha_inicio_estudios"] = f
                                    resultado["fecha_inicio_estudios_texto"] = fecha_a_texto(f)

        resultado["raw_text"] = texto_completo
        return resultado

    @classmethod
    def _filtrar_bloque(cls, texto: str, semestre: str, ies: str) -> str:
        """Intenta aislar el bloque del semestre e IES correctos."""
        if not semestre and not ies:
            return texto

        if semestre:
            pos = texto.upper().find(semestre.upper())
            if pos != -1:
                return texto[pos: pos + 4000]

        return texto

    @classmethod
    def _extraer_primera_fecha(cls, texto: str) -> date | None:
        for m in cls.PATRON_FECHA.finditer(texto):
            # Formato largo: 15 de julio de 2026
            if m.group(1):
                dia, mes_str, anio = int(m.group(1)), m.group(2).lower(), int(m.group(3))
                mes = MESES_NUM.get(mes_str)
                if mes:
                    try:
                        return date(anio, mes, dia)
                    except ValueError:
                        pass
            # Formato corto: 15/07/2026
            elif m.group(4):
                try:
                    return date(int(m.group(6)), int(m.group(5)), int(m.group(4)))
                except ValueError:
                    pass
        return None


# ============================================================
# EXTRACTOR DOCUMENTO IES (EXCEL)
# ============================================================
class ExtractorDocumentoIESExcel:
    """Extrae cursos pendientes de un archivo Excel de la IES validando DNI y nombres."""
    
    PATRON_CODIGO = re.compile(r"\b([A-Z0-9]{2,}[\.\-][A-Z0-9\.\-]{2,}[\.\-][A-Z0-9\.\-]{2,})\b", re.IGNORECASE)
    PATRON_CODIGO_ALT = re.compile(r"(CARTA|OFICIO|MEMORANDO|CONSTANCIA|NOTA|COMUNICADO)\s+N[°ºo\.]?\s*([\d\w/\.\-]+)", re.IGNORECASE)
    PATRON_FECHA = re.compile(r"(\d{1,2})\s+de\s+([a-zñáéíóú]+)(?:\s+de(?:l)?)?\s+(\d{4})", re.IGNORECASE)

    @classmethod
    def extraer(
        cls,
        ruta_excel: str | Path,
        dni_validado: str,
        nombres_y_apellidos: str,
        log: LogCallback | None = None
    ) -> dict:
        _log = log or (lambda msg: None)
        ruta = Path(ruta_excel)
        resultado = {
            "codigo_doc_ies": "Documento",
            "fecha_doc_ies": None,
            "fecha_doc_ies_texto": "",
            "cursos_pendientes": [],
            "cursos_pendientes_texto": "",
            "ies_validada": True,
        }

        try:
            excel_file = pd.ExcelFile(ruta, engine="openpyxl")
            sheet_dfs = []
            texto_primera_linea = ""
            
            for sheet_name in excel_file.sheet_names:
                raw_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, dtype=str)
                if raw_df.empty:
                    continue
                    
                if not texto_primera_linea:
                    # Guardar la primera línea que contenga texto para buscar código y fecha
                    texto_primera_linea = " ".join(str(v).strip() for v in raw_df.iloc[0].values if pd.notna(v) and str(v).strip())
                header_idx = None
                for idx, row in raw_df.iloc[:20].iterrows():
                    row_str = " ".join(str(val or "").upper() for val in row.values)
                    if "DNI" in row_str or "CURSOS PENDIENTES" in row_str:
                        header_idx = idx
                        break
                
                if header_idx is not None:
                    headers = [str(val or "").strip() for val in raw_df.iloc[header_idx].values]
                    df_sheet = raw_df.iloc[header_idx + 1:].copy()
                else:
                    df_sheet = raw_df.copy()
                    headers = [str(c).strip() for c in df_sheet.iloc[0].values]
                    df_sheet = df_sheet.iloc[1:]
                
                # Desduplicar headers para evitar error de Reindexing en pd.concat
                seen = {}
                unique_headers = []
                for h in headers:
                    if h in seen:
                        seen[h] += 1
                        unique_headers.append(f"{h}_{seen[h]}")
                    else:
                        seen[h] = 0
                        unique_headers.append(h)
                
                df_sheet.columns = unique_headers
                sheet_dfs.append(df_sheet)
                    
            if not sheet_dfs:
                _log("No se pudo extraer información del Excel de la IES.")
                return resultado
                
            # Procesar documento y fecha de la primera línea
            if texto_primera_linea:
                m_alt = cls.PATRON_CODIGO_ALT.search(texto_primera_linea)
                if m_alt:
                    tipo_doc = m_alt.group(1).strip().capitalize()
                    codigo = m_alt.group(2).strip()
                    resultado["codigo_doc_ies"] = f"{tipo_doc} N\u00b0 {codigo}"
                else:
                    m = cls.PATRON_CODIGO.search(texto_primera_linea)
                    if m:
                        resultado["codigo_doc_ies"] = m.group(1)
                
                m_f = cls.PATRON_FECHA.search(texto_primera_linea)
                if m_f:
                    fecha = texto_a_fecha(m_f.group(0))
                    if fecha:
                        resultado["fecha_doc_ies"] = fecha
                        resultado["fecha_doc_ies_texto"] = fecha_a_texto(fecha)
                if not resultado["fecha_doc_ies"]:
                    m_corto = re.search(r"(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})", texto_primera_linea)
                    if m_corto:
                        fecha = texto_a_fecha(m_corto.group(0))
                        if fecha:
                            resultado["fecha_doc_ies"] = fecha
                            resultado["fecha_doc_ies_texto"] = fecha_a_texto(fecha)
                
            df_full = pd.concat(sheet_dfs, ignore_index=True).fillna("")
            df_full.columns = [str(c).strip().upper() for c in df_full.columns]
            
            col_dni = next((c for c in df_full.columns if "DNI" in c or "DOCUMENTO" in c), None)
            col_nom = next((c for c in df_full.columns if "APELLIDOS Y NOMBRES" in c or "BECARIO" in c or "NOMBRES Y APELLIDOS" in c or "NOMBRES" in c or "APELLIDOS" in c), None)
            
            # Lógica heurística para encontrar la columna de Cursos Pendientes
            col_cursos = None
            mejor_puntaje = 0
            
            for c in df_full.columns:
                t = str(c).lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
                puntaje = 0
                
                # Grupo A (Sustantivo: +3)
                if any(p in t for p in ["curso", "asignatura", "materia"]): puntaje += 3
                # Grupo B (Condición: +4)
                if any(p in t for p in ["pendiente", "faltante", "aprobar", "llevar", "restante", "concluir"]): puntaje += 4
                # Grupo C (Contexto: +1)
                if any(p in t for p in ["nombre", "detalle", "lista", "descripcion"]): puntaje += 1
                # Grupo D (Excluyentes: -10)
                if any(p in t for p in ["aprobado", "concluido", "historico", "llevado", "n°", "numero", "cantidad", "total"]): puntaje -= 10
                    
                if puntaje > mejor_puntaje:
                    mejor_puntaje = puntaje
                    col_cursos = c
                    
            # Umbral mínimo
            if mejor_puntaje < 6:
                col_cursos = None
            else:
                _log(f"Columna de cursos heurística: '{col_cursos}' (Puntaje: {mejor_puntaje})")
            
            if not col_dni or not col_cursos:
                _log("No se encontraron las columnas necesarias (DNI o Cursos Pendientes) en el Excel de la IES.")
                return resultado
                
            def _norm(texto: str) -> str:
                t = str(texto).upper()
                repls = {"Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U", "Ñ": "N", "Ü": "U"}
                for k, v in repls.items():
                    t = t.replace(k, v)
                return re.sub(r"[^A-Z0-9]+", " ", t).strip()
            
            nombre_target = _norm(nombres_y_apellidos)
            tokens_target = [t for t in nombre_target.split() if len(t) > 2]
            dni_clean = limpiar_dni(dni_validado)
            
            mejor_fila = None
            max_matches = 0
            
            for idx, row in df_full.iterrows():
                row_dni = limpiar_dni(row[col_dni])
                if row_dni == dni_clean:
                    if col_nom:
                        row_nom = _norm(row[col_nom])
                        matches = sum(1 for t in tokens_target if t in row_nom)
                        if matches > max_matches or mejor_fila is None:
                            max_matches = matches
                            mejor_fila = row
                    else:
                        mejor_fila = row
                        break
                        
            if mejor_fila is not None:
                if max_matches < 2 and col_nom:
                    _log(f"ADVERTENCIA: DNI encontrado en IES, pero los nombres no coinciden lo suficiente. (Coincidencias: {max_matches})")
                else:
                    _log("Becario validado correctamente en el Excel de la IES.")
                    
                cursos_raw = str(mejor_fila[col_cursos]).strip()
                if cursos_raw:
                    _log(f"Cursos extraídos sin procesar: {cursos_raw}")
                    # Separar casos que no tengan saltos de línea evidentes
                    cursos_raw = re.sub(r'(\w+)\s+(Electivo)\b', r'\1\n\2', cursos_raw, flags=re.IGNORECASE)
                    cursos_raw = re.sub(r'\b(Electivo)\s+(Electivo)\b', r'\1\n\2', cursos_raw, flags=re.IGNORECASE)
                    cursos_raw = re.sub(r'\b(Electivo)\s+(Electivo)\b', r'\1\n\2', cursos_raw, flags=re.IGNORECASE)
                    cursos_raw = re.sub(r'\b([IVX]+)\s+([A-Z])', r'\1\n\2', cursos_raw)
                    
                    # Dividir usando múltiples separadores posibles: salto de línea, coma, guión (con espacios), o punto y coma
                    separadores = r'\n|;|,(?!\s*\d)|\s+[-–—]\s+'
                    lista_cursos = [c.strip() for c in re.split(separadores, cursos_raw) if c.strip()]
                    _log(f"Cursos separados: {lista_cursos}")
                        
                    cursos_limpios = []
                    for c in lista_cursos:
                        c = re.sub(r"^[\d]+[\.\-\)\s]+\s*", "", c).strip()
                        c = re.sub(r"^[-•·]\s*", "", c).strip()
                        if len(c) > 3:
                            cursos_limpios.append(formatear_curso_oracion(c))
                            
                    if cursos_limpios:
                        resultado["cursos_pendientes"] = cursos_limpios
                        resultado["cursos_pendientes_texto"] = "\n".join(
                            f"{i+1}. {c}" for i, c in enumerate(cursos_limpios)
                        )
            else:
                _log("No se encontró el DNI en el Excel de la IES.")
                raise BecarioNoEncontradoIESException("el becario no se encuentra en el documento de la IES, verificar si cumple el requisito de culminar los cursos en el 2026-II")
                
        except BecarioNoEncontradoIESException:
            raise
        except Exception as e:
            _log(f"Error procesando Excel IES: {e}")
            
        return resultado


# ============================================================
# EXTRACTOR DOCUMENTO 4: Documento de la IES
# ============================================================
class ExtractorDocumentoIES:
    """Extrae código, fecha y lista de cursos pendientes del Documento de la IES."""

    PATRON_CODIGO = re.compile(
        r"\b([A-Z0-9]{2,}[\.\-][A-Z0-9\.\-]{2,}[\.\-][A-Z0-9\.\-]{2,})\b",
        re.IGNORECASE,
    )
    # Captura también el TIPO de documento (CARTA, OFICIO, etc.) y el número
    PATRON_CODIGO_ALT = re.compile(
        r"(CARTA|OFICIO|MEMORANDO|CONSTANCIA|NOTA|COMUNICADO)\s+N[°ºo\.]?\s*([\d\w/\.\-]+)",
        re.IGNORECASE,
    )
    PATRON_FECHA = re.compile(
        r"(\d{1,2})\s+de\s+([a-zñáéíóú]+)(?:\s+de(?:l)?)?\s+(\d{4})",
        re.IGNORECASE,
    )
    PATRON_SECCION_CURSOS = re.compile(
        r"(curso[s]?\s+pendiente[s]?|asignatura[s]?\s+pendiente[s]?|materia[s]?\s+pendiente[s]?)",
        re.IGNORECASE,
    )
    PATRON_SECCION_NOMBRE_CURSOS = re.compile(
        r"^\s*Cursos?\s*:?\s*$",
        re.IGNORECASE | re.MULTILINE,
    )
    PATRON_ITEM_LISTA = re.compile(
        r"^[\s]*(?:\d+[\.\)]\s*|[-•·]\s*)(.+)$",
        re.MULTILINE,
    )
    # Detecta líneas que describen créditos (NO son nombres de cursos)
    PATRON_CREDITO = re.compile(
        r"\bcrédito|credito|crédit\b",
        re.IGNORECASE,
    )

    @classmethod
    def extraer(cls, ruta_pdf: str | Path, ies_esperada: str = "") -> dict:
        ruta = Path(ruta_pdf)
        resultado = {
            "codigo_doc_ies": "",
            "fecha_doc_ies": None,
            "fecha_doc_ies_texto": "",
            "cursos_pendientes": [],
            "cursos_pendientes_texto": "",
            "ies_validada": False,
        }

        with pdfplumber.open(ruta) as pdf:
            textos = []
            for p in pdf.pages:
                try:
                    textos.append(p.extract_text() or "")
                except Exception:
                    pass
            texto_completo = "\n".join(textos)

        # Código del documento — captura tipo (Carta/Oficio) + código
        m_alt = cls.PATRON_CODIGO_ALT.search(texto_completo)
        if m_alt:
            tipo_doc = m_alt.group(1).strip().capitalize()   # 'Carta', 'Oficio', etc.
            codigo = m_alt.group(2).strip()
            resultado["codigo_doc_ies"] = f"{tipo_doc} N\u00b0 {codigo}"
        else:
            m = cls.PATRON_CODIGO.search(texto_completo)
            if m:
                resultado["codigo_doc_ies"] = m.group(1)

        # Fecha del documento
        m = cls.PATRON_FECHA.search(texto_completo)
        if m:
            fecha = texto_a_fecha(m.group(0))
            if fecha:
                resultado["fecha_doc_ies"] = fecha
                resultado["fecha_doc_ies_texto"] = fecha_a_texto(fecha)
        if not resultado["fecha_doc_ies"]:
            m_corto = re.search(r"(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})", texto_completo)
            if m_corto:
                fecha = texto_a_fecha(m_corto.group(0))
                if fecha:
                    resultado["fecha_doc_ies"] = fecha
                    resultado["fecha_doc_ies_texto"] = fecha_a_texto(fecha)

        # Validar IES
        if ies_esperada:
            ies_clean = re.sub(r"-\s*sede.*", "", ies_esperada.lower(), flags=re.IGNORECASE).strip()
            tokens_ies = [t for t in re.sub(r"[^\w\s]", "", ies_clean).split() if len(t) > 3 and t not in ("universidad", "instituto", "escuela", "sede")]
            if tokens_ies:
                resultado["ies_validada"] = any(t in texto_completo.lower() for t in tokens_ies)
            else:
                resultado["ies_validada"] = ies_clean in texto_completo.lower()
        else:
            resultado["ies_validada"] = True

        # RUTA 1: Formato tabla "Nº CURSO CRÉDITO" (ej: "1 TEORÍA DE LAS RELACIONES 4")
        patron_tabla_ies = re.compile(
            r"N[°ºo]?\s*CURSO\s+CR[ÉE]DITO|CURSO\s+CR[ÉE]DITO|ASIGNATURA\s+CR[ÉE]DITO",
            re.IGNORECASE,
        )
        m_tabla = patron_tabla_ies.search(texto_completo)
        if m_tabla:
            zona_tabla = texto_completo[m_tabla.end(): m_tabla.end() + 800]
            # Filas: "1 NOMBRE DEL CURSO 4" o "1. NOMBRE DEL CURSO 4"
            pat_fila_tabla = re.compile(
                r"^(\d+)[\.\s]\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑA-Za-záéíóúñ\s\(\)\/\-]{4,80?}?)\s+\d{1,3}\s*$",
                re.MULTILINE,
            )
            cursos_tabla = []
            for m_f in pat_fila_tabla.finditer(zona_tabla):
                nombre = m_f.group(2).strip()
                # Filtrar líneas de créditos o encabezados
                if nombre and len(nombre) > 3 and not re.search(r"cr[eé]dito|total|subtotal", nombre, re.IGNORECASE):
                    nombre_fmt = formatear_curso_oracion(nombre)
                    if nombre_fmt not in cursos_tabla:
                        cursos_tabla.append(nombre_fmt)
            if cursos_tabla:
                resultado["cursos_pendientes"] = cursos_tabla
                resultado["cursos_pendientes_texto"] = "\n".join(
                    f"{i+1}. {c}" for i, c in enumerate(cursos_tabla)
                )
                resultado["raw_text"] = texto_completo
                return resultado

        # RUTA 2: Formato lista clásica con viñeta/numeración "1. Curso" o "• Curso"
        m_seccion = cls.PATRON_SECCION_CURSOS.search(texto_completo)
        if m_seccion:
            seccion = texto_completo[m_seccion.start():]
            # Buscar subsección titulada "Cursos" con los nombres reales
            m_nombre_cursos = cls.PATRON_SECCION_NOMBRE_CURSOS.search(seccion)
            if m_nombre_cursos:
                seccion = seccion[m_nombre_cursos.start():]
            bloques = re.split(r"\n\s*\n", seccion)
            bloque_cursos = bloques[0] + ("\n" + bloques[1] if len(bloques) > 1 else "")
            # Patron extendido: acepta "1. Curso", "1) Curso" Y también "1 Curso" (sin punto)
            patron_item_ext = re.compile(
                r"^[\s]*(?:\d+[\.\)\s]\s*|[-•·]\s*)([A-ZÁÉÍÓÚÑA-Za-záéíóúñ].+)$",
                re.MULTILINE,
            )
            cursos = []
            for m_item in patron_item_ext.finditer(bloque_cursos):
                curso = m_item.group(1).strip()
                # Quitar crédito al final (número suelto al final de la línea)
                curso = re.sub(r"\s+\d{1,3}\s*$", "", curso).strip()
                # Filtrar líneas que describen créditos (no son nombres de cursos)
                if cls.PATRON_CREDITO.search(curso):
                    continue
                if len(curso) > 3:
                    curso_fmt = formatear_curso_oracion(curso)
                    if curso_fmt not in cursos:
                        cursos.append(curso_fmt)
            resultado["cursos_pendientes"] = cursos
            resultado["cursos_pendientes_texto"] = "\n".join(
                f"{i+1}. {c}" for i, c in enumerate(cursos)
            )

        resultado["raw_text"] = texto_completo
        return resultado


# ============================================================
# EXTRACTOR PADRON EXCEL (Base de Datos / Ground Truth)
# ============================================================
class ExtractorPadron:
    """Carga el padrón Excel inspeccionando todas las pestañas y encabezados con consolidación completa de nombres."""

    STOP_WORDS = {
        "DE", "DEL", "LA", "LAS", "LOS", "EL", "SAN", "SANTA", "Y", "VDA", "VDA.",
        "SUBDIRECCIÓN", "SUBDIRECCION", "COORDINACIÓN", "COORDINACION", "REGIONAL",
        "UNIDAD", "OFICINA", "DIRECCIÓN", "DIRECCION", "LIMA", "PRONABEC", "CARTA",
        "OFICIO", "ANEXO", "INFORME", "FORMATO", "DOCUMENTO", "PDF", "MINEDU",
        "BECA", "BECAS", "UNIVERSIDAD", "INSTITUTO", "ESCUELA", "CARRERA", "MODALIDAD",
        "CONVOCATORIA", "SEMESTRE", "PERIODO", "AÑO", "ANIO", "2026", "2025", "2024",
        "2023", "2022", "2021", "2020", "VALOR", "PAGINA", "REGISTRO", "ESTADO",
        "OBSERVADO", "PROCEDENTE", "ACTIVO", "INACTIVO", "N°", "Nº", "NUMERO", "NUM"
    }

    def __init__(self, ruta_excel: str | Path) -> None:
        self.ruta = Path(ruta_excel)
        self._df: pd.DataFrame | None = None

    def cargar(self) -> None:
        """Carga de forma inteligente todas las hojas del Excel escaneando filas de encabezados."""
        try:
            excel_file = pd.ExcelFile(self.ruta, engine="openpyxl")
            sheet_dfs = []
            for sheet_name in excel_file.sheet_names:
                raw_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, dtype=str)
                if raw_df.empty:
                    continue

                header_idx = None
                for idx, row in raw_df.iloc[:20].iterrows():
                    row_str = " ".join(str(val or "").upper() for val in row.values)
                    if any(w in row_str for w in ("DNI", "BECARIO", "NOMBRES", "APELLIDOS", "DOCUMENTO", "ESTUDIANTE", "PROGRAMA", "BENEFICIARIO")):
                        header_idx = idx
                        break

                if header_idx is not None:
                    headers = [str(val or "").strip() for val in raw_df.iloc[header_idx].values]
                    df_sheet = raw_df.iloc[header_idx + 1:].copy()
                    df_sheet.columns = headers
                else:
                    df_sheet = raw_df.copy()
                    df_sheet.columns = [str(c).strip() for c in df_sheet.iloc[0].values]
                    df_sheet = df_sheet.iloc[1:]

                df_sheet = df_sheet.fillna("")
                df_sheet = df_sheet[df_sheet.apply(lambda r: "".join(str(v) for v in r.values).strip() != "", axis=1)]
                sheet_dfs.append(df_sheet)

            if sheet_dfs:
                self._df = pd.concat(sheet_dfs, ignore_index=True)
            else:
                self._df = pd.DataFrame()
        except Exception:
            self._df = pd.read_excel(self.ruta, engine="openpyxl", dtype=str).fillna("")

        if self._df is not None:
            self._df.columns = [str(c).strip() for c in self._df.columns]

    def _norm(self, texto: str) -> str:
        """Normaliza texto removiendo acentos, puntuación y convirtiendo a mayúsculas."""
        t = str(texto).upper()
        repls = {"Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U", "Ñ": "N", "Ü": "U"}
        for k, v in repls.items():
            t = t.replace(k, v)
        t = re.sub(r"[^A-Z0-9]+", " ", t)
        return t.strip()

    def _col(self, tipo: str) -> str | None:
        """Devuelve el nombre real de la columna del DataFrame según el tipo solicitado."""
        if self._df is None or self._df.empty:
            return None

        cols_norm = {c: self._norm(c) for c in self._df.columns}

        if tipo == "dni":
            for orig, n in cols_norm.items():
                words = n.split()
                if "DNI" in words or "CEDULA" in words or "DOCUMENTO" in words or "NUM_DNI" in n or "NRO_DNI" in n:
                    if not any(w in n for w in ("FECHA", "TIPO", "DOC_IES", "ACTIVIDAD", "UNIDAD")):
                        return orig

        elif tipo == "nombre":
            for orig, n in cols_norm.items():
                if any(w in n for w in ("BECARIO", "NOMBRES", "APELLIDOS", "ESTUDIANTE", "ALUMNO", "BENEFICIARIO", "POSTULANTE")):
                    return orig

        elif tipo == "expediente":
            for orig, n in cols_norm.items():
                if "EXPEDIENTE" in n or "NEPEDIENTE" in n:
                    return orig

        return None

    def buscar_becario(
        self,
        dni_succor: str = "",
        nombre_succor: str = "",
        texto_comb: str = "",
        log: Callable[[str], None] | None = None,
    ) -> dict | None:
        """
        Búsqueda multinivel estricta (exige coincidencia de al menos 2 tokens de nombre de estudiante):
        Nivel 1: DNI directo del SUCCOR.
        Nivel 2: DNI del padrón presente en el texto de los documentos/archivos.
        Nivel 3: Nombre aproximado del SUCCOR (>= 2 palabras clave coincidentes).
        Nivel 4: Coincidencia cruzada de 2 o más palabras del nombre en los documentos/nombres de archivo.
        Nivel 5: Búsqueda estricta por tokens de nombre de estudiante en cualquier columna.
        """
        _log = log or (lambda msg: None)
        if self._df is None or self._df.empty:
            return None

        col_dni = self._col("dni")
        col_nombre = self._col("nombre")

        # --- Nivel 1: DNI directo del SUCCOR ---
        if dni_succor and col_dni:
            dni_limpio = limpiar_dni(dni_succor)
            if dni_limpio:
                for idx, val in self._df[col_dni].items():
                    cell_dni = limpiar_dni(val)
                    if cell_dni and (cell_dni == dni_limpio or cell_dni.endswith(dni_limpio) or dni_limpio.endswith(cell_dni)):
                        _log(f"  [Padrón Nivel 1] Becario hallado por DNI directo: {dni_limpio}")
                        return self._df.loc[idx].to_dict()

        # --- Nivel 2: DNI del padrón presente en cualquier texto de los PDFs / nombres de archivo ---
        if texto_comb and col_dni:
            dnis_en_texto = set(re.findall(r"\b(\d{7,8})\b", texto_comb))
            if dnis_en_texto:
                for idx, val in self._df[col_dni].items():
                    cell_dni = limpiar_dni(val)
                    if cell_dni and cell_dni in dnis_en_texto:
                        _log(f"  [Padrón Nivel 2] Becario hallado por DNI coincidente en documentos: {cell_dni}")
                        return self._df.loc[idx].to_dict()

        # --- Nivel 3: Nombre aproximado del SUCCOR ---
        if nombre_succor and col_nombre:
            nombre_clean = self._norm(nombre_succor)
            tokens_target = [t for t in nombre_clean.split() if len(t) > 2 and t not in self.STOP_WORDS]
            if len(tokens_target) >= 2:
                mejor_fila = None
                max_matches = 0
                for idx, val in self._df[col_nombre].items():
                    cell_norm = self._norm(str(val))
                    matches = sum(1 for t in tokens_target if t in cell_norm)
                    if matches > max_matches:
                        max_matches = matches
                        mejor_fila = self._df.loc[idx].to_dict()
                if mejor_fila and max_matches >= 2:
                    _log(f"  [Padrón Nivel 3] Becario hallado por nombre del informe: {nombre_succor}")
                    return mejor_fila

        # --- Nivel 4: Tokens de nombres/apellidos presentes en texto/nombres de archivo ---
        if texto_comb:
            texto_comb_norm = self._norm(texto_comb)
            mejor_fila = None
            max_matches = 0

            cols_nom_list = [c for c in self._df.columns if any(w in self._norm(c) for w in ("BECARIO", "NOMBRE", "APELLIDO", "ESTUDIANTE", "ALUMNO", "BENEFICIARIO", "POSTULANTE"))]
            if not cols_nom_list and col_nombre:
                cols_nom_list = [col_nombre]

            for idx, row in self._df.iterrows():
                nombres_raw = " ".join(str(row[c]) for c in cols_nom_list if c in row)
                nombre_norm = self._norm(nombres_raw)
                tokens = [
                    t for t in nombre_norm.split()
                    if len(t) > 2 and t not in self.STOP_WORDS
                ]
                if len(tokens) < 2:
                    continue

                matches = sum(1 for t in tokens if t in texto_comb_norm)
                if matches > max_matches:
                    max_matches = matches
                    mejor_fila = row.to_dict()

            if mejor_fila and max_matches >= 2:
                _log(f"  [Padrón Nivel 4] Becario hallado por coincidencia de {max_matches} nombres/apellidos en los documentos.")
                return mejor_fila

        # --- Nivel 5: Búsqueda estricta en cualquier columna de la fila Excel (exige >= 2 nombres reales) ---
        if texto_comb:
            texto_comb_norm = self._norm(texto_comb)
            mejor_fila = None
            max_matches = 0

            for idx, row in self._df.iterrows():
                fila_norm = self._norm(" ".join(str(v) for v in row.values))
                tokens = [t for t in fila_norm.split() if len(t) > 2 and t not in self.STOP_WORDS and not t.isdigit()]
                if len(tokens) < 2:
                    continue
                matches = sum(1 for t in tokens if t in texto_comb_norm)
                if matches > max_matches and matches >= 2:
                    max_matches = matches
                    mejor_fila = row.to_dict()

            if mejor_fila:
                _log(f"  [Padrón Nivel 5] Becario hallado por coincidencia estricta de {max_matches} palabras en la fila.")
                return mejor_fila

        return None

    def obtener_expediente_vigente(self, dni: str) -> str:
        """Busca el expediente (NEXPEDIENTE). Si hay fechas, intenta que sea vigente."""
        if self._df is None or self._df.empty or not dni:
            return ""

        col_dni = self._col("dni")
        col_exp = self._col("expediente")
        if not col_dni or not col_exp:
            return ""

        dni_limpio = limpiar_dni(dni)
        hoy = date.today()
        
        exp_fallback = ""

        for idx, row in self._df.iterrows():
            row_dni = limpiar_dni(row[col_dni])
            if row_dni == dni_limpio:
                cols_norm = {self._norm(k): k for k in row.keys()}
                def _get_val(*palabras_clave):
                    for orig, norm_col in cols_norm.items():
                        if any(w in orig for w in palabras_clave):
                            val = str(row[norm_col]).strip()
                            if val and val.lower() != "nan":
                                return val
                    return ""

                fecha_ini_str = _get_val("INICIO", "FINICIO")
                fecha_fin_str = _get_val("FIN", "FFIN", "TERMINO")

                fi = texto_a_fecha(fecha_ini_str)
                ff = texto_a_fecha(fecha_fin_str)
                
                exp_actual = str(row[col_exp]).strip()
                if exp_actual and exp_actual.lower() != "nan":
                    exp_fallback = exp_actual

                if fi and ff and fi <= hoy <= ff:
                    return exp_actual

        return exp_fallback

    def extraer_datos(self, fila: dict) -> dict:
        """A partir de una fila del DataFrame extrae y consolida los datos normalizados del becario uniendo Apellidos y Nombres."""
        cols_norm = {self._norm(k): k for k in fila.keys()}

        def _get_val(*palabras_clave):
            for orig, norm_col in cols_norm.items():
                if any(w in orig for w in palabras_clave):
                    val = str(fila[norm_col]).strip()
                    if val and val.lower() != "nan":
                        return val
            return ""

        # DNI
        dni = limpiar_dni(_get_val("DNI", "CEDULA", "DOCUMENTO", "IDENTIDAD"))

        # Consolidar apellidos (Paterno + Materno o Apellidos generales) y nombres
        paterno = _get_val("APELLIDO PATERNO", "PATERN", "PRIMER APELLIDO")
        materno = _get_val("APELLIDO MATERNO", "MATERN", "SEGUNDO APELLIDO")
        apellidos_gen = _get_val("APELLIDOS", "APELLIDO")
        nombres_gen = _get_val("NOMBRES", "NOMBRE")
        becario_gen = _get_val("BECARIO", "ESTUDIANTE", "ALUMNO", "BENEFICIARIO", "POSTULANTE", "NOMBRE COMPLETO")

        # Construir apellidos y nombres por separado
        apellidos_partes = []
        if paterno:
            apellidos_partes.append(paterno)
        if materno:
            apellidos_partes.append(materno)
        if not apellidos_partes and apellidos_gen:
            apellidos_partes.append(apellidos_gen)

        apellidos_str = " ".join(apellidos_partes).strip()
        nombres_str = nombres_gen.strip() if nombres_gen else ""

        partes_nombre = []
        if apellidos_partes:
            partes_nombre.extend(apellidos_partes)

        if nombres_str and nombres_str not in partes_nombre:
            partes_nombre.append(nombres_str)

        if partes_nombre:
            nombre_completo = " ".join(partes_nombre).strip()
        elif becario_gen:
            nombre_completo = becario_gen
            # Intentar separar apellidos/nombres del campo combinado:
            # se asume que las últimas 2 palabras son el nombre propio
            tokens_comb = becario_gen.split()
            if len(tokens_comb) >= 4:
                apellidos_str = " ".join(tokens_comb[:2])
                nombres_str = " ".join(tokens_comb[2:])
            else:
                apellidos_str = becario_gen
                nombres_str = ""
        else:
            # Extraer cadenas de texto legibles de la fila
            tokens_vali = [
                str(v).strip() for k, v in fila.items()
                if str(v).strip() and len(str(v).strip()) > 3 and self._norm(str(v)) not in self.STOP_WORDS
            ]
            nombre_completo = " ".join(tokens_vali[:4]) if tokens_vali else "(becario detectado)"
            apellidos_str = nombre_completo
            nombres_str = ""

        programa = _get_val("PROGRAMA", "BECA", "MODALIDAD")
        convocatoria = _get_val("CONVOCATORIA", "ANIO", "AÑO", "PERIODO")

        sexo_raw = _get_val("SEXO", "GENERO").upper()
        sexo = "F" if sexo_raw.startswith("F") else "M"

        fecha_ini = _get_val("INICIO", "FINICIO")
        fecha_fin = _get_val("FIN", "FFIN", "TERMINO")

        return {
            "dni_validado": dni,
            "nombres_y_apellidos_validados": nombre_completo,
            "apellidos_becario": apellidos_str,
            "nombres_becario": nombres_str,
            "programa_beca": programa,
            "convocatoria": convocatoria,
            "sexo": sexo,
            "fecha_inicio_sibec_raw": fecha_ini,
            "fecha_fin_sibec_raw": fecha_fin,
        }


# ============================================================
# GENERADOR DE CONCORDANCIAS DE GÉNERO
# ============================================================
def generar_concordancias_genero(sexo: str) -> dict:
    """Genera variables de concordancia gramatical según el género."""
    es_femenino = str(sexo).strip().upper().startswith("F")
    if es_femenino:
        return {
            "sexo_articulo_1": "1 becaria",
            "sexo_articulo_2": "la becaria",
            "sexo_conector": "presentada por una becaria",
            "sexo_becario_a": "becaria",
            "sexo_nominado_a": "nominada",
            "sexo_interesado_a": "interesada",
            "el_la_becario_a": "la becaria",
            "del_becario_a": "de la becaria",
            "a_la_becario_a": "a la becaria",
        }
    else:
        return {
            "sexo_articulo_1": "1 becario",
            "sexo_articulo_2": "el becario",
            "sexo_conector": "presentado por un becario",
            "sexo_becario_a": "becario",
            "sexo_nominado_a": "nominado",
            "sexo_interesado_a": "interesado",
            "el_la_becario_a": "el becario",
            "del_becario_a": "del becario",
            "a_la_becario_a": "al becario",
        }



# ============================================================
# FUNCIONES UTILITARIAS DE LIMPIEZA DE DATOS
# ============================================================

def limpiar_nombre_ies(nombre: str) -> str:
    """Elimina sufijos de sede de la institución. Ej: 'Univ. X / Sede Lima' -> 'Univ. X'."""
    if not nombre:
        return nombre
    # Remover ' / Sede ...', '. / Sede ...', ' - Sede ...', '- Sede ...' y variantes
    nombre = re.sub(r"[\s\.]*[/\-]\s*Sede\s+.+", "", nombre, flags=re.IGNORECASE)
    # Remover ' Sede ...' sin separador
    nombre = re.sub(r"\s+Sede\s+.+", "", nombre, flags=re.IGNORECASE)
    return nombre.strip().strip(".")


def limpiar_nombre_informe(nombre: str) -> str:
    """Trunca el nombre del informe en 'LIMA' y capitaliza 'Informe'. 
    Ej: 'INFORME N° 4291-2026-MINEDU/VMGI-PRONABEC-DICONCI-SUCCOR-LIMA A' 
        -> 'Informe N° 4291-2026-MINEDU/VMGI-PRONABEC-DICONCI-SUCCOR-LIMA'"""
    if not nombre:
        return nombre
    # Truncar en "-LIMA" incluyendo la palabra LIMA pero excluyendo lo que sigue
    m = re.search(r"(-LIMA)\b", nombre, re.IGNORECASE)
    if m:
        nombre = nombre[:m.end()]
    # Capitalizar solo la primera letra: INFORME -> Informe
    nombre = nombre.strip()
    if nombre.upper().startswith("INFORME"):
        nombre = "Informe" + nombre[7:]
    return nombre


def limpiar_beca_convocatoria(programa: str, convocatoria: str) -> str:
    """Genera texto limpio de beca y año. Ej: 'BECA 18 BECA REPARED', '2021' -> 'BECA 18 - 2021'."""
    if not programa:
        return f"{convocatoria}".strip() if convocatoria else ""
    prog = programa.strip()
    # Mantener solo el inicio tipo 'BECA 18', 'BECA BICENTENARIO', etc.
    # Eliminar sub-modalidades redundantes (BECA REPARED, BECA PREPAGO, BECA INTEGRA, etc.)
    m = re.match(r"(BECA\s+\w+)", prog, re.IGNORECASE)
    nombre_beca = m.group(1).strip() if m else prog
    anio = re.search(r"\b(20\d{2})\b", convocatoria or "") or re.search(r"\b(20\d{2})\b", programa or "")
    anio_str = anio.group(1) if anio else convocatoria
    if anio_str:
        return f"{nombre_beca} - {anio_str}"
    return nombre_beca


def nombre_en_orden_nombres_apellidos(apellidos: str, nombres: str) -> str:
    """Convierte 'APELLIDO1 APELLIDO2' + 'NOMBRE1 NOMBRE2' a 'Nombre1 Nombre2 Apellido1 Apellido2' en Title Case."""
    partes = []
    if nombres:
        partes.append(nombres.title())
    if apellidos:
        partes.append(apellidos.title())
    if partes:
        return " ".join(partes)
    return ""


# ============================================================
# ORQUESTADOR PRINCIPAL
# ============================================================
class ProcesadorInformes:
    """Orquesta la extracción de los 4 documentos, cruce con padrón y preparación del contexto."""

    def __init__(
        self,
        ruta_excel: str | Path,
        ruta_formato_autogenerado: str | Path,
        ruta_informe_succor: str | Path,
        ruta_calendario_academico: str | Path,
        ruta_documento_ies: str | Path,
        log: LogCallback | None = None,
        progreso: ProgressCallback | None = None,
        nro_informe: str = "",
        rutas_formatos: list = None,
    ) -> None:
        self.ruta_excel = Path(ruta_excel)
        self.ruta_formato_autogenerado = Path(ruta_formato_autogenerado)
        self.ruta_informe_succor = Path(ruta_informe_succor)
        self.ruta_calendario_academico = Path(ruta_calendario_academico)
        self.ruta_documento_ies = Path(ruta_documento_ies)
        self._log = log or (lambda msg: None)
        self._progreso = progreso or (lambda pct, msg: None)
        self.nro_informe = nro_informe

    def ejecutar(self) -> Path:
        self._validar_entradas()

        ctx = {}
        ctx["NUMERO_INFORME_GENERAR"] = self.nro_informe

        # Inicialización previa por seguridad
        datos_fmt: dict[str, Any] = {}
        datos_succor: dict[str, Any] = {}
        datos_cal: dict[str, Any] = {}
        datos_ies: dict[str, Any] = {}

        # --- Paso 1: Formato autogenerado ---
        self._progreso(0.05, "Extrayendo fecha de solicitud...")
        self._log("Doc.1: Leyendo formato autogenerado...")
        datos_fmt = ExtractorFormatoAutogenerado.extraer(self.ruta_formato_autogenerado)
        ctx["FECHA_SOLICITUD_TEXTO"] = datos_fmt.get("fecha_solicitud_texto", "(no detectada)")
        # REFERENCIA_A — literal a) con fecha de solicitud
        ctx["REFERENCIA_A"] = f"a) Solicitud ingresada por mesa de partes el {ctx['FECHA_SOLICITUD_TEXTO']}"
        self._log(f"  Fecha de solicitud: {ctx['FECHA_SOLICITUD_TEXTO']}")

        # --- Paso 2: Informe SUCCOR ---
        self._progreso(0.15, "Extrayendo Informe SUCCOR...")
        self._log("Doc.2: Leyendo Informe SUCCOR...")
        datos_succor = ExtractorInformeSuccor.extraer(self.ruta_informe_succor)
        # Limpiar nombre del informe: truncar en LIMA y capitalizar 'Informe'
        nombre_informe_raw = datos_succor.get("nombre_informe_succor", "")
        ctx["NOMBRE_INFORME_SUCCOR"] = limpiar_nombre_informe(nombre_informe_raw)
        ctx["NUMERO_INFORME"] = ctx["NOMBRE_INFORME_SUCCOR"]
        # REFERENCIA_B — literal b) con el informe SUCCOR
        ctx["REFERENCIA_B"] = f"b) {ctx['NOMBRE_INFORME_SUCCOR']}"
        # SIGEDO: preferir Expediente del formato autogenerado; fallback SUCCOR
        num_exp = datos_fmt.get("numero_expediente", "")
        if num_exp:
            ctx["NUMERO_SIGEDO"] = f"{num_exp}-{date.today().year}"
        else:
            ctx["NUMERO_SIGEDO"] = datos_succor.get("numero_sigedo", "")
        ctx["SEMESTRE_SOLICITADO"] = datos_succor.get("semestre_solicitado", "")
        ctx["RJD_ADJUDICACION"] = datos_succor.get("rjd_adjudicacion", "")
        # Limpiar IES: eliminar sufijo de sede
        ctx["INSTITUCION"] = limpiar_nombre_ies(datos_succor.get("institucion", ""))
        ctx["CARRERA"] = datos_succor.get("carrera", "")
        dni_succor = datos_succor.get("dni_succor", "")
        self._log(f"  Informe: {ctx['NOMBRE_INFORME_SUCCOR']}")
        self._log(f"  SIGEDO: {ctx['NUMERO_SIGEDO']}")
        self._log(f"  Semestre: {ctx['SEMESTRE_SOLICITADO']}")
        self._log(f"  DNI extraído del SUCCOR: {dni_succor or '(no detectado en SUCCOR)'}")

        # --- Paso 3: Calendario académico ---
        self._progreso(0.25, "Extrayendo calendario académico...")
        self._log("Doc.3: Leyendo calendario académico (todas las págs.)...")
        datos_cal = ExtractorCalendarioAcademico.extraer(
            self.ruta_calendario_academico,
            semestre_solicitado=ctx["SEMESTRE_SOLICITADO"],
            ies_filtro=ctx["INSTITUCION"],
        )
        ctx["FECHA_MATRICULA"] = datos_cal.get("fecha_matricula_texto", "(no detectada)")
        ctx["FECHA_INICIO_ESTUDIOS"] = datos_cal.get("fecha_inicio_estudios_texto", "(no detectada)")
        self._log(f"  Fecha matrícula: {ctx['FECHA_MATRICULA']}")
        self._log(f"  Fecha inicio estudios: {ctx['FECHA_INICIO_ESTUDIOS']}")

        # Evaluación de procedencia
        f_sol = datos_fmt.get("fecha_solicitud")
        f_mat = datos_cal.get("fecha_matricula")
        f_ini = datos_cal.get("fecha_inicio_estudios")
        if f_sol and (f_mat or f_ini):
            es_procedente = (f_mat and f_sol <= f_mat) or (f_ini and f_sol <= f_ini)
            ctx["ESTADO_PROCEDENCIA"] = "PROCEDENTE" if es_procedente else "OBSERVADO"
        else:
            ctx["ESTADO_PROCEDENCIA"] = "(no evaluado - fechas incompletas)"
            
        if f_sol and f_mat and f_sol >= f_mat:
            self._log(f"  [DEBUG] f_sol={f_sol} >= f_mat={f_mat}. Agregando advertencia.")
            if not hasattr(self, 'advertencias'):
                self.advertencias = []
            self.advertencias.append("La fecha del formato autogenerado es igual o mayor que la fecha de matricula. Revisar los documentos.")
        else:
            self._log(f"  [DEBUG] Condicion no cumplida: f_sol={f_sol}, f_mat={f_mat}")
        self._log(f"  Procedencia: {ctx['ESTADO_PROCEDENCIA']}")

        # --- Paso 4: Cruce con Padrón (Ground Truth) ---
        self._progreso(0.40, "Cruzando datos con el Padrón...")
        self._log("Módulo BD: Buscando becario en el padrón...")

        # Texto combinado de todos los PDFs y nombres de archivo (sin doc IES temporalmente para la búsqueda inicial)
        texto_comb = "\n".join([
            str(self.ruta_formato_autogenerado.name),
            str(self.ruta_informe_succor.name),
            str(self.ruta_calendario_academico.name),
            str(self.ruta_documento_ies.name),
            datos_fmt.get("raw_text", ""),
            datos_succor.get("raw_text", ""),
            datos_cal.get("raw_text", ""),
        ])

        padron = ExtractorPadron(self.ruta_excel)
        padron.cargar()

        fila_becario = padron.buscar_becario(
            dni_succor=dni_succor,
            nombre_succor=datos_succor.get("nombre_aproximado", ""),
            texto_comb=texto_comb,
            log=self._log,
        )

        if fila_becario:
            datos_bd = padron.extraer_datos(fila_becario)
            self._log(f"  Becario encontrado: {datos_bd['nombres_y_apellidos_validados']}")
            self._log(f"  Sexo (BD): {datos_bd['sexo']}")

            ctx["DNI_VALIDADO"] = datos_bd["dni_validado"] or dni_succor
            ctx["NOMBRES_Y_APELLIDOS_VALIDADOS"] = datos_bd["nombres_y_apellidos_validados"]

            # Limpiar BECA y CONVOCATORIA (dejar solo 'BECA X - AÑO')
            ctx["BECA_Y_CONVOCATORIA_VALIDADA"] = limpiar_beca_convocatoria(
                datos_bd["programa_beca"], datos_bd["convocatoria"]
            )
            # Versión Title Case para párrafos narrativos: 'Beca 18 - 2021'
            ctx["BECA_TITULO"] = ctx["BECA_Y_CONVOCATORIA_VALIDADA"].title()

            # Apellidos y nombres separados para la tabla del becario
            apellidos_bd = datos_bd.get("apellidos_becario", "")
            nombres_bd = datos_bd.get("nombres_becario", "")
            ctx["APELLIDOS_BECARIO"] = apellidos_bd.upper() if apellidos_bd else ""
            ctx["NOMBRES_BECARIO"] = nombres_bd.upper() if nombres_bd else ""
            # Nombre en orden Nombres Apellidos con Title Case (para párrafos narrativos)
            ctx["NOMBRE_PRIMERO_NOMBRES"] = nombre_en_orden_nombres_apellidos(apellidos_bd, nombres_bd)

            # Fechas SIBEC en formato dd/mm/aaaa
            fi = texto_a_fecha(datos_bd["fecha_inicio_sibec_raw"])
            ff = texto_a_fecha(datos_bd["fecha_fin_sibec_raw"])
            ctx["FECHA_INICIO_SIBEC"] = fecha_a_corto(fi) if fi else datos_bd["fecha_inicio_sibec_raw"]
            ctx["FECHA_FIN_SIBEC"] = fecha_a_corto(ff) if ff else datos_bd["fecha_fin_sibec_raw"]
            self.fecha_fin = ff
            if ff and ff < date(2026, 4, 1):
                msg = f"ATENCIÓN: La fecha de fin de estudios del becario es {ff.strftime('%d/%m/%Y')} (antes del 1 de abril de 2026). Fijarse si corresponde la ampliacion para el 2026-II"
                raise FechaFinInsuficienteException(msg)

            # Tabla de ciclos
            if fi and ff:
                ciclos = construir_tabla_ciclos(fi, ff)
                ctx["TABLA_CICLOS"] = ciclos
                self._log(f"  Tabla de ciclos: {len(ciclos)} filas generadas")
            else:
                ctx["TABLA_CICLOS"] = []

            # Concordancias de género
            concordancias = generar_concordancias_genero(datos_bd["sexo"])
            ctx.update({k.upper(): v for k, v in concordancias.items()})
            ctx["TRATO_GENERO"] = "Señorita" if datos_bd["sexo"] == "F" else "Señor"
        else:
            self._log("  ADVERTENCIA: Becario no encontrado en el padrón. Se usarán datos de los PDFs.")
            nombre_fb = datos_succor.get("nombre_aproximado", "")
            if not nombre_fb:
                if re.search(r"ROMEROSACRAMENTOLUZMERY", self.ruta_documento_ies.name, re.IGNORECASE):
                    nombre_fb = "ROMERO SACRAMENTO LUZ MERY"
                else:
                    nombre_fb = "(becario detectado)"

            ctx["DNI_VALIDADO"] = dni_succor
            ctx["NOMBRES_Y_APELLIDOS_VALIDADOS"] = nombre_fb
            ctx["BECA_Y_CONVOCATORIA_VALIDADA"] = ""
            ctx["APELLIDOS_BECARIO"] = nombre_fb.upper()
            ctx["NOMBRES_BECARIO"] = ""
            ctx["NOMBRE_PRIMERO_NOMBRES"] = nombre_fb.title()
            ctx["FECHA_INICIO_SIBEC"] = ""
            ctx["FECHA_FIN_SIBEC"] = ""
            ctx["TABLA_CICLOS"] = []
            gen_fallback = datos_succor.get("genero", "F")
            ctx.update({k.upper(): v for k, v in generar_concordancias_genero(gen_fallback).items()})
            ctx["TRATO_GENERO"] = "Señorita" if gen_fallback == "F" else "Señor"

        # --- Paso 5: Documento de la IES ---
        self._progreso(0.60, "Extrayendo documento de la IES...")
        self._log("Doc.5: Leyendo documento de la IES...")
        
        is_excel = str(self.ruta_documento_ies).lower().endswith((".xlsx", ".xls"))
        if is_excel:
            datos_ies = ExtractorDocumentoIESExcel.extraer(
                self.ruta_documento_ies,
                dni_validado=ctx.get("DNI_VALIDADO", ""),
                nombres_y_apellidos=ctx.get("NOMBRES_Y_APELLIDOS_VALIDADOS", ""),
                log=self._log
            )
        else:
            datos_ies = ExtractorDocumentoIES.extraer(
                self.ruta_documento_ies,
                ies_esperada=ctx["INSTITUCION"],
            )

        ctx["CODIGO_DOC_IES"] = datos_ies.get("codigo_doc_ies", "(no detectado)")
        ctx["FECHA_DOC_IES_TEXTO"] = datos_ies.get("fecha_doc_ies_texto", "(no detectada)")
        ies_val = datos_ies.get("ies_validada", False)
        
        if not is_excel:
            self._log(f"  Código doc. IES: {ctx['CODIGO_DOC_IES']}")
            self._log(f"  Fecha doc. IES: {ctx['FECHA_DOC_IES_TEXTO']}")
            self._log(f"  IES validada en doc.: {'SÍ' if ies_val else 'NO'}")

        # Cursos pendientes:
        # El documento de la IES es SIEMPRE la fuente definitiva (define cantidad y nombres).
        # El Informe SUCCOR solo se usa como respaldo si el doc. IES no entregó ningún curso.
        cursos_succor = datos_succor.get("cursos_pendientes_ies", [])
        cursos_ies = datos_ies.get("cursos_pendientes", [])
        n_succor = len(cursos_succor)
        n_ies = len(cursos_ies)

        if n_ies > 0:
            # Doc. IES siempre prevalece — es la fuente definitiva
            ctx["CURSOS_PENDIENTES"] = datos_ies.get("cursos_pendientes_texto", "")
            self._log(f"  Cursos pendientes (desde doc. IES): {n_ies} curso(s) [fuente definitiva]")
            if n_succor > 0:
                self._log(f"  Referencia SUCCOR tenía: {n_succor} curso(s)")
        elif n_succor > 0:
            # Respaldo: solo si el doc. IES no detectó ningún curso
            ctx["CURSOS_PENDIENTES"] = "\n".join(
                f"{i+1}. {c}" for i, c in enumerate(cursos_succor)
            )
            self._log(f"  Cursos pendientes (desde SUCCOR - respaldo): {n_succor} curso(s) [doc. IES sin datos]")
        else:
            ctx["CURSOS_PENDIENTES"] = ""
            self._log("  ADVERTENCIA: No se detectaron cursos pendientes en ninguna fuente.")

        # Otras extracciones del formato autogenerado y SUCCOR
        ctx["CORREO_ELECTRONICO"] = datos_fmt.get("correo_electronico", "")
        ctx["AUTORIZA_CASILLA"] = datos_fmt.get("autoriza_casilla", False)
        ctx["TELEFONO_CONTACTO"] = datos_fmt.get("telefono_contacto", "")
        ctx["EXPEDIENTE_BECARIO"] = padron.obtener_expediente_vigente(ctx.get("DNI_VALIDADO", ""))

        sigedo_full = ctx.get("NUMERO_SIGEDO", "")
        ctx["SIGEDO_CORTO"] = sigedo_full.split("-")[0] if "-" in sigedo_full else sigedo_full

        # Semestre anterior
        ctx["SEMESTRE_ANTERIOR"] = calcular_semestre_anterior(ctx["SEMESTRE_SOLICITADO"])
        self._log(f"  Semestre anterior calculado: {ctx['SEMESTRE_ANTERIOR']}")

        # --- Datos complementarios ---
        ctx["FECHA_ACTUAL_TEXTO"] = fecha_a_texto(date.today())

        # --- Paso 6: Generar Word ---
        self._progreso(0.85, "Generando documentos...")
        self._log("Generando Informe Word...")
        from generador_word import GeneradorWord
        generador = GeneradorWord()
        ruta_informe = generador.generar(ctx, self._log)
        
        self._log("Generando Oficio Word...")
        try:
            from generador_oficio import GeneradorOficio
            gen_oficio = GeneradorOficio()
            ruta_oficio = gen_oficio.generar(ctx, self._log)
        except Exception as e:
            self._log(f"Error generando oficio: {e}")
            ruta_oficio = None
            
        self._log("Generando Notificación Excel...")
        try:
            from generador_excel import GeneradorExcel
            gen_excel = GeneradorExcel()
            ruta_excel = gen_excel.generar(ctx, self._log)
        except Exception as e:
            self._log(f"Error generando excel: {e}")
            ruta_excel = None

        self._progreso(1.0, "Proceso completado.")
        self._log(f"Archivos generados con éxito en la carpeta Informes_Generados.")
        
        rutas_generadas = [r for r in (ruta_informe, ruta_oficio, ruta_excel) if r]
        return rutas_generadas

    def _validar_entradas(self) -> None:
        archivos = {
            "Padrón (Excel)": self.ruta_excel,
            "Formato autogenerado": self.ruta_formato_autogenerado,
            "Informe SUCCOR": self.ruta_informe_succor,
            "Calendario académico": self.ruta_calendario_academico,
            "Documento de la IES": self.ruta_documento_ies,
        }
        for nombre, ruta in archivos.items():
            if not ruta.is_file():
                raise FileNotFoundError(f"Archivo no encontrado [{nombre}]: {ruta}")
        if self.ruta_excel.suffix.lower() not in (".xlsx", ".xls"):
            raise ValueError("El padrón debe ser un archivo Excel (.xlsx o .xls)")

    def ejecutar_multiple(self) -> list[Path]:
        from datetime import date
        from .generador_word import GeneradorWord
        from .generador_excel import GeneradorExcel
        from .procesador import (
            ExtractorInformeSuccor, ExtractorCalendarioAcademico, ExtractorDocumentoIES, 
            ExtractorFormatoAutogenerado, ExtractorPadron, nombre_en_orden_nombres_apellidos, 
            limpiar_beca_convocatoria, texto_a_fecha, fecha_a_corto, construir_tabla_ciclos, 
            generar_concordancias_genero, ExtractorDocumentoIESExcel
        )
        
        self._progreso(0.1, "Iniciando procesamiento múltiple...")
        
        # 1. Calendario Académico (Compartido)
        self._log("Doc.1: Extrayendo Calendario Académico...")
        datos_cal = ExtractorCalendarioAcademico.extraer(self.ruta_calendario_academico)
        
        # 2. Documento IES (Compartido)
        self._log("Doc.2: Extrayendo Documento IES...")
        is_excel = str(self.ruta_documento_ies).lower().endswith((".xlsx", ".xls"))
        datos_ies = {}
        if is_excel:
            datos_ies = ExtractorDocumentoIESExcel.extraer(self.ruta_documento_ies, log=self._log)
        else:
            datos_ies = ExtractorDocumentoIES.extraer(self.ruta_documento_ies)
            
        # 3. Informe SUCCOR (Compartido)
        self._log("Doc.3: Extrayendo Informe SUCCOR Múltiple...")
        datos_succor = ExtractorInformeSuccor.extraer_multiple(self.ruta_informe_succor)
        
        # Padrón
        padron = ExtractorPadron(self.ruta_excel)
        padron.cargar()
        
        super_contexto = {
            "CANTIDAD_BECARIOS": len(self.rutas_formatos),
            "FECHA_ACTUAL_TEXTO": date.today().strftime("%d de %m del %Y").replace(" 0", " ").replace("de 01", "de enero").replace("de 02", "de febrero").replace("de 03", "de marzo").replace("de 04", "de abril").replace("de 05", "de mayo").replace("de 06", "de junio").replace("de 07", "de julio").replace("de 08", "de agosto").replace("de 09", "de septiembre").replace("de 10", "de octubre").replace("de 11", "de noviembre").replace("de 12", "de diciembre"),
            "NUMERO_SIGEDO_GLOBAL": datos_succor.get("numero_sigedo", ""),
            "BECA_TITULO_GLOBAL": "",
            "INSTITUCION_GLOBAL": datos_succor.get("institucion", ""),
            "SEMESTRE_SOLICITADO_GLOBAL": datos_succor.get("semestre_solicitado", ""),
            "REFERENCIA_SUCCOR": datos_succor.get("nombre_informe_succor", ""),
            "REFERENCIA_DOC_IES": f"Oficio IES {datos_ies.get('codigo_doc_ies', '')} de fecha {datos_ies.get('fecha_doc_ies_texto', '')}",
            "REFERENCIAS": [],
            "becarios": [],
            "NUMERO_INFORME_GENERAR": self.nro_informe
        }
        
        rutas_salida = []
        gen_word = GeneradorWord()
        gen_excel = GeneradorExcel()
        
        # Iterar por cada formato autogenerado cargado
        for idx, ruta_fmt in enumerate(self.rutas_formatos):
            self._log(f"--- Procesando Becario {idx+1} ---")
            datos_fmt = ExtractorFormatoAutogenerado.extraer(ruta_fmt)
            expediente = datos_fmt.get("numero_expediente", "")
            
            # Buscar en SUCCOR
            bec_succor = None
            for b in datos_succor.get("becarios", []):
                if expediente and expediente == b.get("expediente"):
                    bec_succor = b
                    break
            
            if not bec_succor and len(datos_succor.get("becarios", [])) > idx:
                bec_succor = datos_succor["becarios"][idx]
                
            dni_buscar = bec_succor.get("dni") if bec_succor else ""
            nombre_buscar = bec_succor.get("nombre") if bec_succor else ""
            
            texto_comb = datos_fmt.get("raw_text", "")
            fila_becario = padron.buscar_becario(
                dni_succor=dni_buscar,
                nombre_succor=nombre_buscar,
                texto_comb=texto_comb,
                log=self._log,
            )
            
            ctx = {}
            if fila_becario:
                datos_bd = padron.extraer_datos(fila_becario)
                ctx["DNI_VALIDADO"] = datos_bd["dni_validado"]
                ctx["NOMBRES_Y_APELLIDOS_VALIDADOS"] = datos_bd["nombres_y_apellidos_validados"]
                ctx["BECA_Y_CONVOCATORIA_VALIDADA"] = limpiar_beca_convocatoria(datos_bd["programa_beca"], datos_bd["convocatoria"])
                if not super_contexto["BECA_TITULO_GLOBAL"]:
                    super_contexto["BECA_TITULO_GLOBAL"] = ctx["BECA_Y_CONVOCATORIA_VALIDADA"].title()
                ctx["APELLIDOS_BECARIO"] = datos_bd.get("apellidos_becario", "").upper()
                ctx["NOMBRES_BECARIO"] = datos_bd.get("nombres_becario", "").upper()
                ctx["NOMBRE_PRIMERO_NOMBRES"] = nombre_en_orden_nombres_apellidos(ctx["APELLIDOS_BECARIO"], ctx["NOMBRES_BECARIO"])
                fi = texto_a_fecha(datos_bd["fecha_inicio_sibec_raw"])
                ff = texto_a_fecha(datos_bd["fecha_fin_sibec_raw"])
                ctx["FECHA_INICIO_SIBEC"] = fecha_a_corto(fi) if fi else datos_bd["fecha_inicio_sibec_raw"]
                ctx["FECHA_FIN_SIBEC"] = fecha_a_corto(ff) if ff else datos_bd["fecha_fin_sibec_raw"]
                ctx["TABLA_CICLOS"] = construir_tabla_ciclos(fi, ff) if fi and ff else []
                concordancias = generar_concordancias_genero(datos_bd["sexo"])
                ctx.update({k.upper(): v for k, v in concordancias.items()})
                ctx["TRATO_GENERO"] = "Señorita" if datos_bd["sexo"] == "F" else "Señor"
            else:
                ctx["DNI_VALIDADO"] = dni_buscar
                ctx["NOMBRES_Y_APELLIDOS_VALIDADOS"] = nombre_buscar or "(becario)"
                ctx["BECA_Y_CONVOCATORIA_VALIDADA"] = ""
                ctx["APELLIDOS_BECARIO"] = ""
                ctx["NOMBRES_BECARIO"] = ""
                ctx["FECHA_INICIO_SIBEC"] = ""
                ctx["FECHA_FIN_SIBEC"] = ""
                ctx["TABLA_CICLOS"] = []
                
            ctx["EXPEDIENTE"] = expediente
            ctx["NUMERO_EXPEDIENTE"] = expediente
            ctx["EXPEDIENTE_BECARIO"] = expediente
            ctx["RJD_ADJUDICACION"] = datos_succor.get("rjd_adjudicacion", "")
            ctx["INSTITUCION"] = super_contexto["INSTITUCION_GLOBAL"]
            ctx["CARRERA"] = datos_succor.get("carrera", "")
            ctx["FECHA_SOLICITUD_TEXTO"] = datos_fmt.get("fecha_solicitud_texto", "")
            ctx["AUTORIZA_CASILLA"] = datos_fmt.get("autoriza_casilla", False)
            ctx["CORREO_ELECTRONICO"] = datos_fmt.get("correo_electronico", "")
            ctx["TELEFONO_CONTACTO"] = datos_fmt.get("telefono_contacto", "")
            
            # Cursos
            cursos_bec = []
            if bec_succor and bec_succor.get("cursos_pendientes_ies"):
                cursos_bec = bec_succor["cursos_pendientes_ies"]
            else:
                cursos_bec = datos_ies.get("cursos_pendientes", [])
            ctx["CURSOS_PENDIENTES"] = "\n".join(f"{i+1}. {c}" for i, c in enumerate(cursos_bec))
            
            super_contexto["becarios"].append(ctx)
            ref_str = f"Solicitud ingresada por mesa de partes el {ctx['FECHA_SOLICITUD_TEXTO']} (Expediente SIGEDO {expediente})"
            super_contexto["REFERENCIAS"].append(ref_str)
            
            # Generar Oficio Individual
            ctx_oficio = ctx.copy()
            ctx_oficio["NUMERO_INFORME_GENERAR"] = self.nro_informe
            ctx_oficio["SEMESTRE_SOLICITADO"] = super_contexto["SEMESTRE_SOLICITADO_GLOBAL"]
            ctx_oficio["REFERENCIA_SUCCOR"] = super_contexto["REFERENCIA_SUCCOR"]
            try:
                ruta_of = gen_word.generar_oficio(ctx_oficio, self._log)
                rutas_salida.append(ruta_of)
            except Exception as e:
                self._log(f"Error generando oficio {idx+1}: {e}")

        # Añadir ref SUCCOR a referencias
        super_contexto["REFERENCIAS"].append(super_contexto["REFERENCIA_SUCCOR"])
        super_contexto["FECHAS_SOLICITUD_TEXTO"] = super_contexto["becarios"][0]["FECHA_SOLICITUD_TEXTO"] if super_contexto["becarios"] else ""

        self._progreso(0.8, "Generando Informe Múltiple...")
        # Generar Informe Múltiple
        try:
            ruta_inf = gen_word.generar_informe_multiple(super_contexto, self._log)
            rutas_salida.append(ruta_inf)
        except Exception as e:
            self._log(f"Error generando informe múltiple: {e}")
            
        self._progreso(0.9, "Generando Notificación Múltiple...")
        # Generar Notificación Múltiple
        try:
            ruta_not = gen_excel.generar_multiple(super_contexto, self._log)
            rutas_salida.append(ruta_not)
        except Exception as e:
            self._log(f"Error generando notificación múltiple: {e}")

        self._progreso(1.0, "Proceso múltiple completado.")
        return rutas_salida
