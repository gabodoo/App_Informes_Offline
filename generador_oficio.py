from __future__ import annotations

import sys
import re
import shutil
from datetime import datetime, date
from pathlib import Path
from typing import Callable, Any

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


def obtener_ruta_plantilla_oficio() -> Path:
    """Busca y retorna la ruta existente de plantilla_oficio.docx probando múltiples ubicaciones."""
    rutas_candidatas = [
        BASE_DIR / "plantillas" / "plantilla_oficio.docx",
        BASE_DIR / "dist" / "main" / "plantillas" / "plantilla_oficio.docx",
        BASE_DIR / "dist" / "plantillas" / "plantilla_oficio.docx",
        Path(__file__).resolve().parent / "plantillas" / "plantilla_oficio.docx",
    ]
    for ruta in rutas_candidatas:
        if ruta.is_file():
            return ruta
    return BASE_DIR / "plantillas" / "plantilla_oficio.docx"


class GeneradorOficio:
    """Generador híbrido robusto para el Oficio oficial de Ampliación de Estudios.
    Aplica docxtpl (Jinja2) y reemplazo profundo con python-docx, cubriendo todas las sintaxis
    de marcadores ({{VAR}}, [VAR], <VAR>, ${VAR}, {VAR}) y fallbacks de texto estático."""

    def __init__(self):
        self.ruta_plantilla = obtener_ruta_plantilla_oficio()

    def generar(self, contexto: dict, log_callback: LogCallback | None = None) -> Path:
        def log(msg):
            if log_callback:
                log_callback(msg)

        if not self.ruta_plantilla.exists():
            log(f"No se encontró la plantilla del oficio en {self.ruta_plantilla}")
            raise FileNotFoundError(f"Falta plantilla: {self.ruta_plantilla}")

        log(f"Cargando plantilla de oficio: {self.ruta_plantilla.name}...")

        # 1. Normalización de SIGEDO (extraer dígitos limpios para evitar repeticiones o -2026-2026)
        sigedo_raw = str(
            contexto.get("NUMERO_SIGEDO")
            or contexto.get("NUMERO_SIGEDO_GLOBAL")
            or contexto.get("SIGEDO")
            or ""
        ).strip()
        m_sig = re.search(r"(\d{4,8})", sigedo_raw)
        if m_sig:
            sigedo_corto = m_sig.group(1)
            reemplazo_sigedo = f"{sigedo_corto}-2026"
        else:
            sigedo_corto = contexto.get("SIGEDO_CORTO") or (sigedo_raw.split("-")[0] if sigedo_raw else "")
            reemplazo_sigedo = sigedo_raw

        # 2. Trato de género estricto: "Señorita" para femenino, "Señor" para masculino
        sexo_val = str(contexto.get("SEXO") or "").strip().upper()
        trato_val = str(contexto.get("TRATO_GENERO") or contexto.get("TRATO") or "").strip()

        if trato_val.lower() in ("señorita", "senorita"):
            trato = "Señorita"
        elif trato_val.lower() in ("señor", "senor"):
            trato = "Señor"
        elif sexo_val.startswith("F") or sexo_val in ("2", "MUJER", "FEMENINO"):
            trato = "Señorita"
        else:
            trato = "Señor"

        # 3. Nombres y Apellidos del becario (en mayúsculas, en orden Nombres Apellidos, sin comas)
        n_bec = str(contexto.get("NOMBRES_BECARIO", "")).strip()
        a_bec = str(contexto.get("APELLIDOS_BECARIO", "")).strip()
        if n_bec and a_bec:
            reemplazo_nombres = f"{n_bec} {a_bec}".strip().upper()
        else:
            raw_nom = str(
                contexto.get("NOMBRE_PRIMERO_NOMBRES")
                or contexto.get("NOMBRES_Y_APELLIDOS_VALIDADOS")
                or contexto.get("DESTINATARIO")
                or contexto.get("NOMBRES_Y_APELLIDOS")
                or ""
            ).strip().upper()
            if "," in raw_nom:
                partes = [p.strip() for p in raw_nom.split(",", 1)]
                reemplazo_nombres = f"{partes[1]} {partes[0]}".strip().upper()
            else:
                reemplazo_nombres = raw_nom

        if "," in reemplazo_nombres:
            partes = [p.strip() for p in reemplazo_nombres.split(",", 1)]
            reemplazo_nombres = f"{partes[1]} {partes[0]}".strip().upper()

        # 4. Fecha de solicitud (proveniente del formato autogenerado)
        val_fsol = str(
            contexto.get("FECHA_SOLICITUD_TEXTO")
            or contexto.get("FECHA_SOLICITUD")
            or contexto.get("FECHA_SOL")
            or ""
        ).strip()
        if not val_fsol or val_fsol in ("(no detectada)", "None"):
            val_fsol = "17 de julio de 2026"

        reemplazo_fecha_escrito = f"Escrito de fecha {val_fsol}"

        # 5. Informe de la SUS autogenerado
        nro_inf_gen = str(contexto.get("NUMERO_INFORME_GENERAR", "")).strip()
        if nro_inf_gen and nro_inf_gen != "S-N":
            reemplazo_informe_sus = f"Informe N° {nro_inf_gen}-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS"
        else:
            reemplazo_informe_sus = "Informe N° -2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS"

        # 6. Semestre solicitado y Asunto
        semestre = str(contexto.get("SEMESTRE_SOLICITADO") or "2026-II").strip()
        asunto_oficio = str(
            contexto.get("ASUNTO_OFICIO")
            or f"Ampliación excepcional de la duración del periodo de estudios por el semestre {semestre}"
        ).strip()

        # 7. Mapeo exhaustivo de claves y alias para cubrir cualquier marcador
        alias_map = {
            "NUMERO_SIGEDO": [
                "SIGEDO", "EXPEDIENTE_SIGEDO", "NUMERO_SIGEDO_GLOBAL", "SIGEDO_OFICIO"
            ],
            "SIGEDO_CORTO": [
                "SIGEDO_NUMERO", "NUM_SIGEDO"
            ],
            "TRATO_GENERO": [
                "TRATO", "SENOR_SENORITA", "SENOR_O_SENORITA", "TRATO_BECARIO", "TRATO_DESTINATARIO", "SENORITA_SENOR"
            ],
            "NOMBRES_Y_APELLIDOS": [
                "NOMBRES_Y_APELLIDOS_VALIDADOS", "NOMBRE_BECARIO", "DESTINATARIO",
                "NOMBRE_PRIMERO_NOMBRES", "NOMBRE_COMPLETO", "APELLIDOS_Y_NOMBRES", "NOMBRES_APELLIDOS"
            ],
            "FECHA_SOLICITUD_TEXTO": [
                "FECHA_SOLICITUD", "FECHA_SOL", "FECHA_DE_SOLICITUD", "FECHA_ESCRITO"
            ],
            "INFORME_SUS": [
                "NUMERO_INFORME_SUS", "INFORME_DIBEC_SUS", "INFORME_DE_LA_SUS", "INFORME_SUS_COMPLETO"
            ],
            "NUMERO_INFORME_GENERAR": [
                "NUMERO_INFORME", "NRO_INFORME", "NUM_INFORME"
            ],
            "SEMESTRE_SOLICITADO": [
                "SEMESTRE", "PERIODO_AMPLIACION", "SEMESTRE_AMPLIACION"
            ],
            "ASUNTO_OFICIO": [
                "ASUNTO"
            ],
        }

        # 8. Preparar contexto expandido
        ctx_expandido = dict(contexto)
        ctx_expandido["NUMERO_SIGEDO"] = reemplazo_sigedo
        ctx_expandido["SIGEDO"] = reemplazo_sigedo
        ctx_expandido["SIGEDO_CORTO"] = sigedo_corto
        ctx_expandido["TRATO_GENERO"] = trato
        ctx_expandido["NOMBRES_Y_APELLIDOS"] = reemplazo_nombres
        ctx_expandido["NOMBRES_Y_APELLIDOS_VALIDADOS"] = reemplazo_nombres
        ctx_expandido["DESTINATARIO"] = reemplazo_nombres
        ctx_expandido["FECHA_SOLICITUD_TEXTO"] = val_fsol
        ctx_expandido["FECHA_SOLICITUD"] = val_fsol
        ctx_expandido["INFORME_SUS"] = reemplazo_informe_sus
        ctx_expandido["NUMERO_INFORME_GENERAR"] = nro_inf_gen
        ctx_expandido["SEMESTRE_SOLICITADO"] = semestre
        ctx_expandido["ASUNTO_OFICIO"] = asunto_oficio

        for clave_principal, lista_alias in alias_map.items():
            if clave_principal in ctx_expandido:
                val = ctx_expandido[clave_principal]
                for alias in lista_alias:
                    if alias not in ctx_expandido:
                        ctx_expandido[alias] = val

        # 9. Contexto limpio y expansión de casings
        ctx_limpio = {}
        mapa_reemplazos = {}
        for k, v in ctx_expandido.items():
            if v is None:
                val_str = ""
            elif isinstance(v, (list, dict)):
                val_str = str(v)
            else:
                val_str = str(v)
            ctx_limpio[k] = val_str
            ctx_limpio[k.lower()] = val_str
            ctx_limpio[k.upper()] = val_str
            ctx_limpio[k.title()] = val_str
            mapa_reemplazos[k.upper()] = val_str

        nombre_salida = contexto.get("NOMBRE_SALIDA_OFICIO", f"{sigedo_corto}_oficio.docx")
        ruta_salida = BASE_DIR / "Informes_Generados" / nombre_salida
        ruta_salida.parent.mkdir(parents=True, exist_ok=True)

        # PASO 1: Renderizado mediante docxtpl (Jinja2)
        try:
            tpl = DocxTemplate(self.ruta_plantilla)
            tpl.render(ctx_limpio)
            tpl.save(ruta_salida)
        except Exception as err:
            log(f"  [AVISO docxtpl oficio]: {err}. Se aplicará reemplazo directo con python-docx.")
            shutil.copy(self.ruta_plantilla, ruta_salida)

        # PASO 2: Reemplazo profundo con python-docx
        self._reemplazo_profundo_docx(
            ruta_salida=ruta_salida,
            mapa_reemplazos=mapa_reemplazos,
            reemplazo_sigedo=reemplazo_sigedo,
            sigedo_corto=sigedo_corto,
            trato=trato,
            reemplazo_nombres=reemplazo_nombres,
            val_fsol=val_fsol,
            reemplazo_fecha_escrito=reemplazo_fecha_escrito,
            reemplazo_informe_sus=reemplazo_informe_sus,
            nro_inf_gen=nro_inf_gen,
            asunto_oficio=asunto_oficio,
            log=log,
        )

        log(f"Oficio generado exitosamente en: {ruta_salida.name}")
        return ruta_salida

    def _reemplazo_profundo_docx(
        self,
        ruta_salida: Path,
        mapa_reemplazos: dict[str, str],
        reemplazo_sigedo: str,
        sigedo_corto: str,
        trato: str,
        reemplazo_nombres: str,
        val_fsol: str,
        reemplazo_fecha_escrito: str,
        reemplazo_informe_sus: str,
        nro_inf_gen: str,
        asunto_oficio: str,
        log: LogCallback | None = None,
    ) -> None:
        """Aplica reemplazo directo a nivel de runs, párrafos, tablas, encabezados y cuadros flotantes."""
        doc = Document(ruta_salida)

        # Patrones de marcadores con todas las sintaxis comunes ({{VAR}}, [VAR], <VAR>, ${VAR}, {VAR})
        patrones_reemplazo = []
        for clave, valor in mapa_reemplazos.items():
            if not clave:
                continue
            for n in {clave, clave.lower(), clave.title()}:
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

        # Fallbacks estáticos (palabras completas para no corromper subcadenas)
        reemplazos_fallback = [
            ("PERICO DE LOS PALOTES", reemplazo_nombres),
            ("Perico De Los Palotes", reemplazo_nombres.title()),
            ("MAINIA ROSITA SEKUTA KAYACH", reemplazo_nombres),
            ("Mainia Rosita Sekuta Kayach", reemplazo_nombres.title()),
            ("Informe N° 1234-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS", reemplazo_informe_sus),
            ("Informe N° 6541-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS", reemplazo_informe_sus),
            ("Informe N° 1234-2026-MINEDU/VMGI-PRONABEC-DIBEC", reemplazo_informe_sus),
            ("Informe N° 6541-2026-MINEDU/VMGI-PRONABEC-DIBEC", reemplazo_informe_sus),
            ("17 de julio de 2026", val_fsol),
            ("17 DE JULIO DE 2026", val_fsol.upper()),
        ]

        if nro_inf_gen:
            reemplazos_fallback.extend([
                ("Informe N° 1234", f"Informe N° {nro_inf_gen}"),
                ("Informe N° 6541", f"Informe N° {nro_inf_gen}"),
            ])

        # Regex estricta para trato con límite de palabra (evita sustituciones parciales recursivas)
        patron_trato = re.compile(r"\b(?:Señorita|Señor|Señoritaita|Señorilaita)\b")

        patron_escrito = re.compile(
            r"Escrito\s+de\s+fecha\s+[^,\.\n\(\);]+?(?:de\s+\d{4}|\b\d{4}\b)",
            re.IGNORECASE,
        )

        def sanitizar_texto(texto: str) -> str:
            if not texto:
                return texto
            # Eliminar duplicaciones de años en SIGEDO (ej: 63671-2026-2026 -> 63671-2026)
            texto = re.sub(r"-2026(?:-2026)+", "-2026", texto)
            # Reemplazo seguro de trato de género
            if not ("{{" in texto or "}}" in texto):
                texto = patron_trato.sub(trato, texto)
            return texto

        def _reemplazar_texto_run(run):
            texto_nuevo = run.text
            for tag, val in patrones_reemplazo:
                if tag in texto_nuevo:
                    texto_nuevo = texto_nuevo.replace(tag, val)
            for old_val, new_val in reemplazos_fallback:
                if old_val in texto_nuevo:
                    texto_nuevo = texto_nuevo.replace(old_val, new_val)
            if patron_escrito.search(texto_nuevo):
                texto_nuevo = patron_escrito.sub(reemplazo_fecha_escrito, texto_nuevo)
            texto_nuevo = sanitizar_texto(texto_nuevo)
            if texto_nuevo != run.text:
                run.text = texto_nuevo

        def procesar_parrafo(p):
            if not p.text:
                return

            # Si el párrafo es exclusivamente la línea de saludo/trato al becario
            p_strip = p.text.strip()
            if p_strip in ("Señorita", "Señor", "Señoritaita", "Señorilaita", "{{ TRATO_GENERO }}", "{{TRATO_GENERO}}", "{{ TRATO }}", "{{TRATO}}", "[TRATO_GENERO]"):
                p.text = trato
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(11)
                return

            texto_pre = p.text
            for run in p.runs:
                _reemplazar_texto_run(run)

            # Fallback a nivel de párrafo si las etiquetas estaban fragmentadas entre runs
            if p.text == texto_pre:
                texto_orig = p.text
                texto_nuevo = texto_orig
                for tag, val in patrones_reemplazo:
                    if tag in texto_nuevo:
                        texto_nuevo = texto_nuevo.replace(tag, val)
                for old_val, new_val in reemplazos_fallback:
                    if old_val in texto_nuevo:
                        texto_nuevo = texto_nuevo.replace(old_val, new_val)
                if patron_escrito.search(texto_nuevo):
                    texto_nuevo = patron_escrito.sub(reemplazo_fecha_escrito, texto_nuevo)
                texto_nuevo = sanitizar_texto(texto_nuevo)

                if texto_nuevo != texto_orig:
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

        # 1. Párrafos del cuerpo
        for p in doc.paragraphs:
            procesar_parrafo(p)

        # 2. Tablas del cuerpo
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        procesar_parrafo(p)

        # 3. Encabezados y pies de página (incluyendo sus tablas)
        for section in doc.sections:
            for header in [section.header, section.first_page_header]:
                if header:
                    for p in header.paragraphs:
                        procesar_parrafo(p)
                    for table in header.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    procesar_parrafo(p)

            for footer in [section.footer, section.first_page_footer]:
                if footer:
                    for p in footer.paragraphs:
                        procesar_parrafo(p)
                    for table in footer.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    procesar_parrafo(p)

        # 4. Cuadros de texto flotantes (SIGEDO en cuadro superior derecho)
        # Se consolida en el primer nodo de texto y se vacían los restantes para evitar duplicaciones
        for p_elem in doc._element.xpath('.//w:txbxContent//w:p'):
            t_nodes = p_elem.xpath('.//w:t')
            if not t_nodes:
                continue
            full_text = "".join(t.text or "" for t in t_nodes)
            if any(k in full_text.upper() for k in ["SIGEDO", "12345", "57897", "{{ NUMERO_SIGEDO }}", "{{NUMERO_SIGEDO}}", "{{ SIGEDO }}", "{{SIGEDO}}"]):
                t_nodes[0].text = f"SIGEDO: {reemplazo_sigedo}"
                for t in t_nodes[1:]:
                    t.text = ""
            else:
                for t in t_nodes:
                    if t.text:
                        t.text = sanitizar_texto(t.text)

        # 5. Barrido final de nodos <w:t> para etiquetas residuales
        for node in doc._element.xpath('.//w:t'):
            if not node.text:
                continue
            for tag, val in patrones_reemplazo:
                if tag in node.text:
                    node.text = node.text.replace(tag, val)
            if "-2026-2026" in node.text:
                node.text = re.sub(r"-2026(?:-2026)+", "-2026", node.text)

        doc.save(ruta_salida)
