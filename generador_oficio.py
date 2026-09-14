import re
from pathlib import Path
from docx import Document
from docx.shared import Pt
import sys

def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "main" and exe_dir.parent.name.lower() == "dist":
            return exe_dir.parent.parent
        if exe_dir.name.lower() == "dist":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


class GeneradorOficio:
    """Clase para la generación del Oficio a partir de su plantilla."""

    def __init__(self):
        self.ruta_plantilla = get_base_dir() / "plantillas" / "plantilla_oficio.docx"

    def generar(self, contexto: dict, log_callback=None) -> Path:
        def log(msg):
            if log_callback:
                log_callback(msg)

        if not self.ruta_plantilla.exists():
            log(f"No se encontró la plantilla del oficio en {self.ruta_plantilla}")
            raise FileNotFoundError(f"Falta plantilla: {self.ruta_plantilla}")

        log("Cargando plantilla de oficio...")
        doc = Document(self.ruta_plantilla)
        
        # 1. Reemplazos directos
        sigedo_completo = str(contexto.get("NUMERO_SIGEDO") or contexto.get("NUMERO_SIGEDO_GLOBAL") or "")
        sigedo_corto = contexto.get("SIGEDO_CORTO") or (sigedo_completo.split("-")[0] if sigedo_completo else "")
        reemplazo_sigedo = sigedo_completo if sigedo_completo else sigedo_corto
        n_bec = str(contexto.get("NOMBRES_BECARIO", "")).strip()
        a_bec = str(contexto.get("APELLIDOS_BECARIO", "")).strip()
        reemplazo_nombres = f"{n_bec} {a_bec}".strip()
        if not reemplazo_nombres:
            reemplazo_nombres = str(contexto.get("NOMBRES_Y_APELLIDOS_VALIDADOS", "")).upper()
        reemplazo_informe = f"Informe N° {contexto.get('NUMERO_INFORME_GENERAR', '')}-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS"
        reemplazo_fecha = f"Escrito de fecha {contexto.get('FECHA_SOLICITUD_TEXTO', '')}"
        trato = str(contexto.get("TRATO_GENERO", "Señorita"))
        
        patrones_reemplazo = [
            ("57897-2026", reemplazo_sigedo),
            ("Señorita", trato),
            ("MAINIA ROSITA SEKUTA KAYACH", reemplazo_nombres),
            ("Informe N° 6541-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS", reemplazo_informe),
            ("Informe N° 6541-2026-MINEDU/VMGI-PRONABEC-DIBEC", reemplazo_informe),
            ("Informe N° 6541", reemplazo_informe),
            ("57897", sigedo_corto),
        ]

        patron_escrito = re.compile(r"Escrito\s+de\s+fecha\s+[^,\.\n\(\);]+?(?:de\s+\d{4}|\b\d{4}\b)", re.IGNORECASE)

        def _reemplazar_texto_run(run):
            texto_nuevo = run.text
            for tag, val in patrones_reemplazo:
                if tag in texto_nuevo:
                    texto_nuevo = texto_nuevo.replace(tag, val)
            if patron_escrito.search(texto_nuevo):
                texto_nuevo = patron_escrito.sub(reemplazo_fecha, texto_nuevo)
            
            if texto_nuevo != run.text:
                run.text = texto_nuevo

        def procesar_parrafo(p):
            if not p.text:
                return
                
            texto_pre = p.text
            for run in p.runs:
                _reemplazar_texto_run(run)

            # Fallback a nivel de párrafo
            if p.text == texto_pre:
                texto_orig = p.text
                texto_nuevo = texto_orig
                for tag, val in patrones_reemplazo:
                    if tag in texto_nuevo:
                        texto_nuevo = texto_nuevo.replace(tag, val)
                if patron_escrito.search(texto_nuevo):
                    texto_nuevo = patron_escrito.sub(reemplazo_fecha, texto_nuevo)
                
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

        # Procesar documento entero
        for p in doc.paragraphs:
            procesar_parrafo(p)
            
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        procesar_parrafo(p)

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

        from lxml import etree
        for node in doc._element.xpath('.//w:t'):
            if not node.text: continue
        nombre_salida = contexto.get("NOMBRE_SALIDA_OFICIO", f"{sigedo_corto}_oficio.docx")
        ruta_salida = get_base_dir() / "Informes_Generados" / nombre_salida
        ruta_salida.parent.mkdir(parents=True, exist_ok=True)
        
                # Actualizar casillas de texto (SIGEDO en cuadro de texto flotante)
        if reemplazo_sigedo:
            for p_elem in doc._element.xpath('.//w:txbxContent//w:p'):
                for t in p_elem.xpath('.//w:t'):
                    if "SIGEDO" in t.text or "57897" in t.text:
                        t.text = f"SIGEDO: {reemplazo_sigedo}"

        log(f"Guardando Oficio en: {ruta_salida.name}")
        doc.save(ruta_salida)
        return ruta_salida
