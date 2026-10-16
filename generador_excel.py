import sys
import re
from pathlib import Path
from copy import copy
from typing import Any
import openpyxl
from openpyxl.cell.cell import MergedCell


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "main" and exe_dir.parent.name.lower() == "dist":
            return exe_dir.parent.parent
        if exe_dir.name.lower() == "dist":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


def obtener_ruta_plantilla_excel(nombre_archivo: str) -> Path:
    """Busca la plantilla Excel en múltiples rutas candidatas para garantizar compatibilidad total (desarrollo y .exe)."""
    base = get_base_dir()
    candidatos = [
        base / "plantillas" / nombre_archivo,
        base / "dist" / "main" / "plantillas" / nombre_archivo,
        base / "dist" / "plantillas" / nombre_archivo,
        Path(__file__).resolve().parent / "plantillas" / nombre_archivo,
        Path.cwd() / "plantillas" / nombre_archivo,
    ]
    for c in candidatos:
        if c.is_file():
            return c
    return base / "plantillas" / nombre_archivo


def obtener_ruta_plantilla_acto_administrativo() -> Path:
    """Busca la plantilla Excel de notificación de acto administrativo (con o sin guion bajo/espacio)."""
    base = get_base_dir()
    nombres = [
        "plantilla_notificacion_acto_administrativo.xlsx",
        "plantilla_notificacion_acto administrativo.xlsx",
    ]
    candidatos = []
    for nom in nombres:
        candidatos.extend([
            base / "plantillas" / nom,
            base / "dist" / "main" / "plantillas" / nom,
            base / "dist" / "plantillas" / nom,
            Path(__file__).resolve().parent / "plantillas" / nom,
            Path.cwd() / "plantillas" / nom,
        ])
    for c in candidatos:
        if c.is_file():
            return c
    return base / "plantillas" / "plantilla_notificacion_acto_administrativo.xlsx"


def actualizar_texto_items_anexo(texto_original: str, num_inf: str, fecha_sol_texto: str) -> str:
    """
    Actualiza el texto del campo NUMERO DE (ITEM) DEL ARCHIVO PARA ADJUNTAR EN EL CORREO COMO ANEXO.
    Modifica únicamente:
    - II) INFORME Nº [numero]-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS con el número de informe generado.
    - IV) Solicitud de fecha [fecha] con la fecha de solicitud del formato autogenerado.
    Funciona independientemente de los datos ficticios, sanitizados o números de prueba que contenga la plantilla.
    """
    num_inf_limpio = str(num_inf).strip()
    m_num = re.search(r"^\d+", num_inf_limpio)
    if m_num:
        num_inf_limpio = m_num.group(0)

    f_texto = str(fecha_sol_texto).strip()
    if not f_texto or f_texto in ("(no detectada)", ""):
        f_texto = "21 de julio de 2026"

    nuevo_inf_txt = f"INFORME Nº {num_inf_limpio}-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS"
    nuevo_sol_txt = f"Solicitud de fecha {f_texto}"

    if not texto_original or not str(texto_original).strip():
        return (
            f"I) Oficio N° xxx-2026-MINEDU/VMGI-PRONABEC-DIBEC, "
            f"II) {nuevo_inf_txt}, "
            f"III) ANEXO 01: FICHA DE REGISTRO - MESA DE PARTES DIGITAL, "
            f"IV) {nuevo_sol_txt}"
        )

    nuevo_texto = str(texto_original)

    # 1. Reemplazo del Item II
    # Intento A: Reemplazar el bloque completo del Item II entre 'II)' y el siguiente item 'III)'
    patron_entre_ii_y_iii = re.compile(
        r"(II[\)\.\-]\s*)(.*?)(?=(?:\s*[,;\n\r]+\s*|\s+)III[\)\.\-])",
        re.IGNORECASE | re.DOTALL,
    )
    if patron_entre_ii_y_iii.search(nuevo_texto):
        nuevo_texto = patron_entre_ii_y_iii.sub(rf"\g<1>{nuevo_inf_txt}", nuevo_texto)
    else:
        # Intento B: Si no hay delimitador 'III)', buscar patrón general de Informe en II)
        patron_ii = re.compile(
            r"(II[\)\.\-]\s*INFORME\s*N[°ºo\.]?\s*)[^\n\r,–—]+?(-MINEDU/VMGI-PRONABEC-DIBEC(?:-SUS)?)",
            re.IGNORECASE,
        )
        if patron_ii.search(nuevo_texto):
            nuevo_texto = patron_ii.sub(rf"\g<1>{num_inf_limpio}-2026\g<2>", nuevo_texto)
        else:
            patron_ii_alt = re.compile(
                r"(II[\)\.\-]\s*INFORME\s*N[°ºo\.]?\s*)[^\n\r,]+",
                re.IGNORECASE,
            )
            if patron_ii_alt.search(nuevo_texto):
                nuevo_texto = patron_ii_alt.sub(rf"\g<1>{num_inf_limpio}-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS", nuevo_texto)
            elif "8552" in nuevo_texto:
                nuevo_texto = nuevo_texto.replace("8552", str(num_inf_limpio))
            elif "1234" in nuevo_texto:
                nuevo_texto = re.sub(r"\b1234\b", str(num_inf_limpio), nuevo_texto)

    # Limpieza residual de valores de prueba de informes
    if "1234" in nuevo_texto:
        nuevo_texto = re.sub(r"\b1234\b", str(num_inf_limpio), nuevo_texto)

    # 2. Reemplazo del Item IV
    # Intento A: Reemplazar todo el contenido del Item IV en su línea
    patron_iv_linea = re.compile(
        r"(IV[\)\.\-]\s*)[^\n\r]+",
        re.IGNORECASE,
    )
    if patron_iv_linea.search(nuevo_texto):
        nuevo_texto = patron_iv_linea.sub(rf"\g<1>{nuevo_sol_txt}", nuevo_texto)
    else:
        # Intento B: Si está en una sola línea separado por comas
        patron_iv_coma = re.compile(
            r"(IV[\)\.\-]\s*)[^,;]+",
            re.IGNORECASE,
        )
        if patron_iv_coma.search(nuevo_texto):
            nuevo_texto = patron_iv_coma.sub(rf"\g<1>{nuevo_sol_txt}", nuevo_texto)
        elif "21 de julio de 2026" in nuevo_texto:
            nuevo_texto = nuevo_texto.replace("21 de julio de 2026", str(f_texto))

    if "21 de julio de 2026" in nuevo_texto:
        nuevo_texto = nuevo_texto.replace("21 de julio de 2026", str(f_texto))

    # 3. Soporte para placeholders explícitos que se hayan colocado en el anexo
    nuevo_texto = re.sub(r"\{\{INFORME(?:_SUS)?\}\}|\{INFORME(?:_SUS)?\}", nuevo_inf_txt, nuevo_texto, flags=re.IGNORECASE)
    nuevo_texto = re.sub(r"\{\{FECHA(?:_SOLICITUD)?\}\}|\{FECHA(?:_SOLICITUD)?\}", f_texto, nuevo_texto, flags=re.IGNORECASE)

    return nuevo_texto


def _detectar_fila_encabezado_y_datos(ws) -> tuple[int, int]:
    """
    Detecta de forma robusta la fila de encabezados y la fila de datos en la hoja,
    independientemente de los valores de prueba o sanitización que contenga.
    Retorna (fila_encabezado, fila_datos).
    """
    palabras_clave_hdr = [
        "SIGEDO", "TIPO DE USUARIO", "DNI", "RUC", "DESTINATARIO",
        "CORREO", "TELEFONO", "ITEM", "ANEXO", "AGOTAMIENTO", "EXPEDIENTE"
    ]

    fila_hdr = 2
    max_coincidencias = 0

    # 1. Encontrar la fila con mayor número de títulos de columna (fila de encabezados)
    for r in range(1, min(ws.max_row + 1, 10)):
        coincidencias = 0
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r, column=c).value or "").strip().upper()
            if not val or val.startswith("{{") or val.startswith("{"):
                continue
            if any(k in val for k in palabras_clave_hdr):
                coincidencias += 1
        if coincidencias > max_coincidencias and coincidencias >= 3:
            max_coincidencias = coincidencias
            fila_hdr = r

    # 2. La fila de datos por defecto es la inmediatamente posterior al encabezado (ej. fila 3)
    fila_datos = fila_hdr + 1

    # 3. Verificación de seguridad: solo buscar filas de datos entre fila_hdr + 1 y las notas explicativas
    # Detenerse inmediatamente al encontrar las filas de instrucciones (ej. PARA AGOTAMIENTO o * Si registras)
    for r in range(fila_hdr + 1, min(ws.max_row + 1, 15)):
        primera_celda = str(ws.cell(row=r, column=1).value or "").strip().upper()
        if "PARA AGOTAMIENTO" in primera_celda or primera_celda.startswith("*"):
            break
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r, column=c).value or "").strip()
            # Buscar estructura estricta de anexo II) o I) al inicio de una línea o ítem (evitando falsos positivos con "(SI)")
            if re.search(r"(?:^|[\r\n;,]+)\s*I{1,3}\)\s+", val):
                fila_datos = r
                return fila_hdr, fila_datos

    return fila_hdr, fila_datos



def _mapear_columnas_acto(ws, ultima_fila_hdr: int) -> dict:
    """Mapea nombres de campo a índices de columna escaneando las filas de encabezado."""
    cols = {
        "sigedo": None,
        "tipo_usuario": None,
        "dni": None,
        "destinatario": None,
        "correo": None,
        "telefono": None,
        "item_anexo": None,
        "agotamiento": None,
        "expediente": None,
        "cargo": None,
    }

    for c in range(1, ws.max_column + 1):
        # Leer directamente de la fila de encabezado real para evitar banners de fila 1
        hdr = str(ws.cell(row=ultima_fila_hdr, column=c).value or "").strip().upper()
        if not hdr:
            continue

        if "TIPO" in hdr and "USUARIO" in hdr:
            if cols["tipo_usuario"] is None: cols["tipo_usuario"] = c
        elif "SIGEDO" in hdr:
            if cols["sigedo"] is None: cols["sigedo"] = c
        elif ("DNI" in hdr or "RUC" in hdr or "DOCUMENTO DE IDENTIDAD" in hdr) and "TIPO" not in hdr:
            if cols["dni"] is None: cols["dni"] = c
        elif "CARGO" in hdr and "DESTINATARIO" in hdr:
            if cols["cargo"] is None: cols["cargo"] = c
        elif any(k in hdr for k in ("DESTINATARIO", "NOMBRES Y APELLIDOS", "APELLIDOS Y NOMBRES", "NOMBRE COMPLETO")) and "OFICINA" not in hdr and "REMITENTE" not in hdr and "CARGO" not in hdr:
            if cols["destinatario"] is None: cols["destinatario"] = c
        elif any(k in hdr for k in ("ADJUNTAR", "ITEM", "ANEXO", "ARCHIVO PARA ADJUNTAR")):
            if cols["item_anexo"] is None: cols["item_anexo"] = c
        elif any(k in hdr for k in ("CORREO", "EMAIL", "E-MAIL", "URL")) and "NOTIFICACIONES" not in hdr and "ADJUNTAR" not in hdr and "ITEM" not in hdr and "ANEXO" not in hdr:
            if cols["correo"] is None: cols["correo"] = c
        elif any(k in hdr for k in ("TELEFONO", "TELÉFONO", "CELULAR")):
            if cols["telefono"] is None: cols["telefono"] = c
        elif "AGOTAMIENTO" in hdr:
            if cols["agotamiento"] is None: cols["agotamiento"] = c
        elif "EXPEDIENTE" in hdr and any(k in hdr for k in ("BECARIO", "PADRON", "PADRÓN", "SIBEC")):
            if cols["expediente"] is None: cols["expediente"] = c

    return cols



def get_salida_dir() -> Path:
    base = get_base_dir()
    salida = base / "Informes_Generados"
    salida.mkdir(parents=True, exist_ok=True)
    return salida


def asignar_valor_celda(ws, r: int, c: int, valor: Any) -> None:
    """Asigna valor a una celda de manera segura, resolviendo la celda raíz si es MergedCell."""
    if r is None or c is None:
        return
    try:
        cell = ws.cell(row=r, column=c)
        if isinstance(cell, MergedCell) or type(cell).__name__ == "MergedCell":
            for rng in ws.merged_cells.ranges:
                if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                    ws.cell(row=rng.min_row, column=rng.min_col).value = valor
                    return
            return
        cell.value = valor
    except AttributeError:
        # Si openpyxl considera la celda de solo lectura (MergedCell no detectada)
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                try:
                    ws.cell(row=rng.min_row, column=rng.min_col).value = valor
                except Exception:
                    pass
                return
    except Exception:
        pass


def obtener_valor_celda(ws, r: int, c: int) -> Any:
    """Obtiene el valor de una celda, buscando en la celda raíz si es una MergedCell."""
    if r is None or c is None:
        return None
    try:
        cell = ws.cell(row=r, column=c)
        if isinstance(cell, MergedCell) or type(cell).__name__ == "MergedCell":
            for rng in ws.merged_cells.ranges:
                if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                    return ws.cell(row=rng.min_row, column=rng.min_col).value
            return None
        return cell.value
    except Exception:
        return None


class GeneradorExcel:
    """Clase para la generación de la Notificación Excel a partir de su plantilla."""

    def __init__(self):
        self.ruta_plantilla = None

    def generar(self, contexto: dict, log_callback=None) -> Path:
        def log(msg):
            if log_callback:
                log_callback(msg)

        es_casilla = bool(contexto.get("AUTORIZA_CASILLA", False))
        nombre_plantilla = "plantilla_notificacion_casilla.xlsx" if es_casilla else "plantilla_notificacion.xlsx"
        self.ruta_plantilla = obtener_ruta_plantilla_excel(nombre_plantilla)

        if not self.ruta_plantilla.exists():
            # Fallback a la otra plantilla si no se encuentra la elegida
            alt_nombre = "plantilla_notificacion.xlsx" if es_casilla else "plantilla_notificacion_casilla.xlsx"
            alt_ruta = obtener_ruta_plantilla_excel(alt_nombre)
            if alt_ruta.exists():
                log(f"Aviso: Plantilla {nombre_plantilla} no encontrada. Usando alternativa {alt_nombre}.")
                self.ruta_plantilla = alt_ruta
            else:
                log(f"No se encontró la plantilla de notificación en {self.ruta_plantilla}")
                raise FileNotFoundError(f"Falta plantilla: {self.ruta_plantilla}")

        log(f"Cargando plantilla de notificación Excel ({self.ruta_plantilla.name})...")
        wb = openpyxl.load_workbook(self.ruta_plantilla)
        ws = wb.active
        
        sigedo_full = str(contexto.get("NUMERO_SIGEDO") or contexto.get("NUMERO_SIGEDO_GLOBAL") or "").strip()
        sigedo_corto = str(contexto.get("SIGEDO_CORTO") or "").strip()
        if not sigedo_corto and sigedo_full:
            sigedo_corto = sigedo_full.split("-")[0].strip()

        dni = str(contexto.get("DNI_VALIDADO", "")).strip()
        dni_val = int(dni) if dni.isdigit() else dni

        n_bec = str(contexto.get("NOMBRES_BECARIO", "")).strip()
        a_bec = str(contexto.get("APELLIDOS_BECARIO", "")).strip()
        if n_bec and a_bec:
            nombres = f"{n_bec} {a_bec}".strip().upper()
        else:
            raw_nom = str(contexto.get("NOMBRE_PRIMERO_NOMBRES") or contexto.get("NOMBRES_Y_APELLIDOS_VALIDADOS", "")).strip().upper()
            if "," in raw_nom:
                partes = [p.strip() for p in raw_nom.split(",", 1)]
                nombres = f"{partes[1]} {partes[0]}".strip().upper()
            else:
                nombres = raw_nom
        if "," in nombres:
            partes = [p.strip() for p in nombres.split(",", 1)]
            nombres = f"{partes[1]} {partes[0]}".strip().upper()

        correo = str(contexto.get("CORREO_ELECTRONICO", "")).strip()
        tel_str = str(contexto.get("TELEFONO_CONTACTO", "")).strip()
        tel_val = int(tel_str) if tel_str.isdigit() else (tel_str or "-")

        num_inf = str(contexto.get("NUMERO_INFORME_GENERAR", "")).strip()
        informe_generado = f"INFORME Nº {num_inf}-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS"

        exp_b = str(contexto.get("EXPEDIENTE_PADRON") or contexto.get("EXPEDIENTE_BECARIO") or contexto.get("EXPEDIENTE") or contexto.get("NUMERO_EXPEDIENTE", "")).strip()
        exp_b = re.sub(r"\.0+$", "", exp_b).strip()
        exp_b_val = int(exp_b) if exp_b.isdigit() else exp_b

        semestre = str(contexto.get("SEMESTRE_SOLICITADO", "2026-II")).strip()
        asunto_notif = str(contexto.get("ASUNTO_OFICIO") or f"Ampliación excepcional de la duración del periodo de estudios por el semestre {semestre}").strip()
        oficio_doc = "OFICIO N° xxx-2026-MINEDU/VMGI-PRONABEC-DIBEC"

        # 1. Determinación de la fila de datos y mapeo inteligente de columnas
        col_dni = None
        col_dest = None
        col_correo = None
        col_sig = None
        col_tel = None
        col_exp = None
        col_asunto = None
        col_doc = None
        col_msj = None
        col_adj = None

        ultima_fila_header = 2 if es_casilla else 3

        # Mapear columnas combinando textos de todas las filas de encabezado (para columnas con celdas combinadas)
        for c in range(1, ws.max_column + 1):
            hdr_comb = " ".join(str(ws.cell(row=r, column=c).value or "").strip().upper() for r in range(1, ultima_fila_header + 1))
            if not hdr_comb:
                continue

            if ("DNI" in hdr_comb or "D.N.I." in hdr_comb or "DOCUMENTO DE IDENTIDAD" in hdr_comb) and "TIPO" not in hdr_comb:
                if col_dni is None: col_dni = c
            elif any(w in hdr_comb for w in ("DESTINATARIO", "NOMBRES Y APELLIDOS", "APELLIDOS Y NOMBRES", "NOMBRE DEL BECARIO", "NOMBRE COMPLETO")) and "OFICINA" not in hdr_comb and "NOTIFICACIONES" not in hdr_comb:
                if col_dest is None: col_dest = c
            elif any(w in hdr_comb for w in ("CORREO ELECTRÓNICO", "CORREO ELECTRONICO", "CORREO", "EMAIL", "E-MAIL")) and "NOTIFICACIONES" not in hdr_comb:
                if col_correo is None: col_correo = c
            elif "SIGEDO" in hdr_comb or ("EXPEDIENTE" in hdr_comb and "BECARIO" not in hdr_comb and "PADRON" not in hdr_comb):
                if col_sig is None: col_sig = c
            elif any(w in hdr_comb for w in ("TELÉFONO", "TELEFONO", "CELULAR")):
                if col_tel is None: col_tel = c
            elif "EXPEDIENTE" in hdr_comb and any(w in hdr_comb for w in ("BECARIO", "PADRON", "SIBEC")):
                if col_exp is None: col_exp = c
            elif any(w in hdr_comb for w in ("TIPIFICACIÓN", "TIPIFICACION", "ASUNTO")):
                if col_asunto is None: col_asunto = c
            elif any(w in hdr_comb for w in ("N° DE DOCUMENTO", "NUMERO DE DOCUMENTO", "DOCUMENTO EMITIDO")):
                if col_doc is None: col_doc = c
            elif any(w in hdr_comb for w in ("DESCRIPCIÓN", "DESCRIPCION", "MENSAJE DE EMISIÓN", "MENSAJE DE EMISION")):
                if col_msj is None: col_msj = c
            elif any(w in hdr_comb for w in ("ADJUNTAR EN EL CORREO", "NUMERO DE (ITEM)", "ARCHIVO PARA ADJUNTAR")):
                if col_adj is None: col_adj = c

        # Valores de respaldo por defecto si algún encabezado no se reconoció
        if es_casilla:
            fila_datos = 3
            if not col_dni: col_dni = 2
            if not col_dest: col_dest = 3
            if not col_exp: col_exp = 4
            if not col_sig: col_sig = 5
            if not col_adj: col_adj = 8
            if not col_tel: col_tel = 11

            if col_dni: asignar_valor_celda(ws, fila_datos, col_dni, dni_val)
            if col_dest: asignar_valor_celda(ws, fila_datos, col_dest, nombres)
            if col_exp: asignar_valor_celda(ws, fila_datos, col_exp, exp_b_val)
            if col_sig: asignar_valor_celda(ws, fila_datos, col_sig, sigedo_full or sigedo_corto)
            if col_tel: asignar_valor_celda(ws, fila_datos, col_tel, tel_val)
            if col_adj:
                desc_doc_base = str(ws.cell(row=fila_datos, column=col_adj).value or "")
                patron_doc = re.compile(r"INFORME N[°ºo\.]?\s*\d*-2026-MINEDU/VMGI-PRONABEC-DIBEC(?:-SUS)?", re.IGNORECASE)
                if patron_doc.search(desc_doc_base):
                    asignar_valor_celda(ws, fila_datos, col_adj, patron_doc.sub(informe_generado, desc_doc_base))
                else:
                    asignar_valor_celda(ws, fila_datos, col_adj, f"OFICIO Nº xxx-2026-MINEDU/VMGI-PRONABEC-DIBEC, {informe_generado}")

            nombre_salida = f"{sigedo_corto or dni or 'notificacion'}_notificacion_casilla.xlsx"
        else:
            if not col_sig: col_sig = 2
            if not col_dni: col_dni = 5
            if not col_dest: col_dest = 7
            if not col_asunto: col_asunto = 8
            if not col_doc: col_doc = 9
            if not col_msj: col_msj = 10
            if not col_correo: col_correo = 12
            if not col_tel: col_tel = 13
            if not col_adj: col_adj = 14

            # En plantilla_notificacion.xlsx (notificación individual por correo):
            # Las filas 4 y 5 corresponden a la notificación (con celdas combinadas verticalmente en varias columnas).
            # Para el campo 'NUMERO DE (ITEM) DEL ARCHIVO PARA ADJUNTAR EN EL CORREO' (col_adj):
            # - Fila 4: Item I (El Informe SUS)
            # - Fila 5: Item II (El Oficio)
            for f in (4, 5):
                if col_sig: asignar_valor_celda(ws, f, col_sig, sigedo_full or sigedo_corto)
                if col_dni: asignar_valor_celda(ws, f, col_dni, dni_val)
                if col_dest: asignar_valor_celda(ws, f, col_dest, nombres)
                if col_asunto: asignar_valor_celda(ws, f, col_asunto, asunto_notif)
                if col_doc: asignar_valor_celda(ws, f, col_doc, oficio_doc)
                if col_msj: asignar_valor_celda(ws, f, col_msj, asunto_notif)
                if col_correo: asignar_valor_celda(ws, f, col_correo, correo)
                if col_tel: asignar_valor_celda(ws, f, col_tel, tel_val)

            if col_adj:
                asignar_valor_celda(ws, 4, col_adj, f"I) {informe_generado}")
                asignar_valor_celda(ws, 5, col_adj, f"II) {oficio_doc}")

            nombre_salida = f"{sigedo_corto or dni or 'notificacion'}_notificacion.xlsx"

        # 2. Reemplazo universal de marcadores y datos residuales de plantilla en todas las celdas
        marcadores = {
            "{{DNI}}": str(dni),
            "{DNI}": str(dni),
            "{{DESTINATARIO}}": nombres,
            "{DESTINATARIO}": nombres,
            "{{NOMBRES}}": nombres,
            "{NOMBRES}": nombres,
            "{{BECARIO}}": nombres,
            "{BECARIO}": nombres,
            "{{CORREO}}": correo,
            "{CORREO}": correo,
            "{{CORREO_ELECTRONICO}}": correo,
            "{CORREO_ELECTRONICO}": correo,
            "{{EMAIL}}": correo,
            "{EMAIL}": correo,
            "{{SIGEDO}}": str(sigedo_full or sigedo_corto),
            "{SIGEDO}": str(sigedo_full or sigedo_corto),
            "{{INFORME}}": f"INFORME Nº {num_inf}",
            "{INFORME}": f"INFORME Nº {num_inf}",
            "{{OFICIO}}": oficio_doc,
            "{OFICIO}": oficio_doc,
            "{{EXPEDIENTE}}": str(exp_b),
            "{EXPEDIENTE}": str(exp_b),
            "{{TELEFONO}}": str(tel_val),
            "{TELEFONO}": str(tel_val),
            # Residuos de ejemplo de la plantilla original
            "78069508": str(dni),
            "DIEGO CARBAJAL CCORAHUA": nombres,
            "diegocdacc864@gmail.com": correo,
        }

        patron_inf = re.compile(r"Informe\s+N[°ºo\.]?\s*\d*-2026-MINEDU/VMGI-PRONABEC-DIBEC(?:-SUS)?", re.IGNORECASE)
        patron_inf_generico = re.compile(r"Informe\s+N[°ºo\.]?\s*\d+", re.IGNORECASE)

        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell) or type(cell).__name__ == "MergedCell":
                    continue
                if cell.value is None:
                    continue

                try:
                    texto_original = str(cell.value)
                    
                    # A. Reemplazo de marcadores y residuos conocidos
                    for tag, replacement in marcadores.items():
                        if tag in texto_original and replacement:
                            if texto_original.strip() == tag:
                                cell.value = int(replacement) if str(replacement).isdigit() and tag in ("{{DNI}}", "{DNI}", "78069508") else replacement
                                texto_original = str(cell.value)
                            else:
                                texto_original = texto_original.replace(tag, str(replacement))
                                cell.value = texto_original

                    # B. Reemplazo de SIGEDO residual
                    if "59583" in texto_original or "60692" in texto_original:
                        texto_a_reemplazar = "59583" if "59583" in texto_original else "60692"
                        if isinstance(cell.value, (int, float)):
                            solo_nums = re.sub(r"\D", "", str(sigedo_corto))
                            cell.value = int(solo_nums) if solo_nums else (int(sigedo_corto) if sigedo_corto.isdigit() else sigedo_corto)
                        else:
                            cell.value = texto_original.replace(texto_a_reemplazar, str(sigedo_corto))

                    # C. Reemplazo de número de informe en celdas numéricas aisladas
                    if (str(cell.value).strip() in ("6539", "6757", "6541")) and isinstance(cell.value, (int, float)):
                        cell.value = int(num_inf) if num_inf.isdigit() else num_inf

                    # D. Reemplazo de cadenas de Informe
                    texto_actual = str(cell.value)
                    if patron_inf.search(texto_actual):
                        cell.value = patron_inf.sub(informe_generado, texto_actual)
                    elif patron_inf_generico.search(texto_actual) and "INFORME" in texto_actual.upper():
                        cell.value = patron_inf_generico.sub(f"Informe Nº {num_inf}", texto_actual)
                except (AttributeError, TypeError):
                    pass


        salida_dir = get_salida_dir()
        ruta_salida = salida_dir / nombre_salida
        
        log(f"Guardando Notificación en: {ruta_salida.name}")
        wb.save(ruta_salida)
        return ruta_salida

    def generar_acto_administrativo(self, contexto: dict, log_callback=None) -> Path:
        """
        Genera la Notificación de Acto Administrativo para Informes No Viables
        utilizando plantilla_notificacion_acto_administrativo.xlsx.
        Maneja pestañas: NOT.CORREO, NOT.CASILLA y NOT.PERSONAL según el canal indicado
        en el formato autogenerado (correo vs casilla electrónica).
        """
        def log(msg):
            if log_callback:
                log_callback(msg)

        self.ruta_plantilla = obtener_ruta_plantilla_acto_administrativo()
        if not self.ruta_plantilla.exists():
            log(f"No se encontró la plantilla de acto administrativo en {self.ruta_plantilla}")
            raise FileNotFoundError(f"Falta plantilla: {self.ruta_plantilla}")

        log(f"Cargando plantilla de notificación acto administrativo ({self.ruta_plantilla.name})...")
        wb = openpyxl.load_workbook(self.ruta_plantilla)

        # 1. Identificar las tres pestañas de la plantilla
        ws_correo = None
        ws_casilla = None
        ws_personal = None
        for sheet in wb.worksheets:
            t = re.sub(r"[\s\.\-_]", "", sheet.title).upper()
            if "CORREO" in t:
                ws_correo = sheet
            elif "CASILLA" in t:
                ws_casilla = sheet
            elif "PERSONAL" in t:
                ws_personal = sheet

        if not ws_correo and len(wb.worksheets) >= 1:
            ws_correo = wb.worksheets[0]
        if not ws_casilla and len(wb.worksheets) >= 2:
            ws_casilla = wb.worksheets[1]
        if not ws_personal and len(wb.worksheets) >= 3:
            ws_personal = wb.worksheets[2]

        # 2. Extracción y preparación de datos del contexto
        sigedo_full = str(contexto.get("NUMERO_SIGEDO") or contexto.get("NUMERO_SIGEDO_GLOBAL") or "").strip()
        sigedo_corto = str(contexto.get("SIGEDO_CORTO") or "").strip()
        if not sigedo_corto and sigedo_full:
            sigedo_corto = sigedo_full.split("-")[0].strip()

        # En Notificación (Casilla y Correo), SIGEDO N° debe incluir el sufijo de año (ej. "66274 - 2026")
        try:
            from datetime import date
            anio_actual = date.today().year
        except Exception:
            anio_actual = 2026

        if sigedo_full and "-" in sigedo_full:
            partes_sig = [p.strip() for p in sigedo_full.split("-", 1)]
            sigedo_val = f"{partes_sig[0]} - {partes_sig[1]}"
        elif sigedo_corto:
            sigedo_val = f"{sigedo_corto} - {anio_actual}"
        else:
            sigedo_val = sigedo_full or ""

        sexo_str = str(contexto.get("SEXO") or "").strip().upper()
        es_fem = sexo_str.startswith("F") or sexo_str in ("2", "MUJER", "FEMENINO")
        tipo_usuario = "CIUDADANA" if es_fem else "CIUDADANO"

        dni = str(contexto.get("DNI_VALIDADO", "")).strip()
        dni_val = int(dni) if dni.isdigit() else dni

        n_bec = str(contexto.get("NOMBRES_BECARIO", "")).strip()
        a_bec = str(contexto.get("APELLIDOS_BECARIO", "")).strip()
        if n_bec and a_bec:
            nombres = f"{n_bec} {a_bec}".strip().upper()
        else:
            raw_nom = str(contexto.get("NOMBRE_PRIMERO_NOMBRES") or contexto.get("NOMBRES_Y_APELLIDOS_VALIDADOS", "")).strip().upper()
            if "," in raw_nom:
                partes = [p.strip() for p in raw_nom.split(",", 1)]
                nombres = f"{partes[1]} {partes[0]}".strip().upper()
            else:
                nombres = raw_nom
        if "," in nombres:
            partes = [p.strip() for p in nombres.split(",", 1)]
            nombres = f"{partes[1]} {partes[0]}".strip().upper()

        correo_fmt = str(
            contexto.get("CORREO_FORMATO_AUTOGENERADO")
            or contexto.get("CORREO_ELECTRONICO")
            or contexto.get("CORREO")
            or contexto.get("EMAIL")
            or contexto.get("CORREO_PADRON")
            or ""
        ).strip()
        correo = correo_fmt

        tel_str = str(contexto.get("TELEFONO_FORMATO_AUTOGENERADO") or contexto.get("TELEFONO_CONTACTO") or "").strip()
        tel_val = int(tel_str) if tel_str.isdigit() else (tel_str or "-")

        num_inf = str(contexto.get("NUMERO_INFORME_GENERAR") or "").strip()
        if not num_inf:
            m_num = re.search(r"\d+", str(contexto.get("NUMERO_INFORME", "")))
            num_inf = m_num.group(0) if m_num else ""

        fecha_sol_texto = str(contexto.get("FECHA_SOLICITUD_TEXTO") or contexto.get("FECHA_SOLICITUD") or "").strip()
        if not fecha_sol_texto or fecha_sol_texto in ("(no detectada)", ""):
            try:
                from datetime import date
                from procesador import fecha_a_texto
                fecha_sol_texto = fecha_a_texto(date.today())
            except Exception:
                fecha_sol_texto = "21 de julio de 2026"

        exp_b = str(contexto.get("EXPEDIENTE_PADRON") or contexto.get("EXPEDIENTE_BECARIO") or contexto.get("EXPEDIENTE") or "").strip()
        exp_b = re.sub(r"\.0+$", "", exp_b).strip()
        exp_b_val = int(exp_b) if exp_b.isdigit() else exp_b

        # 3. Determinar si señala notificación por casilla electrónica o por correo
        es_casilla = bool(contexto.get("AUTORIZA_CASILLA_FORMATO", contexto.get("AUTORIZA_CASILLA", False)))

        if es_casilla:
            log("Formato autogenerado indica notificación por Casilla Electrónica. Invirtiendo orden de pestañas...")
            agotamiento_casilla = "SI"
            agotamiento_correo = "NO"
            hojas_ordenadas = [ws_casilla, ws_correo, ws_personal]
        else:
            log("Formato autogenerado indica notificación por Correo Electrónico. Manteniendo orden estándar de pestañas...")
            agotamiento_correo = "SI" if correo else "NO"
            agotamiento_casilla = "NO"
            hojas_ordenadas = [ws_correo, ws_casilla, ws_personal]

        # Reordenar pestañas en el libro
        hojas_validas = [h for h in hojas_ordenadas if h is not None]
        otras_hojas = [h for h in wb.worksheets if h not in hojas_validas]
        wb._sheets = hojas_validas + otras_hojas
        wb.active = hojas_validas[0]

        # 4. Llenar pestaña NOT.CORREO
        if ws_correo:
            fila_hdr_c, fila_correo = _detectar_fila_encabezado_y_datos(ws_correo)
            cols_c = _mapear_columnas_acto(ws_correo, fila_hdr_c)

            # Fallbacks exactos para la estructura de la plantilla NOT.CORREO
            if not cols_c["sigedo"]: cols_c["sigedo"] = 2
            if not cols_c["tipo_usuario"]: cols_c["tipo_usuario"] = 5
            if not cols_c["dni"]: cols_c["dni"] = 6
            if not cols_c["destinatario"]: cols_c["destinatario"] = 7
            if not cols_c["correo"]: cols_c["correo"] = 13
            if not cols_c["telefono"]: cols_c["telefono"] = 15
            if not cols_c["item_anexo"]: cols_c["item_anexo"] = 16
            if not cols_c["agotamiento"]: cols_c["agotamiento"] = 18

            if cols_c["sigedo"]: asignar_valor_celda(ws_correo, fila_correo, cols_c["sigedo"], sigedo_val)
            if cols_c["tipo_usuario"]: asignar_valor_celda(ws_correo, fila_correo, cols_c["tipo_usuario"], tipo_usuario)
            if cols_c["dni"]: asignar_valor_celda(ws_correo, fila_correo, cols_c["dni"], dni_val)
            # En NOT.CORREO la Col G (7) es DESTINATARIO (Nombres y Apellidos), nunca el número de expediente
            if cols_c["destinatario"]: asignar_valor_celda(ws_correo, fila_correo, cols_c["destinatario"], nombres)
            if cols_c["correo"]: asignar_valor_celda(ws_correo, fila_correo, cols_c["correo"], correo)
            if cols_c["telefono"]: asignar_valor_celda(ws_correo, fila_correo, cols_c["telefono"], tel_val)
            if cols_c["item_anexo"]:
                val_orig = obtener_valor_celda(ws_correo, fila_correo, cols_c["item_anexo"])
                val_act = actualizar_texto_items_anexo(str(val_orig or ""), num_inf, fecha_sol_texto)
                asignar_valor_celda(ws_correo, fila_correo, cols_c["item_anexo"], val_act)
            if cols_c["agotamiento"]: asignar_valor_celda(ws_correo, fila_correo, cols_c["agotamiento"], agotamiento_correo)

        # 5. Llenar pestaña NOT.CASILLA
        if ws_casilla:
            fila_hdr_cas, fila_casilla = _detectar_fila_encabezado_y_datos(ws_casilla)
            cols_cas = _mapear_columnas_acto(ws_casilla, fila_hdr_cas)

            # Fallbacks exactos para la estructura de la plantilla NOT.CASILLA
            if not cols_cas["sigedo"]: cols_cas["sigedo"] = 2
            if not cols_cas["tipo_usuario"]: cols_cas["tipo_usuario"] = 5
            if not cols_cas["dni"]: cols_cas["dni"] = 6
            if not cols_cas["expediente"]: cols_cas["expediente"] = 7
            if not cols_cas["destinatario"]: cols_cas["destinatario"] = 8
            if not cols_cas["telefono"]: cols_cas["telefono"] = 12
            if not cols_cas["agotamiento"]: cols_cas["agotamiento"] = 13
            if not cols_cas["item_anexo"]: cols_cas["item_anexo"] = 14

            if cols_cas["sigedo"]: asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["sigedo"], sigedo_val)
            if cols_cas["tipo_usuario"]: asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["tipo_usuario"], tipo_usuario)
            if cols_cas["dni"]: asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["dni"], dni_val)
            if cols_cas["expediente"]: asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["expediente"], exp_b_val)
            if cols_cas["destinatario"]: asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["destinatario"], nombres)
            if cols_cas["item_anexo"]:
                val_orig_cas = obtener_valor_celda(ws_casilla, fila_casilla, cols_cas["item_anexo"])
                val_act_cas = actualizar_texto_items_anexo(str(val_orig_cas or ""), num_inf, fecha_sol_texto)
                asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["item_anexo"], val_act_cas)
            if cols_cas["telefono"]: asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["telefono"], tel_val)
            if cols_cas["agotamiento"]: asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["agotamiento"], agotamiento_casilla)

        # 6. Reemplazo de marcadores genéricos en celdas de NOT.CORREO y NOT.CASILLA (NOT.PERSONAL se mantiene intacta)
        nuevo_inf_tag = f"INFORME Nº {num_inf}-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS"
        marcadores_acto = {
            "{{DNI}}": str(dni),
            "{DNI}": str(dni),
            "{{DESTINATARIO}}": nombres,
            "{DESTINATARIO}": nombres,
            "{{NOMBRES}}": nombres,
            "{NOMBRES}": nombres,
            "{{BENEFICIARIO}}": nombres,
            "{BENEFICIARIO}": nombres,
            "{{BECARIO}}": nombres,
            "{BECARIO}": nombres,
            "{{SIGEDO}}": str(sigedo_val),
            "{SIGEDO}": str(sigedo_val),
            "{{EXPEDIENTE}}": str(exp_b),
            "{EXPEDIENTE}": str(exp_b),
            "{{EXPEDIENTE_BECARIO}}": str(exp_b),
            "{EXPEDIENTE_BECARIO}": str(exp_b),
            "{{CORREO}}": correo,
            "{CORREO}": correo,
            "{{EMAIL}}": correo,
            "{EMAIL}": correo,
            "{{CORREO_ELECTRONICO}}": correo,
            "{CORREO_ELECTRONICO}": correo,
            "periquitopinpin@gmail.com": correo,
            "diegocdacc864@gmail.com": correo,
            "{{TELEFONO}}": str(tel_val),
            "{TELEFONO}": str(tel_val),
            "{{CELULAR}}": str(tel_val),
            "{CELULAR}": str(tel_val),
            "{{TIPO_USUARIO}}": tipo_usuario,
            "{TIPO_USUARIO}": tipo_usuario,
            "{{INFORME}}": nuevo_inf_tag,
            "{INFORME}": nuevo_inf_tag,
            "{{INFORME_SUS}}": nuevo_inf_tag,
            "{INFORME_SUS}": nuevo_inf_tag,
            "{{FECHA_SOLICITUD}}": fecha_sol_texto,
            "{FECHA_SOLICITUD}": fecha_sol_texto,
            "{{FECHA}}": fecha_sol_texto,
            "{FECHA}": fecha_sol_texto,
        }
        for ws_target in (ws_correo, ws_casilla):
            if not ws_target:
                continue
            for row in ws_target.iter_rows():
                for cell in row:
                    if isinstance(cell, MergedCell) or type(cell).__name__ == "MergedCell":
                        continue
                    if cell.value is None:
                        continue
                    try:
                        t_orig = str(cell.value)
                        for tag, rep in marcadores_acto.items():
                            if tag.lower() in t_orig.lower() and rep:
                                patron_tag = re.compile(re.escape(tag), re.IGNORECASE)
                                if t_orig.strip().upper() == tag.upper():
                                    cell.value = int(rep) if str(rep).isdigit() and "DNI" in tag.upper() else rep
                                    t_orig = str(cell.value)
                                else:
                                    t_orig = patron_tag.sub(str(rep), t_orig)
                                    cell.value = t_orig
                    except Exception:
                        pass

        # Salvaguarda final: En NOT.CORREO la Col 7 (DESTINATARIO) nunca debe retener
        # el número de expediente por plantilla, y correo/anexo deben quedar garantizados
        if ws_correo:
            col_dest_c = cols_c.get("destinatario") or 7
            asignar_valor_celda(ws_correo, fila_correo, col_dest_c, nombres)
            if cols_c.get("sigedo"):
                asignar_valor_celda(ws_correo, fila_correo, cols_c["sigedo"], sigedo_val)
            if cols_c.get("correo") and correo:
                asignar_valor_celda(ws_correo, fila_correo, cols_c["correo"], correo)
            if cols_c.get("item_anexo"):
                val_c_final = obtener_valor_celda(ws_correo, fila_correo, cols_c["item_anexo"])
                val_c_act = actualizar_texto_items_anexo(str(val_c_final or ""), num_inf, fecha_sol_texto)
                asignar_valor_celda(ws_correo, fila_correo, cols_c["item_anexo"], val_c_act)

        if ws_casilla:
            if cols_cas.get("sigedo"):
                asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["sigedo"], sigedo_val)
            if cols_cas.get("destinatario"):
                asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["destinatario"], nombres)
            if cols_cas.get("expediente"):
                asignar_valor_celda(ws_casilla, fila_casilla, cols_cas["expediente"], exp_b_val)

        # 7. Guardar archivo generado
        nombre_salida = f"{sigedo_corto or dni or 'notificacion'}_notificacion_acto_administrativo.xlsx"
        salida_dir = get_salida_dir()
        ruta_salida = salida_dir / nombre_salida

        log(f"Guardando Notificación Acto Administrativo en: {ruta_salida.name}")
        wb.save(ruta_salida)
        return ruta_salida


    def generar_multiple(self, super_contexto: dict, log_callback=None) -> list[Path]:
        def log(msg):
            if log_callback: log_callback(msg)

        becarios = super_contexto.get("becarios", [])
        becarios_casilla = [b for b in becarios if b.get("AUTORIZA_CASILLA")]
        becarios_normales = [b for b in becarios if not b.get("AUTORIZA_CASILLA")]

        rutas_generadas = []

        if becarios_casilla:
            log(f"Se detectaron {len(becarios_casilla)} becario(s) con autorización de casilla electrónica.")
            ruta_cas = self.generar_casilla_multiple(super_contexto, becarios_casilla, log_callback=log_callback)
            if ruta_cas:
                rutas_generadas.append(ruta_cas)

        if becarios_normales:
            log(f"Se detectaron {len(becarios_normales)} becario(s) para notificación por correo electrónico.")
            ruta_norm = self.generar_correo_multiple(super_contexto, becarios_normales, log_callback=log_callback)
            if ruta_norm:
                rutas_generadas.append(ruta_norm)

        return rutas_generadas

    def generar_casilla_multiple(self, super_contexto: dict, becarios: list, log_callback=None) -> Path:
        def log(msg):
            if log_callback: log_callback(msg)

        ruta_plantilla = obtener_ruta_plantilla_excel("plantilla_notificacion_casilla_multiple.xlsx")
        if not ruta_plantilla.exists():
            ruta_plantilla = obtener_ruta_plantilla_excel("plantilla_notificacion_casilla.xlsx")

        if not ruta_plantilla.exists():
            log(f"No se encontró la plantilla de casilla múltiple en {ruta_plantilla}")
            raise FileNotFoundError(f"Falta plantilla: {ruta_plantilla}")

        log("Cargando plantilla de notificación Excel casilla múltiple...")
        wb = openpyxl.load_workbook(ruta_plantilla)
        ws = wb.active

        fila_base = 3
        valores_base = [ws.cell(row=fila_base, column=c).value for c in range(1, 12)]

        filas_necesarias = len(becarios)
        ultima_fila_requerida = fila_base + max(filas_necesarias, 1) - 1
        for r in range(ws.max_row, ultima_fila_requerida, -1):
            ws.delete_rows(r)

        num_inf = super_contexto.get("NUMERO_INFORME_GENERAR", "")
        informe_generado = f"INFORME Nº {num_inf}-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS"
        sigedo_global = super_contexto.get("NUMERO_SIGEDO_GLOBAL", "")

        def clonar_estilo(origen, destino):
            if origen.has_style:
                destino.font = copy(origen.font)
                destino.border = copy(origen.border)
                destino.fill = copy(origen.fill)
                destino.number_format = copy(origen.number_format)
                destino.protection = copy(origen.protection)
                destino.alignment = copy(origen.alignment)

        base_height = ws.row_dimensions[fila_base].height or 63.75

        # Col 8 (H): N° DE DOCUMENTO
        desc_doc_base = str(valores_base[7] or "")
        patron_doc = re.compile(r"INFORME N[°ºo\.]?\s*\d*-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS", re.IGNORECASE)
        if patron_doc.search(desc_doc_base):
            col_h_mod = patron_doc.sub(informe_generado, desc_doc_base)
        else:
            col_h_mod = re.sub(r"INFORME N[°ºo\.]?\s*\d+", f"INFORME Nº {num_inf}", desc_doc_base, flags=re.IGNORECASE)
            if col_h_mod == desc_doc_base and "INFORME" in desc_doc_base.upper():
                col_h_mod = f"OFICIO Nº xxx-2026-MINEDU/VMGI-PRONABEC-DIBEC, {informe_generado}"

        for i, becario in enumerate(becarios):
            fila_actual = fila_base + i
            ws.row_dimensions[fila_actual].height = base_height

            dni = str(becario.get("DNI_VALIDADO", "")).strip()
            dni_val = int(dni) if dni.isdigit() else dni

            n_bec = str(becario.get("NOMBRES_BECARIO", "")).strip()
            a_bec = str(becario.get("APELLIDOS_BECARIO", "")).strip()
            nombres = f"{n_bec} {a_bec}".strip()
            if not nombres:
                nombres = str(becario.get("NOMBRES_Y_APELLIDOS_VALIDADOS", "")).upper()

            exp_bec = becario.get("EXPEDIENTE_PADRON") or becario.get("EXPEDIENTE_BECARIO") or becario.get("EXPEDIENTE") or becario.get("NUMERO_EXPEDIENTE", "")
            exp_bec_str = re.sub(r"\.0+$", "", str(exp_bec or "")).strip()
            exp_bec_val = int(exp_bec_str) if exp_bec_str.isdigit() else exp_bec_str

            sigedo_val = super_contexto.get("NUMERO_SIGEDO_GLOBAL") or becario.get("NUMERO_SIGEDO", "")

            tel = str(becario.get("TELEFONO_CONTACTO", "")).strip() or "-"
            tel_val = int(tel) if tel.isdigit() else tel

            valores_fila = [
                i + 1,                               # 1: N°
                dni_val,                             # 2: DNI
                nombres,                             # 3: NOMBRE COMPLETO
                exp_bec_val,                         # 4: N° EXPEDIENTE DE BECARIO
                sigedo_val,                          # 5: SIGEDO
                valores_base[5],                     # 6: OFICINA/UNIDAD
                valores_base[6],                     # 7: TIPIFICACION - TITULO
                col_h_mod,                           # 8: N° DE DOCUMENTO
                valores_base[8],                     # 9: DESCRIPCIÓN
                valores_base[9],                     # 10: OBSERVACIONES
                tel_val,                             # 11: TELEFONO
            ]

            for c in range(1, 12):
                target_cell = ws.cell(row=fila_actual, column=c)
                ref_cell = ws.cell(row=fila_base, column=c)
                if fila_actual != fila_base and not isinstance(target_cell, MergedCell) and type(target_cell).__name__ != "MergedCell":
                    clonar_estilo(ref_cell, target_cell)
                asignar_valor_celda(ws, fila_actual, c, valores_fila[c - 1])

        nombre_salida = f"{sigedo_global}_notificacion_casilla_multiple.xlsx"
        ruta_salida = get_salida_dir() / nombre_salida

        log(f"Guardando Notificación Casilla Múltiple en: {ruta_salida.name}")
        wb.save(ruta_salida)
        return ruta_salida

    def generar_correo_multiple(self, super_contexto: dict, becarios: list, log_callback=None) -> Path:
        def log(msg):
            if log_callback: log_callback(msg)

        ruta_plantilla = obtener_ruta_plantilla_excel("plantilla_notificacion_multiple.xlsx")
        if not ruta_plantilla.exists():
            ruta_plantilla = obtener_ruta_plantilla_excel("plantilla_notificacion.xlsx")

        if not ruta_plantilla.exists():
            log(f"No se encontró la plantilla múltiple en {ruta_plantilla}")
            raise FileNotFoundError(f"Falta plantilla: {ruta_plantilla}")

        log("Cargando plantilla de notificación Excel múltiple...")
        wb = openpyxl.load_workbook(ruta_plantilla)
        ws = wb.active

        # Detección inteligente de fila_base en caso de encabezados combinados (fila 3/4 -> datos en 5)
        ultima_hdr = 2
        for r in range(1, 6):
            fila_vals = [str(ws.cell(row=r, column=c).value or "").strip().upper() for c in range(1, ws.max_column + 1)]
            if any(any(k in v for k in ("SIGEDO", "DNI", "DESTINATARIO", "N°", "REMITENTE", "CORREO", "ASUNTO", "DOCUMENTO")) for v in fila_vals):
                ultima_hdr = max(ultima_hdr, r)
        fila_base = ultima_hdr + 1 if ultima_hdr >= 3 else 3
        valores_base = [ws.cell(row=fila_base, column=c).value for c in range(1, max(14, ws.max_column + 1))]

        filas_necesarias = len(becarios)

        # Limpiar filas posteriores de la plantilla que excedan el número de becarios
        ultima_fila_requerida = fila_base + max(filas_necesarias, 1) - 1
        for r in range(ws.max_row, ultima_fila_requerida, -1):
            ws.delete_rows(r)

        num_inf = super_contexto.get("NUMERO_INFORME_GENERAR", "")
        informe_generado = f"Informe Nº {num_inf}-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS"
        sigedo_global = super_contexto.get("NUMERO_SIGEDO_GLOBAL", "")

        def clonar_estilo(origen, destino):
            if origen.has_style:
                destino.font = copy(origen.font)
                destino.border = copy(origen.border)
                destino.fill = copy(origen.fill)
                destino.number_format = copy(origen.number_format)
                destino.protection = copy(origen.protection)
                destino.alignment = copy(origen.alignment)

        base_height = ws.row_dimensions[fila_base].height or 63.75

        # Preparar Col K (NUMERO DE (ITEM) DEL ARCHIVO PARA ADJUNTAR EN EL CORREO)
        raw_col_k = str(valores_base[10] or "")
        patron_col_k = re.compile(r"(Informe[^\d]*?)\d+([^,]*?DIBEC-SUS)", re.IGNORECASE)
        if patron_col_k.search(raw_col_k):
            col_k_mod = patron_col_k.sub(rf"\g<1>{num_inf}\g<2>", raw_col_k)
        elif any(n in raw_col_k for n in ["6757", "6539", "6541"]):
            col_k_mod = re.sub(r"\b(?:6757|6539|6541)\b", str(num_inf), raw_col_k)
        elif raw_col_k:
            col_k_mod = raw_col_k
        else:
            col_k_mod = f"I) {informe_generado}, II) Oficio Nº xxx-2026-MINEDU/VMGI-PRONABEC-DIBEC"

        # Preparar Col H (DESCRIPCIÓN - MENSAJE DE EMISIÓN)
        desc_orig = str(valores_base[7] or "")
        patron_desc = re.compile(r"(Informe[^\d]*?)\d+([^,]*?DIBEC-SUS)", re.IGNORECASE)
        if patron_desc.search(desc_orig):
            desc_mod = patron_desc.sub(rf"\g<1>{num_inf}\g<2>", desc_orig)
        else:
            desc_mod = desc_orig

        for i, becario in enumerate(becarios):
            fila_actual = fila_base + i
            ws.row_dimensions[fila_actual].height = base_height

            dni = str(becario.get("DNI_VALIDADO", "")).strip()
            n_bec = str(becario.get("NOMBRES_BECARIO", "")).strip()
            a_bec = str(becario.get("APELLIDOS_BECARIO", "")).strip()
            if n_bec and a_bec:
                nombres = f"{n_bec} {a_bec}".strip().upper()
            else:
                raw_nom = str(becario.get("NOMBRE_PRIMERO_NOMBRES") or becario.get("NOMBRES_Y_APELLIDOS_VALIDADOS", "")).strip().upper()
                if "," in raw_nom:
                    partes = [p.strip() for p in raw_nom.split(",", 1)]
                    nombres = f"{partes[1]} {partes[0]}".strip().upper()
                else:
                    nombres = raw_nom
            if "," in nombres:
                partes = [p.strip() for p in nombres.split(",", 1)]
                nombres = f"{partes[1]} {partes[0]}".strip().upper()
            correo = str(becario.get("CORREO_ELECTRONICO", "")).strip()
            tel = str(becario.get("TELEFONO_CONTACTO", "")).strip() or "-"

            dni_val = int(dni) if dni.isdigit() else dni

            valores_fila = [
                i + 1,                               # 1: N°
                sigedo_global,                       # 2: SIGEDO
                dni_val,                             # 3: DNI
                valores_base[3],                     # 4: OFICINA REMITENTE
                nombres,                             # 5: DESTINATARIO
                valores_base[5],                     # 6: TIPIFICACIÓN - ASUNTO
                valores_base[6],                     # 7: N° DE DOCUMENTO
                desc_mod,                            # 8: DESCRIPCIÓN - MENSAJE DE EMISIÓN
                correo,                              # 9: CORREO ELECTRÓNICO
                tel,                                 # 10: TELÉFONO
                col_k_mod,                           # 11: NUMERO DE (ITEM) DEL ARCHIVO PARA ADJUNTAR EN EL CORREO
                valores_base[11],                    # 12: OBSERVACIONES
                valores_base[12],                    # 13: TIPO DE NOTIFICACIÓN
            ]

            for c in range(1, 14):
                target_cell = ws.cell(row=fila_actual, column=c)
                ref_cell = ws.cell(row=fila_base, column=c)
                if fila_actual != fila_base and not isinstance(target_cell, MergedCell) and type(target_cell).__name__ != "MergedCell":
                    clonar_estilo(ref_cell, target_cell)
                asignar_valor_celda(ws, fila_actual, c, valores_fila[c - 1])

        nombre_salida = f"{sigedo_global}_notificacion_multiple.xlsx"
        ruta_salida = get_salida_dir() / nombre_salida

        log(f"Guardando Notificación Múltiple en: {ruta_salida.name}")
        wb.save(ruta_salida)
        return ruta_salida
