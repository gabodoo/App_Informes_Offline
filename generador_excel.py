import sys
from pathlib import Path
import openpyxl

def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "main" and exe_dir.parent.name.lower() == "dist":
            return exe_dir.parent.parent
        if exe_dir.name.lower() == "dist":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


class GeneradorExcel:
    """Clase para la generación de la Notificación Excel a partir de su plantilla."""

    def __init__(self):
        self.ruta_plantilla = None

    def generar(self, contexto: dict, log_callback=None) -> Path:
        def log(msg):
            if log_callback:
                log_callback(msg)

        es_casilla = contexto.get("AUTORIZA_CASILLA", False)
        if es_casilla:
            self.ruta_plantilla = get_base_dir() / "plantillas" / "plantilla_notificacion_casilla.xlsx"
        else:
            self.ruta_plantilla = get_base_dir() / "plantillas" / "plantilla_notificacion.xlsx"

        if not self.ruta_plantilla.exists():
            log(f"No se encontró la plantilla de notificación en {self.ruta_plantilla}")
            raise FileNotFoundError(f"Falta plantilla: {self.ruta_plantilla}")

        log("Cargando plantilla de notificación Excel...")
        wb = openpyxl.load_workbook(self.ruta_plantilla)
        ws = wb.active
        
        sigedo_corto = contexto.get("SIGEDO_CORTO", "")
        dni = contexto.get("DNI_VALIDADO", "")
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
        correo = contexto.get("CORREO_ELECTRONICO", "")
        informe_generado = f"INFORME Nº {contexto.get('NUMERO_INFORME_GENERAR', '')}-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS"

        if es_casilla:
            dni_str = str(dni).strip()
            ws["B3"].value = int(dni_str) if dni_str.isdigit() else dni_str
            ws["C3"].value = nombres
            exp_b = str(contexto.get("EXPEDIENTE_PADRON") or contexto.get("EXPEDIENTE_BECARIO") or contexto.get("EXPEDIENTE") or "").strip()
            exp_b = re.sub(r"\.0+$", "", exp_b).strip()
            ws["D3"].value = int(exp_b) if exp_b.isdigit() else exp_b
            ws["E3"].value = contexto.get("NUMERO_SIGEDO", "")
            tel_str = str(contexto.get("TELEFONO_CONTACTO", "")).strip()
            ws["K3"].value = int(tel_str) if tel_str.isdigit() else tel_str
            
            import re
            texto_doc = str(ws["H3"].value or "")
            texto_doc = re.sub(r"INFORME N[°º]\s*\d+-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS", informe_generado, texto_doc, flags=re.IGNORECASE)
            ws["H3"].value = texto_doc
            
            nombre_salida = f"{sigedo_corto}_notificacion_casilla.xlsx"
        else:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    
                    texto_original = str(cell.value)
                    
                    # Reemplazo SIGEDO (59583)
                    if "59583" in texto_original or "60692" in texto_original:
                        texto_a_reemplazar = "59583" if "59583" in texto_original else "60692"
                        if isinstance(cell.value, (int, float)):
                            import re
                            solo_nums = re.sub(r"\D", "", str(sigedo_corto))
                            cell.value = int(solo_nums) if solo_nums else sigedo_corto
                        else:
                            cell.value = texto_original.replace(texto_a_reemplazar, str(sigedo_corto))
                        
                    # Reemplazo de Informe
                    if "6539" in str(cell.value) or "6757" in str(cell.value):
                        if isinstance(cell.value, (int, float)):
                            num = contexto.get('NUMERO_INFORME_GENERAR', '')
                            cell.value = int(num) if str(num).isdigit() else num
                        else:
                            pass # Será manejado por las siguientes condiciones si es un string largo
                            
                    texto = str(cell.value)
                    
                    if "Informe N° 6539-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS" in texto:
                        cell.value = texto.replace("Informe N° 6539-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS", informe_generado)
                    elif "Informe N° 6757-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS" in texto:
                        cell.value = texto.replace("Informe N° 6757-2026-MINEDU/VMGI-PRONABEC-DIBEC-SUS", informe_generado)
                    elif "Informe N° 6539-2026-MINEDU/VMGI-PRONABEC-DIBEC" in texto:
                        cell.value = texto.replace("Informe N° 6539-2026-MINEDU/VMGI-PRONABEC-DIBEC", informe_generado)
                    elif "Informe N° 6539" in texto:
                        cell.value = texto.replace("Informe N° 6539", informe_generado)
                    elif "Informe N° 6757" in texto:
                        cell.value = texto.replace("Informe N° 6757", informe_generado)
                    
                    # Reemplazos "debajo" de una celda
                    texto_upper = texto.upper()
                    if "DNI" == texto_upper.strip() or "D.N.I." in texto_upper:
                        celda_abajo = ws.cell(row=cell.row + 1, column=cell.column)
                        celda_abajo.value = dni
                    
                    if "DESTINATARIO" == texto_upper.strip() or "DESTINATARIO" in texto_upper and "NOTIFICACIONES" not in texto_upper:
                        celda_abajo = ws.cell(row=cell.row + 1, column=cell.column)
                        celda_abajo.value = nombres
                        
                    if ("CORREO ELECTRÓNICO" in texto_upper or "CORREO ELECTRONICO" in texto_upper) and "NOTIFICACIONES" not in texto_upper:
                        celda_abajo = ws.cell(row=cell.row + 1, column=cell.column)
                        celda_abajo.value = correo

            nombre_salida = f"{sigedo_corto}_notificacion.xlsx"
        ruta_salida = get_base_dir() / "Informes_Generados" / nombre_salida
        ruta_salida.parent.mkdir(parents=True, exist_ok=True)
        
        log(f"Guardando Notificación en: {ruta_salida.name}")
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

        ruta_plantilla = get_base_dir() / "plantillas" / "plantilla_notificacion_casilla_multiple.xlsx"
        if not ruta_plantilla.exists():
            ruta_plantilla = get_base_dir() / "plantillas" / "plantilla_notificacion_casilla.xlsx"

        if not ruta_plantilla.exists():
            log(f"No se encontró la plantilla de casilla múltiple en {ruta_plantilla}")
            raise FileNotFoundError(f"Falta plantilla: {ruta_plantilla}")

        log("Cargando plantilla de notificación Excel casilla múltiple...")
        from copy import copy
        import re

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
                if fila_actual != fila_base:
                    clonar_estilo(ref_cell, target_cell)
                target_cell.value = valores_fila[c - 1]

        nombre_salida = f"{sigedo_global}_notificacion_casilla_multiple.xlsx"
        ruta_salida = get_base_dir() / "Informes_Generados" / nombre_salida
        ruta_salida.parent.mkdir(parents=True, exist_ok=True)

        log(f"Guardando Notificación Casilla Múltiple en: {ruta_salida.name}")
        wb.save(ruta_salida)
        return ruta_salida

    def generar_correo_multiple(self, super_contexto: dict, becarios: list, log_callback=None) -> Path:
        def log(msg):
            if log_callback: log_callback(msg)

        ruta_plantilla = get_base_dir() / "plantillas" / "plantilla_notificacion_multiple.xlsx"
        if not ruta_plantilla.exists():
            log(f"No se encontró la plantilla múltiple en {ruta_plantilla}")
            raise FileNotFoundError(f"Falta plantilla: {ruta_plantilla}")

        log("Cargando plantilla de notificación Excel múltiple...")
        from copy import copy
        import re

        wb = openpyxl.load_workbook(ruta_plantilla)
        ws = wb.active

        fila_base = 3
        valores_base = [ws.cell(row=fila_base, column=c).value for c in range(1, 14)]

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
        # Se toma la información base de la plantilla y solo se actualiza el Informe SUS (literal a / ítem I),
        # conservando la información del literal b / ítem II correspondiente al Oficio.
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
                if fila_actual != fila_base:
                    clonar_estilo(ref_cell, target_cell)
                target_cell.value = valores_fila[c - 1]

        nombre_salida = f"{sigedo_global}_notificacion_multiple.xlsx"
        ruta_salida = get_base_dir() / "Informes_Generados" / nombre_salida
        ruta_salida.parent.mkdir(parents=True, exist_ok=True)

        log(f"Guardando Notificación Múltiple en: {ruta_salida.name}")
        wb.save(ruta_salida)
        return ruta_salida
