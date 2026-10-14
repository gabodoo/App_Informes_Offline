import re
from pathlib import Path
from datetime import datetime

import pandas as pd
import pdfplumber

def limpiar_dni(val: str | int | float | None) -> str:
    """Limpia y formatea un DNI asegurando 8 dígitos numéricos."""
    if val is None or pd.isna(val):
        return ""
    val_str = str(val).strip().replace(".0", "")
    val_str = re.sub(r"\D", "", val_str)
    if len(val_str) < 8:
        val_str = val_str.zfill(8)
    return val_str[:8]


def formatear_nombre_ies(nombre: str) -> str:
    """Formatea el nombre de la IES a tipo título respetando conectores en minúsculas."""
    if not nombre:
        return ""
    palabras = str(nombre).strip().split()
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


class ExtractorInformeSuccor:
    """Extrae metadatos del Informe de la SUCCOR para el módulo de Actualización."""
    PATRON_NOMBRE_INFORME = re.compile(
        r"(INFORME\s+(?:N[°ºo\.]?|NRO\.?|NÚMERO|N)?\s*[\d]+[-\s\w/\-\.]+(?:SUCCOR[-\w]*)?)",
        re.IGNORECASE,
    )
    PATRON_SIGEDO = re.compile(
        r"(?:SIGEDO|Expediente|CUT|Trámite|Tramite)\s*(?:N[°ºo\.]?\s*)?([A-Z0-9\.\-]*\d[A-Z0-9\.\-]*)",
        re.IGNORECASE,
    )

    @staticmethod
    def _limpiar(valor: str) -> str:
        v = re.sub(r"\s+", " ", str(valor).replace("\n", " ")).strip(" :-\t")
        v = re.sub(r'\b(N[°º])([^\s])', r'\1 \2', v, flags=re.IGNORECASE)
        return v

    @classmethod
    def limpiar_nombre_succor(cls, nom: str) -> str:
        if not nom:
            return ""
        nom = cls._limpiar(nom)
        nom = re.sub(r"^\s*(?:INFORME|Informe)\s+(?:N[°o\.]*|N\.o|No\.?|NRO\.?|NÚMERO|NUMERO)?\s*", "Informe N° ", nom, flags=re.IGNORECASE)
        if not nom.startswith("Informe N° ") and re.match(r"^\s*(?:INFORME|Informe)\b", nom, re.I):
            nom = re.sub(r"^\s*(?:INFORME|Informe)\s*", "Informe N° ", nom, flags=re.IGNORECASE)
        
        m_lima = re.search(r"^(.*?SUCCOR[-\s/]*LIMA)\b", nom, flags=re.IGNORECASE)
        if m_lima:
            nom = m_lima.group(1).strip()
        else:
            m_lima2 = re.search(r"^(.*?\bLIMA)\b", nom, flags=re.IGNORECASE)
            if m_lima2 and "SUCCOR" in nom.upper():
                nom = m_lima2.group(1).strip()
        nom = re.sub(r"(\bLIMA\b)[\s\w\.\-]*$", r"\1", nom, flags=re.IGNORECASE).strip()
        return nom

    @classmethod
    def extraer(cls, ruta_pdf: str | Path) -> dict:
        ruta = Path(ruta_pdf)
        resultado = {
            "nombre_informe_succor": "",
            "numero_informe": "",
            "numero_sigedo": "",
        }

        with pdfplumber.open(ruta) as pdf:
            textos = [p.extract_text() or "" for p in pdf.pages]
            texto_completo = "\n".join(textos)

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

        if resultado.get("nombre_informe_succor"):
            nom = resultado["nombre_informe_succor"]
            nom = re.sub(r"^\s*(?:INFORME|Informe)\s+(?:N[º°o\.]*|N\.o|No)\s*", "Informe N° ", nom, flags=re.IGNORECASE)
            if not nom.startswith("Informe N° ") and re.match(r"^\s*(?:INFORME|Informe)\b", nom, re.I):
                nom = re.sub(r"^\s*(?:INFORME|Informe)\s*", "Informe N° ", nom, flags=re.IGNORECASE)
            nom = cls.limpiar_nombre_succor(nom)
            resultado["nombre_informe_succor"] = nom
            resultado["numero_informe"] = nom

        m_sig = cls.PATRON_SIGEDO.search(texto_completo)
        if m_sig:
            resultado["numero_sigedo"] = m_sig.group(1)

        return resultado


class ProcesadorActualizacionSuspendido:
    """Procesador para la opción Actualización -> Suspendido a Realizando Estudios."""

    def __init__(self, ruta_padron, ruta_informe_succor, ruta_anexo_succor, rutas_resoluciones, nro_informe, log, progreso):
        self.ruta_padron = Path(ruta_padron)
        self.ruta_informe_succor = Path(ruta_informe_succor)
        self.ruta_anexo_succor = Path(ruta_anexo_succor)
        self.rutas_resoluciones = [Path(r) for r in rutas_resoluciones if str(r).strip()]
        self.nro_informe = nro_informe
        self.log = log
        self.progreso = progreso
        self.contexto = {}

    def ejecutar(self):
        self.log("=== INICIANDO PROCESAMIENTO ACTUALIZACIÓN: SUSPENDIDO A REALIZANDO ===")
        self.progreso(0.1, "Extrayendo DNIs del Anexo SUCCOR...")
        dnis_anexo = self._extraer_dnis_anexo()
        
        self.progreso(0.2, "Buscando becarios en el Padrón...")
        becarios_data = self._buscar_en_padron(dnis_anexo)

        self.progreso(0.4, "Extrayendo datos del Informe SUCCOR...")
        datos_succor = self._extraer_datos_succor()

        self.progreso(0.6, "Extrayendo Resoluciones...")
        resoluciones = self._extraer_resoluciones()

        n_becarios = len(becarios_data)
        str_becarios = f"{n_becarios} becarios" if n_becarios > 1 else "1 becario"

        # Becas e instituciones únicas para el asunto
        becas_unicas = list(dict.fromkeys([b.get("beca_asunto") for b in becarios_data if b.get("beca_asunto")]))
        inst_unicas = list(dict.fromkeys([b.get("institucion_asunto") for b in becarios_data if b.get("institucion_asunto")]))

        str_becas = " y ".join(becas_unicas) if becas_unicas else "Beca 18 - 2022"
        str_inst = " y ".join(inst_unicas) if inst_unicas else "la institución"

        hoy = datetime.now()
        meses = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        fecha_str = f"{hoy.day} de {meses[hoy.month]} de {hoy.year}"

        asunto_completo = f"Modificación de Estado, Condición y Situación de {str_becarios} de la {str_becas} de la {str_inst}"

        self.contexto = {
            "nro_informe_word": self.nro_informe,
            "NUMERO_INFORME": self.nro_informe,
            "NUMERO_INFORME_GENERAR": self.nro_informe,
            "CANTIDAD_BECARIOS": str_becarios,
            "BECA_Y_CONVOCATORIA": str_becas,
            "INSTITUCION": str_inst,
            "REFERENCIA": datos_succor["numero_informe"],
            "NOMBRE_INFORME_SUCCOR": datos_succor["numero_informe"],
            "NUMERO_SIGEDO": datos_succor["numero_sigedo"],
            "FECHA_ACTUAL_TEXTO": fecha_str,
            "TEXTO_SUCCOR": datos_succor["texto_numeral_2_1"],
            "RESOLUCIONES": resoluciones,
            "ASUNTO": asunto_completo,
            "becarios": becarios_data,
            "informe_succor": datos_succor,
        }

        rutas_generadas = []
        self.progreso(0.7, "Generando Excel de Actualización...")
        from generador_excel_actualizacion import GeneradorExcelActualizacionSuspendido
        gen_excel = GeneradorExcelActualizacionSuspendido()
        ruta_ex = gen_excel.generar(self.contexto, self.log)
        if ruta_ex: rutas_generadas.append(ruta_ex)

        self.progreso(0.8, "Generando Informe Word de Actualización...")
        from generador_word_actualizacion import GeneradorWordActualizacionSuspendido
        gen_word = GeneradorWordActualizacionSuspendido()
        ruta_wd = gen_word.generar(self.contexto, self.log)
        if ruta_wd: rutas_generadas.append(ruta_wd)

        self.progreso(1.0, "Finalizado")
        return rutas_generadas

    def _extraer_dnis_anexo(self):
        self.log(f"Leyendo anexo: {self.ruta_anexo_succor.name}")
        df = pd.read_excel(self.ruta_anexo_succor)
        dnis = []
        for col in df.columns:
            if 'dni' in str(col).lower() or 'documento' in str(col).lower():
                for val in df[col]:
                    if pd.notnull(val):
                        v_limpio = limpiar_dni(val)
                        if len(v_limpio) == 8:
                            dnis.append(v_limpio)
                break
        
        if not dnis:
            for col in df.columns:
                for val in df[col]:
                    if pd.notnull(val):
                        v_limpio = limpiar_dni(val)
                        if len(v_limpio) == 8:
                            dnis.append(v_limpio)
                            
        dnis = list(dict.fromkeys(dnis))
        self.log(f"Se encontraron {len(dnis)} DNIs en el anexo.")
        if not dnis:
            raise Exception("No se encontraron DNIs válidos en el Anexo SUCCOR.")
        return dnis

    def _buscar_en_padron(self, dnis_buscar):
        self.log(f"Leyendo padrón: {self.ruta_padron.name}")
        try:
            df = pd.read_excel(self.ruta_padron, engine='openpyxl')
        except Exception as e:
            self.log(f"Error al leer padrón: {e}")
            raise

        df.columns = [str(c).strip().upper() for c in df.columns]
        
        col_dni = next((c for c in df.columns if c in ['DNI', 'N° DNI', 'NRO DNI', 'DOCUMENTO']), None)
        if not col_dni:
            raise Exception("No se encontró columna DNI en el padrón.")
        
        df[col_dni] = df[col_dni].apply(limpiar_dni)
        
        becarios = []
        for dni in dnis_buscar:
            fila = df[df[col_dni] == dni]
            if fila.empty:
                self.log(f"ADVERTENCIA: DNI {dni} no encontrado en el padrón.")
                continue
            
            fila = fila.iloc[0]
            
            def get_val(cols):
                for c in cols:
                    if c in df.columns:
                        v = fila[c]
                        return str(v).strip() if pd.notnull(v) else ""
                return ""

            expediente = get_val(['NEXPEDIE', 'NEXPEDIENTE', 'EXPEDIENTE', 'N° EXPEDIENTE', 'NRO EXPEDIENTE', 'NRO_EXPEDIENTE', 'EXP'])
            if not expediente:
                col_exp = next((c for c in df.columns if 'EXPEDIE' in c), None)
                if col_exp:
                    expediente = str(fila[col_exp]).strip() if pd.notnull(fila[col_exp]) else ""

            paterno = get_val(['PATERNO', 'APELLIDO PATERNO', 'APELLIDOS'])
            materno = get_val(['MATERNO', 'APELLIDO MATERNO'])
            nombres = get_val(['NOMBRES', 'NOMBRE'])
            
            if paterno and materno and nombres:
                becario_nombre = f"{paterno} {materno} {nombres}"
            elif paterno and nombres:
                becario_nombre = f"{paterno} {nombres}"
            else:
                becario_nombre = get_val(['BECARIO', 'BECARIO/A', 'NOMBRES Y APELLIDOS'])

            programa = get_val(['PROGRAMA', 'BECA', 'MODALIDAD', 'TIPO DE BECA'])
            anio = get_val(['AÑO_CONV', 'AÑO_CONVOCATORIA', 'AÑO CONVOCATORIA', 'AÑO DE CONVOCATORIA', 'ANIO_CONVOCATORIA', 'ANIO', 'CONVOCATORIA', 'AÑO'])
            institucion = get_val(['INSTITUCI', 'INSTITUCIÓN', 'IES', 'INSTITUCION', 'UNIVERSIDAD', 'INSTITUTO'])
            if not institucion:
                col_inst = next((c for c in df.columns if 'INSTITUC' in c or 'IES' in c), None)
                if col_inst:
                    institucion = str(fila[col_inst]).strip() if pd.notnull(fila[col_inst]) else ""

            # Extraer año de 4 dígitos (ej. 2023) evitando cadenas largas de convocatoria
            m_anio = re.search(r"\b(20\d{2})\b", str(anio))
            if m_anio:
                anio_limpio = m_anio.group(1)
            else:
                m_anio2 = re.search(r"\b(20\d{2})\b", str(programa))
                anio_limpio = m_anio2.group(1) if m_anio2 else str(anio).strip().replace(".0", "")

            prog_upper = str(programa).upper()
            if "18" in prog_upper or "BECA 18" in prog_upper:
                prog_nombre = "Beca 18"
            elif "PERMANENCIA" in prog_upper:
                prog_nombre = "Beca Permanencia"
            elif "CONTINUIDAD" in prog_upper:
                prog_nombre = "Beca Continuidad"
            else:
                prog_nombre = " ".join([p.capitalize() for p in str(programa).split()]) if programa else "Beca 18"

            beca_limpia = f"{prog_nombre} - {anio_limpio}" if anio_limpio else prog_nombre
            beca_asunto = beca_limpia
            beca_cuadro = beca_limpia

            # Formato específico para Excel: PROGRAMA - AÑO_CONVOCATORIA tal como viene en padrón (Imagen 1 y 3)
            prog_ex = str(programa).strip().upper()
            anio_ex = str(anio).strip().replace(".0", "").upper()
            if prog_ex and anio_ex:
                beca_excel = f"{prog_ex} - {anio_ex}"
            elif prog_ex:
                beca_excel = prog_ex
            elif anio_ex:
                beca_excel = anio_ex
            else:
                beca_excel = beca_limpia.upper()

            institucion_asunto = formatear_nombre_ies(institucion)
            institucion_cuadro = institucion.upper()

            # Extracción del campo FIN del padrón (Imagen 3: Columna AN 'FIN')
            col_fin = next((c for c in df.columns if c in ['FIN', 'FECHA FIN SEGÚN SIBEC', 'FECHA_FIN', 'FECHA FIN']), None)
            fin_val = fila.get(col_fin) if col_fin else None
            if pd.isnull(fin_val):
                fin_val = fila.get('FIN')
            if pd.isnull(fin_val):
                fin_val = fila.get('FECHA FIN SEGÚN SIBEC')
            if pd.isnull(fin_val):
                fin_val = fila.get('FECHA_FIN')
                
            fin_str = ""
            if pd.notnull(fin_val):
                if isinstance(fin_val, (pd.Timestamp, datetime)):
                    fin_str = fin_val.strftime("%d/%m/%Y")
                else:
                    val_s = str(fin_val).strip()
                    m_date = re.match(r"^(\d{4})-(\d{2})-(\d{2})", val_s)
                    if m_date:
                        fin_str = f"{m_date.group(3)}/{m_date.group(2)}/{m_date.group(1)}"
                    else:
                        fin_str = val_s

            estado = get_val(['ESTADO_BECA', 'ESTADO'])
            condicion = get_val(['CONDICIÓN', 'CONDICION'])
            situacion = get_val(['SITUACIÓN', 'SITUACION'])

            becarios.append({
                "expediente": expediente,
                "dni": dni,
                "becario": becario_nombre.upper(),
                "beca": beca_cuadro,
                "beca_excel": beca_excel.upper(),
                "beca_asunto": beca_asunto,
                "institucion": institucion_cuadro,
                "institucion_asunto": institucion_asunto,
                "fin": fin_str,
                "estado": estado.upper(),
                "condicion": condicion.upper(),
                "situacion": situacion.upper()
            })
        
        self.log(f"Se encontraron {len(becarios)} becarios en el padrón.")
        return becarios

    def _extraer_datos_succor(self):
        self.log(f"Leyendo Informe SUCCOR: {self.ruta_informe_succor.name}")
        datos = {
            "numero_informe": "",
            "numero_sigedo": "",
            "texto_numeral_2_1": ""
        }
        with pdfplumber.open(self.ruta_informe_succor) as pdf:
            textos = [p.extract_text() or "" for p in pdf.pages]
            texto_completo = "\n".join(textos)

        ext = ExtractorInformeSuccor.extraer(self.ruta_informe_succor)
        
        # 1. Extracción de SIGEDO (Imagen 2: SIGEDO: 75568-2026)
        sigedo_val = ""
        m_sig = re.search(r"SIGEDO\s*[:\s]+(\d+-\d{4})", texto_completo, re.IGNORECASE)
        if m_sig:
            sigedo_val = m_sig.group(1).strip()
        else:
            m_sig2 = re.search(r"SIGEDO\s*[:\s]+(\d{4,8})", texto_completo, re.IGNORECASE)
            if m_sig2:
                sigedo_val = f"{m_sig2.group(1).strip()}-2026"
            elif ext.get("numero_sigedo"):
                s = str(ext.get("numero_sigedo")).strip()
                sigedo_val = s if "-" in s else f"{s}-2026"
            else:
                m_fn = re.search(r"(\d+-\d{4})", self.ruta_informe_succor.name)
                if m_fn:
                    sigedo_val = m_fn.group(1)
                else:
                    m_fn2 = re.search(r"\b(\d{5,8})\b", self.ruta_informe_succor.name)
                    if m_fn2:
                        sigedo_val = f"{m_fn2.group(1)}-2026"
        datos["numero_sigedo"] = sigedo_val

        # 2. Extracción de Informe SUCCOR para REFERENCIA (Imagen 3 -> Imagen 4)
        m_inf = re.search(r"(?:INFORME|Informe)\s*(?:N[°ºo\.]*|N\.o|No\.?)?\s*([\d]+[-\s\w/\-\.]+(?:SUCCOR[-\w]*|LIMA)?)", texto_completo, re.IGNORECASE)
        if m_inf:
            raw_inf = m_inf.group(0).strip()
            num_succor = ExtractorInformeSuccor.limpiar_nombre_succor(raw_inf)
        elif ext.get("numero_informe"):
            num_succor = ExtractorInformeSuccor.limpiar_nombre_succor(ext.get("numero_informe"))
        elif ext.get("nombre_informe_succor"):
            num_succor = ExtractorInformeSuccor.limpiar_nombre_succor(ext.get("nombre_informe_succor"))
        else:
            num_succor = "Informe N° S/N-2026-MINEDU/VMGI-PRONABEC-DICONCI-SUCCOR-LIMA"

        num_succor = re.sub(r"^\s*INFORME\b", "Informe", num_succor, flags=re.IGNORECASE)
        if not num_succor.startswith("Informe N° "):
            num_succor = re.sub(r"^Informe\s*(?:N[°ºo\.]*)?\s*", "Informe N° ", num_succor, flags=re.IGNORECASE)
        datos["numero_informe"] = num_succor

        # 3. Numeral 2.1
        m_21 = re.search(r"(?:2\.1\.?|2\.1)\s*(.*?)(?:2\.2\.?|2\.2|$)", texto_completo, re.DOTALL | re.IGNORECASE)
        if m_21:
            texto_21 = m_21.group(1).strip()
            texto_21 = re.sub(r'\s+', ' ', texto_21)
            datos["texto_numeral_2_1"] = texto_21
        else:
            self.log("ADVERTENCIA: No se encontró el numeral 2.1 en el Informe SUCCOR explícitamente.")
            datos["texto_numeral_2_1"] = "No se pudo extraer el numeral 2.1 del Informe de la SUCCOR."

        return datos

    def _extraer_resoluciones(self):
        resoluciones_nums = []
        anio_resolucion = "2026"

        for ruta_item in self.rutas_resoluciones:
            val_str = str(ruta_item).strip()
            if not val_str:
                continue

            ruta = Path(val_str)
            if ruta.exists() and ruta.suffix.lower() == '.pdf':
                self.log(f"Analizando resolución: {ruta.name}")
                try:
                    with pdfplumber.open(ruta) as pdf:
                        paginas_leer = pdf.pages[:3]
                        texto_completo = "\n".join([p.extract_text() or "" for p in paginas_leer])

                    num_encontrado = None

                    # 1. Patrón completo: N° 2654-2026-MINEDU...
                    m_completo = re.search(r"N[°ºo\.\s]*(\d+)\s*-\s*(20\d{2})\s*-\s*MINEDU", texto_completo, re.IGNORECASE)
                    if m_completo:
                        num_encontrado = m_completo.group(1).strip()
                        anio_resolucion = m_completo.group(2).strip()

                    # 2. Patrón: Resolución Jefatural ... N° 2654
                    if not num_encontrado:
                        m_rj = re.search(r"(?:RESOLUCI[OÓ]N\s+JEFATURAL|RESOLUCI[OÓ]N|RJ)[\s\S]{0,60}?N[°ºo\.\s]*(\d+)", texto_completo, re.IGNORECASE)
                        if m_rj:
                            num_encontrado = m_rj.group(1).strip()
                            m_y = re.search(r"\b(20\d{2})\b", texto_completo)
                            if m_y:
                                anio_resolucion = m_y.group(1)

                    # 3. Patrón: N° 2654-2026
                    if not num_encontrado:
                        m_nro_anio = re.search(r"N[°ºo\.\s]*(\d+)\s*-\s*(20\d{2})", texto_completo, re.IGNORECASE)
                        if m_nro_anio:
                            num_encontrado = m_nro_anio.group(1).strip()
                            anio_resolucion = m_nro_anio.group(2).strip()

                    # 4. Patrón: N° 2654
                    if not num_encontrado:
                        m_nro = re.search(r"\bN[°ºo\.]\s*(\d{2,6})\b", texto_completo, re.IGNORECASE)
                        if m_nro:
                            num_encontrado = m_nro.group(1).strip()

                    # 5. Fallback al nombre del archivo
                    if not num_encontrado:
                        m_stem = re.search(r"(\d{2,6})", ruta.stem)
                        if m_stem:
                            num_encontrado = m_stem.group(1)
                        else:
                            num_encontrado = ruta.stem

                    if num_encontrado:
                        resoluciones_nums.append(str(num_encontrado).strip())

                except Exception as e:
                    self.log(f"Error leyendo resolución {ruta.name}: {e}")
                    m_stem = re.search(r"(\d+)", ruta.stem)
                    resoluciones_nums.append(m_stem.group(1) if m_stem else ruta.stem)
            else:
                # Si no es un archivo PDF existente pero el usuario escribió un texto o número directamente
                m_txt = re.search(r"(\d+)", val_str)
                if m_txt:
                    resoluciones_nums.append(m_txt.group(1))
                    m_y = re.search(r"\b(20\d{2})\b", val_str)
                    if m_y:
                        anio_resolucion = m_y.group(1)
                else:
                    resoluciones_nums.append(val_str)

        if not resoluciones_nums:
            return f"Resolución Jefatural N° S/N-{anio_resolucion}-MINEDU/VMGI-PRONABEC-DIBEC"

        # Mantener el orden y eliminar duplicados
        nums_unicos = list(dict.fromkeys([n for n in resoluciones_nums if n]))
        if len(nums_unicos) == 1:
            return f"Resolución Jefatural N° {nums_unicos[0]}-{anio_resolucion}-MINEDU/VMGI-PRONABEC-DIBEC"
        else:
            nums_str = ", ".join(nums_unicos)
            return f"Resolución Jefatural Nros {nums_str}-{anio_resolucion}-MINEDU/VMGI-PRONABEC-DIBEC"


# Alias para compatibilidad de nomenclatura
ProcesadorSuspendidoRegular = ProcesadorActualizacionSuspendido

