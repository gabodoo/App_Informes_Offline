"""
Interfaz grafica con CustomTkinter para la aplicacion de informes offline (Camuflada como Sigedo Lima).
"""

from __future__ import annotations

import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from procesador import ProcesadorInformes, BecarioNoEncontradoIESException, FechaFinInsuficienteException, BecarioNoCulminariaAmpliacionException


class AppInformes(ctk.CTk):
    """Ventana principal de la aplicacion."""

    APP_TITLE = "Sigedo Lima"
    APP_SIZE = "1000x800"
    MIN_SIZE = (900, 700)

    # Colores Sigedo
    COLOR_AZUL_OSCURO = "#004b87"
    COLOR_NARANJA = "#f39c12"
    COLOR_VERDE = "#27ae60"
    COLOR_AZUL_CLARO = "#2980b9"
    COLOR_ROSA = "#e84393"
    COLOR_FONDO_BLANCO = "#ffffff"
    COLOR_TEXTO_OSCURO = "#333333"
    COLOR_BORDE = "#e0e0e0"

    def __init__(self) -> None:
        super().__init__()

        # Forzar modo claro para parecer web
        ctk.set_appearance_mode("Light")
        self.configure(fg_color=self.COLOR_FONDO_BLANCO)

        self.title(self.APP_TITLE)
        self.geometry(self.APP_SIZE)
        self.minsize(*self.MIN_SIZE)

        # Variables Individual
        self._nro_informe = tk.StringVar()
        self._ruta_excel = tk.StringVar()
        self._ruta_formato_autogenerado = tk.StringVar()
        self._ruta_informe_succor = tk.StringVar()
        self._ruta_calendario_academico = tk.StringVar()
        self._ruta_documento_ies = tk.StringVar()

        # Variables Múltiple
        self._nro_informe_mult = tk.StringVar()
        self._ruta_excel_mult = tk.StringVar()
        self._formatos_autogenerados_mult = [] # Lista de StringVars
        self._ruta_informe_succor_mult = tk.StringVar()
        self._ruta_calendario_academico_mult = tk.StringVar()
        self._ruta_documento_ies_mult = tk.StringVar()

        self._modo_actual = "individual"
        self._procesando = False

        self._construir_ui()

    def _construir_ui(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- Sidebar (Menú lateral simulado) ---
        sidebar = ctk.CTkFrame(self, width=200, fg_color=self.COLOR_FONDO_BLANCO, border_width=1, border_color=self.COLOR_BORDE, corner_radius=0)
        sidebar.grid(row=0, column=0, rowspan=3, sticky="nsew")
        sidebar.grid_propagate(False)

        # Logo PRONABEC simulado
        logo_lbl = ctk.CTkLabel(sidebar, text="PRONABEC", font=ctk.CTkFont(size=22, weight="bold"), text_color=self.COLOR_AZUL_OSCURO)
        logo_lbl.pack(pady=(20, 30), padx=20, anchor="w")

        self.btn_individual = ctk.CTkButton(sidebar, text="✎ Individual", fg_color="transparent", text_color=self.COLOR_ROSA, anchor="w", hover_color="#f0f0f0", font=ctk.CTkFont(size=14), command=lambda: self._cambiar_modo("individual"))
        self.btn_individual.pack(fill="x", pady=5, padx=10)

        self.btn_multiple = ctk.CTkButton(sidebar, text="📥 Múltiple", fg_color="transparent", text_color=self.COLOR_TEXTO_OSCURO, anchor="w", hover_color="#f0f0f0", font=ctk.CTkFont(size=14), command=lambda: self._cambiar_modo("multiple"))
        self.btn_multiple.pack(fill="x", pady=5, padx=10)


        # --- Header Azul (Top Bar) ---
        header = ctk.CTkFrame(self, fg_color=self.COLOR_AZUL_OSCURO, corner_radius=0, height=60)
        header.grid(row=0, column=1, sticky="ew")
        header.grid_propagate(False)

        header_lbl = ctk.CTkLabel(
            header,
            text="<   SIGEDO                    MACRO REGION LIMA",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="white",
        )
        header_lbl.pack(side="left", padx=20, pady=15)

        user_lbl = ctk.CTkLabel(header, text="👤 Te damos la Bienvenida", font=ctk.CTkFont(size=11), text_color="white", justify="right")
        user_lbl.pack(side="right", padx=20, pady=10)

        # --- Toolbar (Botones de colores simulados) ---
        toolbar = ctk.CTkFrame(self, fg_color=self.COLOR_FONDO_BLANCO, corner_radius=0, height=50)
        toolbar.grid(row=1, column=1, sticky="ew", padx=10, pady=(10,0))
        
        self._btn_generar = ctk.CTkButton(toolbar, text="Generar", fg_color=self.COLOR_NARANJA, hover_color="#d68910", text_color="white", font=ctk.CTkFont(size=13, weight="bold"), corner_radius=0, command=self._iniciar_generacion)
        self._btn_generar.pack(side="left", padx=2, fill="y", ipadx=10)

        self._btn_nuevo = ctk.CTkButton(toolbar, text="Nuevo", fg_color=self.COLOR_AZUL_CLARO, hover_color="#1f618d", text_color="white", font=ctk.CTkFont(size=13, weight="bold"), corner_radius=0, command=self._limpiar_para_nuevo_informe, state="disabled")
        self._btn_nuevo.pack(side="left", padx=2, fill="y", ipadx=10)


        # --- Contenedor de Vistas ---
        self.main_container = ctk.CTkFrame(self, fg_color=self.COLOR_FONDO_BLANCO)
        self.main_container.grid(row=2, column=1, sticky="nsew", padx=20, pady=20)
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(0, weight=1)

        self.frames = {}
        
        # Vista Individual
        self.frames["individual"] = self._construir_vista_individual(self.main_container)
        self.frames["individual"].grid(row=0, column=0, sticky="nsew")
        
        # Vista Múltiple
        self.frames["multiple"] = self._construir_vista_multiple(self.main_container)
        self.frames["multiple"].grid(row=0, column=0, sticky="nsew")

        # --- Progreso y Log (Común) ---
        status_frame = ctk.CTkFrame(self.main_container, fg_color=self.COLOR_FONDO_BLANCO)
        status_frame.grid(row=1, column=0, sticky="nsew", pady=(20, 0))
        status_frame.grid_columnconfigure(0, weight=1)
        status_frame.grid_rowconfigure(2, weight=1)

        self._lbl_estado = ctk.CTkLabel(status_frame, text="Estado: Listo", font=ctk.CTkFont(size=12), text_color=self.COLOR_TEXTO_OSCURO, anchor="w")
        self._lbl_estado.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        self._barra_progreso = ctk.CTkProgressBar(status_frame, progress_color=self.COLOR_AZUL_CLARO, height=8)
        self._barra_progreso.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        self._barra_progreso.set(0)

        self._log_text = ctk.CTkTextbox(status_frame, font=ctk.CTkFont(family="Consolas", size=12), border_width=1, border_color=self.COLOR_BORDE, fg_color="#fafafa", text_color="#333", corner_radius=0, state="disabled")
        self._log_text.grid(row=2, column=0, sticky="nsew")

        self._cambiar_modo("individual")

    def _cambiar_modo(self, modo: str) -> None:
        self._modo_actual = modo
        if modo == "individual":
            self.btn_individual.configure(text_color=self.COLOR_ROSA)
            self.btn_multiple.configure(text_color=self.COLOR_TEXTO_OSCURO)
        else:
            self.btn_multiple.configure(text_color=self.COLOR_ROSA)
            self.btn_individual.configure(text_color=self.COLOR_TEXTO_OSCURO)
        
        self.frames[modo].tkraise()

    def _construir_vista_individual(self, parent) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(parent, fg_color=self.COLOR_FONDO_BLANCO)
        frame.grid_columnconfigure(0, weight=1)
        
        form_frame = ctk.CTkFrame(frame, fg_color=self.COLOR_FONDO_BLANCO, border_width=1, border_color=self.COLOR_BORDE, corner_radius=0)
        form_frame.grid(row=0, column=0, sticky="ew")
        form_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(form_frame, text="Nro. Informe", font=ctk.CTkFont(size=13, weight="bold"), text_color=self.COLOR_TEXTO_OSCURO).grid(row=0, column=0, padx=15, pady=(20, 10), sticky="w")
        nro_entry = ctk.CTkEntry(form_frame, textvariable=self._nro_informe, placeholder_text="Ej: 6348", height=34, border_color=self.COLOR_BORDE, corner_radius=2, fg_color="#fcfcfc")
        nro_entry.grid(row=0, column=1, sticky="w", padx=(0, 15), pady=(20, 10))

        self._crear_fila_seleccion(form_frame, fila=1, etiqueta="Cargar padrón (.xlsx)", variable=self._ruta_excel, comando=lambda: self._seleccionar_archivo(self._ruta_excel, "Padrón (Excel)", [("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]), placeholder="Padrón/Base de datos de becarios...")
        self._crear_fila_seleccion(form_frame, fila=2, etiqueta="Formato autogenerado (.pdf)", variable=self._ruta_formato_autogenerado, comando=lambda: self._seleccionar_archivo(self._ruta_formato_autogenerado, "Formato autogenerado", [("PDF", "*.pdf"), ("Todos", "*.*")]), placeholder="PDF con fecha/hora de ingreso...")
        self._crear_fila_seleccion(form_frame, fila=3, etiqueta="Informe SUCCOR (.pdf)", variable=self._ruta_informe_succor, comando=lambda: self._seleccionar_archivo(self._ruta_informe_succor, "Informe SUCCOR", [("PDF", "*.pdf"), ("Todos", "*.*")]), placeholder="Informe SUCCOR del becario...")
        self._crear_fila_seleccion(form_frame, fila=4, etiqueta="Calendario académico (.pdf)", variable=self._ruta_calendario_academico, comando=lambda: self._seleccionar_archivo(self._ruta_calendario_academico, "Calendario académico", [("PDF", "*.pdf"), ("Todos", "*.*")]), placeholder="Calendario académico de la IES...")
        self._crear_fila_seleccion(form_frame, fila=5, etiqueta="Documento de la IES (.pdf/.xlsx)", variable=self._ruta_documento_ies, comando=lambda: self._seleccionar_archivo(self._ruta_documento_ies, "Documento IES", [("PDF o Excel", "*.pdf *.xlsx *.xls"), ("Todos", "*.*")]), placeholder="Documento/Carta emitida por la IES...")
        
        return frame

    def _construir_vista_multiple(self, parent) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(parent, fg_color=self.COLOR_FONDO_BLANCO)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        
        # Make the form scrollable to handle multiple files
        scrollable_frame = ctk.CTkScrollableFrame(frame, fg_color=self.COLOR_FONDO_BLANCO, border_width=1, border_color=self.COLOR_BORDE, corner_radius=0)
        scrollable_frame.grid(row=0, column=0, sticky="nsew")
        scrollable_frame.grid_columnconfigure(1, weight=1)

        # Usamos las mismas variables para los documentos compartidos para facilitar
        ctk.CTkLabel(scrollable_frame, text="Nro. Informe (Múltiple)", font=ctk.CTkFont(size=13, weight="bold"), text_color=self.COLOR_TEXTO_OSCURO).grid(row=0, column=0, padx=15, pady=(20, 10), sticky="w")
        nro_entry = ctk.CTkEntry(scrollable_frame, textvariable=self._nro_informe_mult, placeholder_text="Ej: 6348", height=34, border_color=self.COLOR_BORDE, corner_radius=2, fg_color="#fcfcfc")
        nro_entry.grid(row=0, column=1, sticky="w", padx=(0, 15), pady=(20, 10))

        self._crear_fila_seleccion(scrollable_frame, fila=1, etiqueta="Cargar padrón (.xlsx)", variable=self._ruta_excel_mult, comando=lambda: self._seleccionar_archivo(self._ruta_excel_mult, "Padrón (Excel)", [("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]), placeholder="Padrón/Base de datos de becarios...")
        self._crear_fila_seleccion(scrollable_frame, fila=2, etiqueta="Informe SUCCOR Compartido (.pdf)", variable=self._ruta_informe_succor_mult, comando=lambda: self._seleccionar_archivo(self._ruta_informe_succor_mult, "Informe SUCCOR", [("PDF", "*.pdf"), ("Todos", "*.*")]), placeholder="Informe SUCCOR (Múltiples becarios)...")
        self._crear_fila_seleccion(scrollable_frame, fila=3, etiqueta="Calendario académico (.pdf)", variable=self._ruta_calendario_academico_mult, comando=lambda: self._seleccionar_archivo(self._ruta_calendario_academico_mult, "Calendario académico", [("PDF", "*.pdf"), ("Todos", "*.*")]), placeholder="Calendario académico de la IES...")
        self._crear_fila_seleccion(scrollable_frame, fila=4, etiqueta="Documento de la IES (.pdf/.xlsx)", variable=self._ruta_documento_ies_mult, comando=lambda: self._seleccionar_archivo(self._ruta_documento_ies_mult, "Documento IES", [("PDF o Excel", "*.pdf *.xlsx *.xls"), ("Todos", "*.*")]), placeholder="Documento/Carta emitida por la IES...")
        
        # Sección dinámica para Formatos Autogenerados
        separator = ctk.CTkFrame(scrollable_frame, height=2, fg_color=self.COLOR_BORDE)
        separator.grid(row=5, column=0, columnspan=2, sticky="ew", padx=15, pady=20)

        lbl_formatos = ctk.CTkLabel(scrollable_frame, text="Formatos Autogenerados (1 por becario)", font=ctk.CTkFont(size=14, weight="bold"), text_color=self.COLOR_AZUL_OSCURO)
        lbl_formatos.grid(row=6, column=0, columnspan=2, padx=15, pady=(0, 10), sticky="w")

        self.frame_formatos_dinamico = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        self.frame_formatos_dinamico.grid(row=7, column=0, columnspan=2, sticky="ew")
        self.frame_formatos_dinamico.grid_columnconfigure(1, weight=1)

        btn_add_formato = ctk.CTkButton(scrollable_frame, text="+ Añadir Formato", width=120, fg_color=self.COLOR_VERDE, hover_color="#219150", command=self._add_formato_autogenerado)
        btn_add_formato.grid(row=8, column=0, padx=15, pady=10, sticky="w")

        # Iniciar con 2 formatos por defecto ya que es "Múltiple"
        self._add_formato_autogenerado()
        self._add_formato_autogenerado()

        return frame

    def _add_formato_autogenerado(self):
        if len(self._formatos_autogenerados_mult) >= 5:
            messagebox.showinfo("Límite", "El sistema admite un máximo de 5 becarios por Informe Múltiple.")
            return

        idx = len(self._formatos_autogenerados_mult)
        var_ruta = tk.StringVar()
        self._formatos_autogenerados_mult.append(var_ruta)

        fila = idx
        
        etiqueta = ctk.CTkLabel(self.frame_formatos_dinamico, text=f"Formato Becario {idx+1}:", font=ctk.CTkFont(size=13, weight="bold"), text_color=self.COLOR_TEXTO_OSCURO)
        etiqueta.grid(row=fila, column=0, padx=15, pady=5, sticky="w")

        fila_controles = ctk.CTkFrame(self.frame_formatos_dinamico, fg_color="transparent")
        fila_controles.grid(row=fila, column=1, padx=(0, 15), pady=5, sticky="ew")
        fila_controles.grid_columnconfigure(0, weight=1)

        ctk.CTkEntry(
            fila_controles,
            textvariable=var_ruta,
            placeholder_text="PDF formato autogenerado...",
            height=34, border_color=self.COLOR_BORDE, corner_radius=2, fg_color="#fcfcfc"
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            fila_controles,
            text="Examinar", width=100, height=34,
            command=lambda v=var_ruta: self._seleccionar_archivo(v, "Formato autogenerado", [("PDF", "*.pdf"), ("Todos", "*.*")]),
            fg_color="#f0f0f0", text_color=self.COLOR_TEXTO_OSCURO, hover_color="#e0e0e0", border_width=1, border_color=self.COLOR_BORDE, corner_radius=2
        ).grid(row=0, column=1)


    def _crear_fila_seleccion(
        self,
        parent: ctk.CTkFrame,
        fila: int,
        etiqueta: str,
        variable: tk.StringVar,
        comando: callable,
        placeholder: str,
    ) -> None:
        ctk.CTkLabel(
            parent,
            text=etiqueta,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=self.COLOR_TEXTO_OSCURO
        ).grid(row=fila, column=0, padx=15, pady=10, sticky="w")

        fila_controles = ctk.CTkFrame(parent, fg_color="transparent")
        fila_controles.grid(
            row=fila, column=1,
            padx=(0, 15),
            pady=10,
            sticky="ew",
        )
        fila_controles.grid_columnconfigure(0, weight=1)

        ctk.CTkEntry(
            fila_controles,
            textvariable=variable,
            placeholder_text=placeholder,
            height=34,
            border_color=self.COLOR_BORDE,
            corner_radius=2,
            fg_color="#fcfcfc"
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            fila_controles,
            text="Examinar",
            width=100,
            height=34,
            command=comando,
            fg_color="#f0f0f0",
            text_color=self.COLOR_TEXTO_OSCURO,
            hover_color="#e0e0e0",
            border_width=1,
            border_color=self.COLOR_BORDE,
            corner_radius=2,
        ).grid(row=0, column=1)

    def _seleccionar_archivo(self, variable: tk.StringVar, titulo: str, tipos: list) -> None:
        ruta = filedialog.askopenfilename(
            title=f"Seleccionar {titulo}",
            filetypes=tipos,
        )
        if ruta:
            variable.set(ruta)
            self._log(f"{titulo} seleccionado: {ruta}")

    # --- Lógica principal ---

    def _limpiar_para_nuevo_informe(self) -> None:
        if self._modo_actual == "individual":
            self._nro_informe.set("")
            self._ruta_formato_autogenerado.set("")
            self._ruta_informe_succor.set("")
            self._ruta_calendario_academico.set("")
            self._ruta_documento_ies.set("")
        else:
            self._nro_informe_mult.set("")
            for var in self._formatos_autogenerados_mult:
                var.set("")
            self._ruta_informe_succor_mult.set("")
            self._ruta_calendario_academico_mult.set("")
            self._ruta_documento_ies_mult.set("")
            
        self._barra_progreso.set(0)
        self._lbl_estado.configure(text="Estado: Listo")
        self._btn_nuevo.configure(state="disabled")
        self._log("-" * 60)
        self._log("Listo para generar un nuevo informe. Padrón conservado.")

    def _log(self, mensaje: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        linea = f"[{timestamp}] {mensaje}\n"

        def _escribir() -> None:
            self._log_text.configure(state="normal")
            self._log_text.insert("end", linea)
            self._log_text.see("end")
            self._log_text.configure(state="disabled")

        self.after(0, _escribir)

    def _actualizar_progreso(self, valor: float, mensaje: str) -> None:
        def _actualizar() -> None:
            self._barra_progreso.set(max(0.0, min(1.0, valor)))
            self._lbl_estado.configure(text=f"Estado: {mensaje}")

        self.after(0, _actualizar)

    def _validar_entradas(self) -> bool:
        if self._modo_actual == "individual":
            if not self._nro_informe.get().strip():
                messagebox.showwarning("Datos incompletos", "Por favor ingrese el Nro. de Informe a generar.")
                return False

            campos = [
                (self._ruta_excel, "Padrón de becarios (.xlsx)"),
                (self._ruta_formato_autogenerado, "Formato autogenerado (.pdf)"),
                (self._ruta_informe_succor, "Informe SUCCOR (.pdf)"),
                (self._ruta_calendario_academico, "Calendario académico (.pdf)"),
                (self._ruta_documento_ies, "Documento de la IES (.pdf/.xlsx)"),
            ]
            for var, nombre in campos:
                if not var.get().strip():
                    messagebox.showwarning("Datos incompletos", f"Cargue el archivo: {nombre}")
                    return False
        else:
            if not self._nro_informe_mult.get().strip():
                messagebox.showwarning("Datos incompletos", "Por favor ingrese el Nro. de Informe a generar.")
                return False

            campos = [
                (self._ruta_excel_mult, "Padrón de becarios (.xlsx)"),
                (self._ruta_informe_succor_mult, "Informe SUCCOR Compartido (.pdf)"),
                (self._ruta_calendario_academico_mult, "Calendario académico (.pdf)"),
                (self._ruta_documento_ies_mult, "Documento de la IES (.pdf/.xlsx)"),
            ]
            for var, nombre in campos:
                if not var.get().strip():
                    messagebox.showwarning("Datos incompletos", f"Cargue el archivo: {nombre}")
                    return False
            
            formatos_llenos = [v for v in self._formatos_autogenerados_mult if v.get().strip()]
            if len(formatos_llenos) < 2:
                messagebox.showwarning("Datos incompletos", "Debe cargar al menos 2 Formatos Autogenerados para generar un Informe Múltiple.")
                return False

        from generador_word import PLANTILLA_PATH
        if not PLANTILLA_PATH.exists():
            messagebox.showerror(
                "Plantilla no encontrada",
                f"No se encontró el archivo de plantilla Word en:\n{PLANTILLA_PATH.resolve()}\n\n"
                "Por favor coloque el archivo 'plantilla_informe.docx' dentro de la carpeta 'plantillas'.",
            )
            return False

        return True

    def _iniciar_generacion(self) -> None:
        if self._procesando:
            return
        if not self._validar_entradas():
            return

        self._procesando = True
        self._btn_generar.configure(state="disabled")
        self._btn_nuevo.configure(state="disabled")
        self._barra_progreso.set(0)
        self._lbl_estado.configure(text="Estado: Iniciando...")
        self._log("=" * 60)
        
        nro = self._nro_informe.get().strip() if self._modo_actual == "individual" else self._nro_informe_mult.get().strip()
        self._log(f"Inicio de generación de informe N° {nro} (Modo: {self._modo_actual})")

        hilo = threading.Thread(target=self._ejecutar_procesamiento, daemon=True)
        hilo.start()

    def _ejecutar_procesamiento(self) -> None:
        try:
            if self._modo_actual == "individual":
                procesador = ProcesadorInformes(
                    ruta_excel=self._ruta_excel.get(),
                    ruta_formato_autogenerado=self._ruta_formato_autogenerado.get(),
                    ruta_informe_succor=self._ruta_informe_succor.get(),
                    ruta_calendario_academico=self._ruta_calendario_academico.get(),
                    ruta_documento_ies=self._ruta_documento_ies.get(),
                    log=self._log,
                    progreso=self._actualizar_progreso,
                    nro_informe=self._nro_informe.get().strip()
                )
                rutas_generadas = procesador.ejecutar()
            else:
                formatos = [v.get() for v in self._formatos_autogenerados_mult if v.get().strip()]
                procesador = ProcesadorInformes(
                    ruta_excel=self._ruta_excel_mult.get(),
                    ruta_formato_autogenerado=formatos[0],
                    ruta_informe_succor=self._ruta_informe_succor_mult.get(),
                    ruta_calendario_academico=self._ruta_calendario_academico_mult.get(),
                    ruta_documento_ies=self._ruta_documento_ies_mult.get(),
                    log=self._log,
                    progreso=self._actualizar_progreso,
                    nro_informe=self._nro_informe_mult.get().strip(),
                    rutas_formatos=formatos
                )
                rutas_generadas = procesador.ejecutar_multiple()
                
            rutas_str = "\n".join(str(r.name) for r in rutas_generadas) if rutas_generadas else "(Sin archivos)"
            
            # Solo mostrar warning si hubo procesador instanciado y tiene advertencias
            if self._modo_actual == "individual" and hasattr(procesador, 'advertencias') and procesador.advertencias:
                adv_text = "\n\n".join(procesador.advertencias)
                self.after(
                    0,
                    lambda m=adv_text: messagebox.showwarning("Atención", m)
                )

            self.after(
                0,
                lambda: messagebox.showinfo(
                    "Completado",
                    f"Operación finalizada. Archivos:\n\n{rutas_str}",
                ),
            )
            self.after(0, lambda: self._btn_nuevo.configure(state="normal"))
        except (BecarioNoEncontradoIESException, FechaFinInsuficienteException, BecarioNoCulminariaAmpliacionException) as e:
            self._log(f"ADVERTENCIA: {e}")
            msg = str(e)
            self.after(
                0,
                lambda m=msg: messagebox.showwarning("Atención", m)
            )
        except Exception as exc:
            self._log(f"ERROR: {exc}")
            import traceback
            self._log(traceback.format_exc())
            self.after(
                0,
                lambda: messagebox.showerror(
                    "Error",
                    f"No se pudo completar el procesamiento:\n{exc}",
                ),
            )
        finally:
            self.after(0, self._finalizar_procesamiento)

    def _finalizar_procesamiento(self) -> None:
        self._procesando = False
        self._btn_generar.configure(state="normal")
        self._btn_nuevo.configure(state="normal")
