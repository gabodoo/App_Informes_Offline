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

from procesador import ProcesadorInformes


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

        self._nro_informe = tk.StringVar()
        self._ruta_excel = tk.StringVar()
        self._ruta_formato_autogenerado = tk.StringVar()
        self._ruta_informe_succor = tk.StringVar()
        self._ruta_calendario_academico = tk.StringVar()
        self._ruta_documento_ies = tk.StringVar()
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

        menus = [("✎ Individual", self.COLOR_ROSA), ("📥 Multiple", self.COLOR_TEXTO_OSCURO)]
        for texto, color in menus:
            btn = ctk.CTkButton(sidebar, text=texto, fg_color="transparent", text_color=color, anchor="w", hover_color="#f0f0f0", font=ctk.CTkFont(size=14))
            btn.pack(fill="x", pady=5, padx=10)


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
        
        # Botones de barra superior para parecerse a Sigedo
        self._btn_generar = ctk.CTkButton(toolbar, text="Generar", fg_color=self.COLOR_NARANJA, hover_color="#d68910", text_color="white", font=ctk.CTkFont(size=13, weight="bold"), corner_radius=0, command=self._iniciar_generacion)
        self._btn_generar.pack(side="left", padx=2, fill="y", ipadx=10)

        self._btn_nuevo = ctk.CTkButton(toolbar, text="Nuevo", fg_color=self.COLOR_AZUL_CLARO, hover_color="#1f618d", text_color="white", font=ctk.CTkFont(size=13, weight="bold"), corner_radius=0, command=self._limpiar_para_nuevo_informe, state="disabled")
        self._btn_nuevo.pack(side="left", padx=2, fill="y", ipadx=10)


        # --- Contenido Principal (Formulario Web) ---
        main_content = ctk.CTkFrame(self, fg_color=self.COLOR_FONDO_BLANCO)
        main_content.grid(row=2, column=1, sticky="nsew", padx=20, pady=20)
        main_content.grid_columnconfigure(0, weight=1)
        main_content.grid_rowconfigure(6, weight=1)

        # Tabla contenedora
        form_frame = ctk.CTkFrame(main_content, fg_color=self.COLOR_FONDO_BLANCO, border_width=1, border_color=self.COLOR_BORDE, corner_radius=0)
        form_frame.grid(row=0, column=0, sticky="ew")
        form_frame.grid_columnconfigure(1, weight=1)

        # Nro. Informe (Nuevo Campo)
        ctk.CTkLabel(form_frame, text="Nro. Informe", font=ctk.CTkFont(size=13, weight="bold"), text_color=self.COLOR_TEXTO_OSCURO).grid(row=0, column=0, padx=15, pady=(20, 10), sticky="w")
        nro_entry = ctk.CTkEntry(form_frame, textvariable=self._nro_informe, placeholder_text="Ej: 6348", height=34, border_color=self.COLOR_BORDE, corner_radius=2, fg_color="#fcfcfc")
        nro_entry.grid(row=0, column=1, sticky="w", padx=(0, 15), pady=(20, 10))

        # Archivos
        self._crear_fila_seleccion(form_frame, fila=1, etiqueta="Cargar padrón (.xlsx)", variable=self._ruta_excel, comando=self._seleccionar_excel, placeholder="Padrón/Base de datos de becarios...")
        self._crear_fila_seleccion(form_frame, fila=2, etiqueta="Formato autogenerado (.pdf)", variable=self._ruta_formato_autogenerado, comando=self._seleccionar_formato_autogenerado, placeholder="PDF con fecha/hora de ingreso...")
        self._crear_fila_seleccion(form_frame, fila=3, etiqueta="Informe SUCCOR (.pdf)", variable=self._ruta_informe_succor, comando=self._seleccionar_informe_succor, placeholder="Informe SUCCOR del becario...")
        self._crear_fila_seleccion(form_frame, fila=4, etiqueta="Calendario académico (.pdf)", variable=self._ruta_calendario_academico, comando=self._seleccionar_calendario_academico, placeholder="Calendario académico de la IES...")
        self._crear_fila_seleccion(form_frame, fila=5, etiqueta="Documento de la IES (.pdf/.xlsx)", variable=self._ruta_documento_ies, comando=self._seleccionar_documento_ies, placeholder="Documento/Carta emitida por la IES...")

        # --- Progreso y Log ---
        status_frame = ctk.CTkFrame(main_content, fg_color=self.COLOR_FONDO_BLANCO)
        status_frame.grid(row=6, column=0, sticky="nsew", pady=(20, 0))
        status_frame.grid_columnconfigure(0, weight=1)
        status_frame.grid_rowconfigure(2, weight=1)

        self._lbl_estado = ctk.CTkLabel(status_frame, text="Estado: Listo", font=ctk.CTkFont(size=12), text_color=self.COLOR_TEXTO_OSCURO, anchor="w")
        self._lbl_estado.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        self._barra_progreso = ctk.CTkProgressBar(status_frame, progress_color=self.COLOR_AZUL_CLARO, height=8)
        self._barra_progreso.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        self._barra_progreso.set(0)

        self._log_text = ctk.CTkTextbox(status_frame, font=ctk.CTkFont(family="Consolas", size=12), border_width=1, border_color=self.COLOR_BORDE, fg_color="#fafafa", text_color="#333", corner_radius=0, state="disabled")
        self._log_text.grid(row=2, column=0, sticky="nsew")

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

    # --- Selectores de archivos ---

    def _seleccionar_excel(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccionar Padrón (Excel)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self._ruta_excel.set(ruta)
            self._log(f"Padrón seleccionado: {ruta}")

    def _seleccionar_formato_autogenerado(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccionar Formato autogenerado",
            filetypes=[("PDF", "*.pdf"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self._ruta_formato_autogenerado.set(ruta)
            self._log(f"Formato autogenerado seleccionado: {ruta}")

    def _seleccionar_informe_succor(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccionar Informe SUCCOR",
            filetypes=[("PDF", "*.pdf"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self._ruta_informe_succor.set(ruta)
            self._log(f"Informe SUCCOR seleccionado: {ruta}")

    def _seleccionar_calendario_academico(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccionar Calendario académico",
            filetypes=[("PDF", "*.pdf"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self._ruta_calendario_academico.set(ruta)
            self._log(f"Calendario académico seleccionado: {ruta}")

    def _seleccionar_documento_ies(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccionar Documento de la IES",
            filetypes=[("PDF o Excel", "*.pdf *.xlsx *.xls"), ("Todos los archivos", "*.*")],
        )
        if ruta:
            self._ruta_documento_ies.set(ruta)
            self._log(f"Documento IES seleccionado: {ruta}")

    # --- Lógica principal ---

    def _limpiar_para_nuevo_informe(self) -> None:
        self._nro_informe.set("")
        self._ruta_formato_autogenerado.set("")
        self._ruta_informe_succor.set("")
        self._ruta_calendario_academico.set("")
        self._ruta_documento_ies.set("")
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
        self._log(f"Inicio de generación de informe N° {self._nro_informe.get().strip()}")

        hilo = threading.Thread(target=self._ejecutar_procesamiento, daemon=True)
        hilo.start()

    def _ejecutar_procesamiento(self) -> None:
        try:
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
            rutas_str = "\n".join(str(r.name) for r in rutas_generadas)
            
            fecha_fin = getattr(procesador, 'fecha_fin', None)
            from datetime import date
            if fecha_fin and fecha_fin < date(2026, 1, 1):
                msg_alerta = f"ATENCIÓN: La fecha de fin en el Informe es {fecha_fin.strftime('%d/%m/%Y')} (antes del 1 de enero de 2026)."
                self.after(
                    0,
                    lambda m=msg_alerta: messagebox.showwarning("Atención - Fecha de fin", m)
                )
                
            self.after(
                0,
                lambda: messagebox.showinfo(
                    "Completado",
                    f"Archivos generados correctamente en la carpeta Informes_Generados:\n\n{rutas_str}",
                ),
            )
            self.after(0, lambda: self._btn_nuevo.configure(state="normal"))
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

