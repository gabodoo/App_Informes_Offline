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
    APP_SIZE = "1000x750"
    MIN_SIZE = (900, 600)

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
        self._tipo_carga_formatos_mult = tk.StringVar(value="Individual")
        self._ruta_formato_consolidado_mult = tk.StringVar()
        self._ruta_informe_succor_mult = tk.StringVar()
        self._ruta_calendario_academico_mult = tk.StringVar()
        self._ruta_documento_ies_mult = tk.StringVar()

        self._modo_actual = "actualizacion"
        self._menu_ampliaciones_abierto = False
        self._procesando = False

        self._construir_ui()

    def _construir_ui(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- Sidebar (Menú lateral simulado) ---
        sidebar = ctk.CTkFrame(self, width=220, fg_color=self.COLOR_FONDO_BLANCO, border_width=1, border_color=self.COLOR_BORDE, corner_radius=0)
        sidebar.grid(row=0, column=0, rowspan=3, sticky="nsew")
        sidebar.grid_propagate(False)

        # Logo PRONABEC simulado
        logo_lbl = ctk.CTkLabel(sidebar, text="PRONABEC", font=ctk.CTkFont(size=22, weight="bold"), text_color=self.COLOR_AZUL_OSCURO)
        logo_lbl.pack(pady=(20, 25), padx=20, anchor="w")

        # 1. Opción principal: Actualización
        self.btn_actualizacion = ctk.CTkButton(
            sidebar,
            text="Actualización",
            fg_color="transparent",
            text_color=self.COLOR_TEXTO_OSCURO,
            anchor="w",
            hover_color="#f0f0f0",
            font=ctk.CTkFont(size=14),
            command=lambda: self._cambiar_modo("actualizacion")
        )
        self.btn_actualizacion.pack(fill="x", pady=4, padx=10)
        self.btn_ecs = self.btn_actualizacion

        # 2. Opción principal: Ampliaciones (desplegable)
        self.btn_ampliaciones = ctk.CTkButton(
            sidebar,
            text="📁 Ampliaciones  ▶",
            fg_color="transparent",
            text_color=self.COLOR_TEXTO_OSCURO,
            anchor="w",
            hover_color="#f0f0f0",
            font=ctk.CTkFont(size=14),
            command=self._toggle_menu_ampliaciones
        )
        self.btn_ampliaciones.pack(fill="x", pady=4, padx=10)

        # Contenedor para subopciones de Ampliaciones (desglosable)
        self.frame_submenu_ampliaciones = ctk.CTkFrame(sidebar, fg_color="transparent")

        self.btn_individual = ctk.CTkButton(
            self.frame_submenu_ampliaciones,
            text="    ✎ Individual",
            fg_color="transparent",
            text_color=self.COLOR_TEXTO_OSCURO,
            anchor="w",
            hover_color="#f0f0f0",
            font=ctk.CTkFont(size=13),
            command=lambda: self._cambiar_modo("individual")
        )
        self.btn_individual.pack(fill="x", pady=2, padx=(10, 10))

        self.btn_multiple = ctk.CTkButton(
            self.frame_submenu_ampliaciones,
            text="    📥 Múltiple",
            fg_color="transparent",
            text_color=self.COLOR_TEXTO_OSCURO,
            anchor="w",
            hover_color="#f0f0f0",
            font=ctk.CTkFont(size=13),
            command=lambda: self._cambiar_modo("multiple")
        )
        self.btn_multiple.pack(fill="x", pady=2, padx=(10, 10))


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

        self._btn_abrir_carpeta = ctk.CTkButton(toolbar, text="📂 Abrir Carpeta", fg_color="#27ae60", hover_color="#1e8449", text_color="white", font=ctk.CTkFont(size=13, weight="bold"), corner_radius=0, command=self._abrir_carpeta_salida)
        self._btn_abrir_carpeta.pack(side="left", padx=2, fill="y", ipadx=10)


        # --- Contenedor de Vistas ---
        self.main_container = ctk.CTkFrame(self, fg_color=self.COLOR_FONDO_BLANCO)
        self.main_container.grid(row=2, column=1, sticky="nsew", padx=15, pady=(10, 15))
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(0, weight=1)

        self.frames = {}
        
        # Vista Individual
        self.frames["individual"] = self._construir_vista_individual(self.main_container)
        self.frames["individual"].grid(row=0, column=0, sticky="nsew")
        
        # Vista Múltiple
        self.frames["multiple"] = self._construir_vista_multiple(self.main_container)
        self.frames["multiple"].grid(row=0, column=0, sticky="nsew")

        # Vista Actualización (pantalla en blanco)
        self.frames["actualizacion"] = self._construir_vista_actualizacion(self.main_container)
        self.frames["actualizacion"].grid(row=0, column=0, sticky="nsew")
        self.frames["actualizacion_ecs"] = self.frames["actualizacion"]

        # --- Progreso y Log (Común) ---
        status_frame = ctk.CTkFrame(self.main_container, fg_color=self.COLOR_FONDO_BLANCO)
        status_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        status_frame.grid_columnconfigure(0, weight=1)
        status_frame.grid_rowconfigure(2, weight=1)

        self._lbl_estado = ctk.CTkLabel(status_frame, text="Módulo: Actualización", font=ctk.CTkFont(size=12), text_color=self.COLOR_TEXTO_OSCURO, anchor="w")
        self._lbl_estado.grid(row=0, column=0, sticky="ew", pady=(0, 4))

        self._barra_progreso = ctk.CTkProgressBar(status_frame, progress_color=self.COLOR_AZUL_CLARO, height=8)
        self._barra_progreso.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self._barra_progreso.set(0)

        self._log_text = ctk.CTkTextbox(status_frame, height=130, font=ctk.CTkFont(family="Consolas", size=12), border_width=1, border_color=self.COLOR_BORDE, fg_color="#fafafa", text_color="#333", corner_radius=0, state="disabled")
        self._log_text.grid(row=2, column=0, sticky="nsew")

        self._cambiar_modo("actualizacion")

    def _toggle_menu_ampliaciones(self, abrir: bool | None = None) -> None:
        if abrir is None:
            nuevo_estado = not self._menu_ampliaciones_abierto
        else:
            nuevo_estado = abrir

        if nuevo_estado:
            if not self._menu_ampliaciones_abierto:
                self.frame_submenu_ampliaciones.pack(fill="x", pady=(0, 5), padx=0)
                self._menu_ampliaciones_abierto = True
            self.btn_ampliaciones.configure(text="📁 Ampliaciones  ▼")
            if self._modo_actual not in ("individual", "multiple"):
                self._cambiar_modo("individual")
        else:
            if self._menu_ampliaciones_abierto:
                self.frame_submenu_ampliaciones.pack_forget()
                self._menu_ampliaciones_abierto = False
            self.btn_ampliaciones.configure(text="📁 Ampliaciones  ▶")

    def _cambiar_modo(self, modo: str) -> None:
        if modo in ("actualizacion", "actualizacion_ecs"):
            self._modo_actual = "actualizacion"
        else:
            self._modo_actual = modo

        # Resetear colores
        self.btn_actualizacion.configure(text_color=self.COLOR_TEXTO_OSCURO)
        self.btn_ampliaciones.configure(text_color=self.COLOR_TEXTO_OSCURO)
        self.btn_individual.configure(text_color=self.COLOR_TEXTO_OSCURO)
        self.btn_multiple.configure(text_color=self.COLOR_TEXTO_OSCURO)

        if self._modo_actual == "actualizacion":
            self.btn_actualizacion.configure(text_color=self.COLOR_ROSA)
            if not self._procesando:
                self._btn_generar.configure(state="normal")
            self._lbl_estado.configure(text="Módulo: Actualización")

        elif self._modo_actual == "individual":
            if not self._menu_ampliaciones_abierto:
                self._toggle_menu_ampliaciones(abrir=True)
            self.btn_ampliaciones.configure(text_color=self.COLOR_AZUL_OSCURO)
            self.btn_individual.configure(text_color=self.COLOR_ROSA)
            if not self._procesando:
                self._btn_generar.configure(state="normal")
            self._lbl_estado.configure(text="Estado: Listo (Modo Individual)")

        elif self._modo_actual == "multiple":
            if not self._menu_ampliaciones_abierto:
                self._toggle_menu_ampliaciones(abrir=True)
            self.btn_ampliaciones.configure(text_color=self.COLOR_AZUL_OSCURO)
            self.btn_multiple.configure(text_color=self.COLOR_ROSA)
            if not self._procesando:
                self._btn_generar.configure(state="normal")
            self._lbl_estado.configure(text="Estado: Listo (Modo Múltiple)")

        target_mode = "actualizacion" if self._modo_actual == "actualizacion" else self._modo_actual
        if target_mode in self.frames:
            self.frames[target_mode].tkraise()

    def _construir_vista_actualizacion(self, parent) -> ctk.CTkFrame:
        # Pantalla en blanco para Actualización
        frame = ctk.CTkFrame(parent, fg_color=self.COLOR_FONDO_BLANCO)
        return frame

    def _construir_vista_individual(self, parent) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(parent, fg_color=self.COLOR_FONDO_BLANCO)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        
        scrollable_frame = ctk.CTkScrollableFrame(frame, fg_color=self.COLOR_FONDO_BLANCO, border_width=1, border_color=self.COLOR_BORDE, corner_radius=0)
        scrollable_frame.grid(row=0, column=0, sticky="nsew")
        scrollable_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(scrollable_frame, text="Nro. Informe", font=ctk.CTkFont(size=13, weight="bold"), text_color=self.COLOR_TEXTO_OSCURO).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        nro_entry = ctk.CTkEntry(scrollable_frame, textvariable=self._nro_informe, placeholder_text="Ej: 6348", height=34, border_color=self.COLOR_BORDE, corner_radius=2, fg_color="#fcfcfc")
        nro_entry.grid(row=0, column=1, sticky="w", padx=(0, 15), pady=(15, 8))

        self._crear_fila_seleccion(scrollable_frame, fila=1, etiqueta="Cargar padrón (.xlsx)", variable=self._ruta_excel, comando=lambda: self._seleccionar_archivo(self._ruta_excel, "Padrón (Excel)", [("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]), placeholder="Padrón/Base de datos de becarios...")
        self._crear_fila_seleccion(scrollable_frame, fila=2, etiqueta="Formato autogenerado (.pdf)", variable=self._ruta_formato_autogenerado, comando=lambda: self._seleccionar_archivo(self._ruta_formato_autogenerado, "Formato autogenerado", [("PDF", "*.pdf"), ("Todos", "*.*")]), placeholder="PDF con fecha/hora de ingreso...")
        self._crear_fila_seleccion(scrollable_frame, fila=3, etiqueta="Informe SUCCOR (.pdf)", variable=self._ruta_informe_succor, comando=lambda: self._seleccionar_archivo(self._ruta_informe_succor, "Informe SUCCOR", [("PDF", "*.pdf"), ("Todos", "*.*")]), placeholder="Informe SUCCOR del becario...")
        self._crear_fila_seleccion(scrollable_frame, fila=4, etiqueta="Calendario académico (.pdf)", variable=self._ruta_calendario_academico, comando=lambda: self._seleccionar_archivo(self._ruta_calendario_academico, "Calendario académico", [("PDF", "*.pdf"), ("Todos", "*.*")]), placeholder="Calendario académico de la IES...")
        self._crear_fila_seleccion(scrollable_frame, fila=5, etiqueta="Documento de la IES (.pdf/.xlsx)", variable=self._ruta_documento_ies, comando=lambda: self._seleccionar_archivo(self._ruta_documento_ies, "Documento IES", [("PDF o Excel", "*.pdf *.xlsx *.xls"), ("PDF", "*.pdf"), ("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]), placeholder="Documento/Carta emitida por la IES (.pdf o .xlsx)...")
        
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
        self._crear_fila_seleccion(scrollable_frame, fila=4, etiqueta="Documento de la IES (.pdf/.xlsx)", variable=self._ruta_documento_ies_mult, comando=lambda: self._seleccionar_archivos(self._ruta_documento_ies_mult, "Documento IES", [("PDF o Excel", "*.pdf *.xlsx *.xls"), ("PDF", "*.pdf"), ("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]), placeholder="Documento/Carta de la IES (uno o varios PDFs, o Excel)...")
        
        # Sección dinámica para Formatos Autogenerados
        separator = ctk.CTkFrame(scrollable_frame, height=2, fg_color=self.COLOR_BORDE)
        separator.grid(row=5, column=0, columnspan=2, sticky="ew", padx=15, pady=20)

        lbl_formatos = ctk.CTkLabel(scrollable_frame, text="Formatos Autogenerados", font=ctk.CTkFont(size=14, weight="bold"), text_color=self.COLOR_AZUL_OSCURO)
        lbl_formatos.grid(row=6, column=0, columnspan=2, padx=15, pady=(0, 8), sticky="w")

        # Menú desplegable para elegir tipo de formato: Individual o Múltiple
        fila_opcion_formato = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        fila_opcion_formato.grid(row=7, column=0, columnspan=2, padx=15, pady=(0, 10), sticky="ew")
        fila_opcion_formato.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            fila_opcion_formato,
            text="Tipo de formato:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=self.COLOR_TEXTO_OSCURO
        ).grid(row=0, column=0, padx=(0, 15), sticky="w")

        self.om_tipo_formatos = ctk.CTkOptionMenu(
            fila_opcion_formato,
            values=["Individual", "Múltiple"],
            variable=self._tipo_carga_formatos_mult,
            command=self._on_tipo_carga_formatos_cambiado,
            height=34,
            width=180,
            fg_color=self.COLOR_AZUL_CLARO,
            button_color=self.COLOR_AZUL_OSCURO,
            button_hover_color="#1d4ed8",
            dropdown_fg_color="#ffffff",
            dropdown_text_color=self.COLOR_TEXTO_OSCURO,
            dropdown_hover_color="#f3f4f6",
            font=ctk.CTkFont(size=13, weight="bold"),
            corner_radius=4,
        )
        self.om_tipo_formatos.grid(row=0, column=1, sticky="w")

        self.lbl_formatos_ayuda = ctk.CTkLabel(
            fila_opcion_formato,
            text="ℹ️ Modo Individual: Cargue los formatos autogenerados (.pdf) de cada becario uno por uno.",
            font=ctk.CTkFont(size=11, slant="italic"),
            text_color="#4b5563"
        )
        self.lbl_formatos_ayuda.grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 0))

        # Contenedor A: Formatos Individuales (uno por uno)
        self.frame_formatos_individual = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        self.frame_formatos_individual.grid(row=8, column=0, columnspan=2, sticky="ew")
        self.frame_formatos_individual.grid_columnconfigure(0, weight=1)

        self.frame_formatos_dinamico = ctk.CTkFrame(self.frame_formatos_individual, fg_color="transparent")
        self.frame_formatos_dinamico.grid(row=0, column=0, sticky="ew")
        self.frame_formatos_dinamico.grid_columnconfigure(1, weight=1)

        self.btn_add_formato = ctk.CTkButton(
            self.frame_formatos_individual,
            text="+ Añadir Formato",
            width=120,
            fg_color=self.COLOR_VERDE,
            hover_color="#219150",
            command=self._add_formato_autogenerado
        )
        self.btn_add_formato.grid(row=1, column=0, padx=15, pady=10, sticky="w")

        # Iniciar con 2 formatos por defecto ya que es "Múltiple"
        self._add_formato_autogenerado()
        self._add_formato_autogenerado()

        # Contenedor B: Formato Consolidado (un solo PDF con 2 a 5 becarios)
        self.frame_formatos_consolidado = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        self.frame_formatos_consolidado.grid_columnconfigure(1, weight=1)

        self._crear_fila_seleccion(
            self.frame_formatos_consolidado,
            fila=0,
            etiqueta="PDF consolidado (2 a 5 becarios):",
            variable=self._ruta_formato_consolidado_mult,
            comando=lambda: self._seleccionar_archivo(
                self._ruta_formato_consolidado_mult,
                "Formato Autogenerado Consolidado",
                [("PDF", "*.pdf"), ("Todos", "*.*")]
            ),
            placeholder="Archivo PDF único que une los formatos autogenerados de todos los becarios..."
        )
        lbl_info_cons = ctk.CTkLabel(
            self.frame_formatos_consolidado,
            text="El sistema reconocerá automáticamente la información de cada becario (1 por 1) sin confundirlos.",
            font=ctk.CTkFont(size=11),
            text_color="#4b5563"
        )
        lbl_info_cons.grid(row=1, column=1, sticky="w", padx=(0, 15), pady=(2, 10))

        # Por defecto, ocultar el contenedor consolidado (modo individual activo)
        self.frame_formatos_consolidado.grid_remove()

        return frame

    def _on_tipo_carga_formatos_cambiado(self, seleccion: str) -> None:
        if seleccion == "Múltiple":
            self.frame_formatos_individual.grid_remove()
            self.frame_formatos_consolidado.grid(row=8, column=0, columnspan=2, sticky="ew")
            self.lbl_formatos_ayuda.configure(
                text="ℹ️ Modo Múltiple: Cargue un archivo PDF consolidado donde se junte la información de 2 a 5 becarios."
            )
            self._log("Modalidad seleccionada: Formatos Autogenerados Múltiples (PDF consolidado)")
        else:
            self.frame_formatos_consolidado.grid_remove()
            self.frame_formatos_individual.grid(row=8, column=0, columnspan=2, sticky="ew")
            self.lbl_formatos_ayuda.configure(
                text="ℹ️ Modo Individual: Cargue los formatos autogenerados (.pdf) de cada becario uno por uno."
            )
            self._log("Modalidad seleccionada: Formatos Autogenerados Individuales (1 por 1)")

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
        ).grid(row=fila, column=0, padx=15, pady=8, sticky="w")

        fila_controles = ctk.CTkFrame(parent, fg_color="transparent")
        fila_controles.grid(
            row=fila, column=1,
            padx=(0, 15),
            pady=8,
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

    def _seleccionar_archivos(self, variable: tk.StringVar, titulo: str, tipos: list) -> None:
        rutas = filedialog.askopenfilenames(
            title=f"Seleccionar {titulo} (puede seleccionar uno o varios archivos)",
            filetypes=tipos,
        )
        if rutas:
            variable.set("; ".join(rutas))
            self._log(f"{titulo} seleccionado(s): {len(rutas)} archivo(s)")

    # --- Lógica principal ---

    def _abrir_carpeta_salida(self) -> None:
        try:
            from generador_word import SALIDA_DIR
            salida = SALIDA_DIR
        except Exception:
            salida = Path("Informes_Generados").resolve()
        salida.mkdir(parents=True, exist_ok=True)
        try:
            import os
            os.startfile(str(salida))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\n{e}")

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
            self._ruta_formato_consolidado_mult.set("")
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
        if self._modo_actual not in ("individual", "multiple"):
            messagebox.showwarning("Seleccione una opción", "Por favor seleccione 'Individual' o 'Múltiple' dentro del menú Ampliaciones.")
            return False

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
            
            tipo_fmt = self._tipo_carga_formatos_mult.get()
            if tipo_fmt == "Múltiple":
                if not self._ruta_formato_consolidado_mult.get().strip():
                    messagebox.showwarning("Datos incompletos", "Cargue el archivo: Formato autogenerado consolidado (.pdf)")
                    return False
            else:
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
        if self._modo_actual not in ("individual", "multiple"):
            messagebox.showinfo("Información", "Seleccione 'Individual' o 'Múltiple' dentro de Ampliaciones para generar informes.")
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
                res = procesador.ejecutar()
                rutas_generadas = [res] if isinstance(res, Path) else (list(res) if res else [])
            else:
                tipo_fmt = self._tipo_carga_formatos_mult.get()
                if tipo_fmt == "Múltiple":
                    formatos = [self._ruta_formato_consolidado_mult.get().strip()]
                    tipo_param = "multiple"
                else:
                    formatos = [v.get() for v in self._formatos_autogenerados_mult if v.get().strip()]
                    tipo_param = "individual"

                procesador = ProcesadorInformes(
                    ruta_excel=self._ruta_excel_mult.get(),
                    ruta_formato_autogenerado=formatos[0],
                    ruta_informe_succor=self._ruta_informe_succor_mult.get(),
                    ruta_calendario_academico=self._ruta_calendario_academico_mult.get(),
                    ruta_documento_ies=self._ruta_documento_ies_mult.get(),
                    log=self._log,
                    progreso=self._actualizar_progreso,
                    nro_informe=self._nro_informe_mult.get().strip(),
                    rutas_formatos=formatos,
                    tipo_formato_autogenerado=tipo_param
                )
                res = procesador.ejecutar_multiple()
                rutas_generadas = list(res) if res else []
                
            alerta_elec = getattr(procesador, 'alerta_electivos', '')
            advertencias_list = getattr(procesador, 'advertencias', [])

            salida_dir = None
            if rutas_generadas:
                salida_dir = rutas_generadas[0].parent
            else:
                try:
                    from generador_word import SALIDA_DIR
                    salida_dir = SALIDA_DIR
                except Exception:
                    salida_dir = Path('Informes_Generados').resolve()

            rutas_str = "\n".join(str(r.name) for r in rutas_generadas) if rutas_generadas else "(Sin archivos)"

            def _notificar_finalizacion():
                # 1. Aviso emergente de cursos electivos pendientes (Individual y Múltiple)
                if alerta_elec:
                    messagebox.showwarning("Aviso", alerta_elec)

                # 2. Ventana emergente de advertencia al concluir si hay alertas adicionales
                if advertencias_list:
                    adv_text = "\n\n".join(advertencias_list)
                    messagebox.showwarning("Atención", adv_text)

                # 3. Notificación de informe(s) completado(s)
                messagebox.showinfo(
                    'Completado',
                    f'Operación finalizada.\n\nCarpeta de salida:\n{salida_dir}\n\nArchivos:\n{rutas_str}',
                )
                self._btn_nuevo.configure(state='normal')

                # 4. Abrir carpeta de salida en el explorador
                if salida_dir and salida_dir.exists():
                    try:
                        import os
                        os.startfile(str(salida_dir))
                    except Exception:
                        pass

            self.after(0, _notificar_finalizacion)
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