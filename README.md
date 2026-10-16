# Wolgang - Generador de Informes Offline

Aplicación de escritorio para la generación de informes

---

## 📁 Estructura del Proyecto

```
App_Informes_Offline/
├── plantillas/                 # Plantillas Word y Excel oficiales requeridas
│   ├── plantilla_informe.docx
│   ├── plantilla_informe_multiple.docx
│   ├── plantilla_oficio.docx
│   ├── plantilla_notificacion.xlsx
│   ├── plantilla_notificacion_casilla.xlsx
│   ├── plantilla_notificacion_multiple.xlsx
│   ├── plantilla_notificacion_casilla_multiple.xlsx
│   ├── plantilla_notificacion_acto_administrativo.xlsx
│   ├── plantilla_informe_no_viable.docx
│   ├── Plantilla_Informe_suspendido_realizando_estudios.docx
│   └── Plantilla _Actualizacion_suspendido_realizando_estudios.xlsx
├── Informes_Generados/         # Carpeta donde se guardan los archivos generados
├── main.py                     # Punto de entrada principal
├── gui.py                      # Interfaz gráfica (CustomTkinter)
├── procesador.py               # Lógica de Ampliaciones (Individual, Múltiple y No Viable)
├── procesador_actualizacion.py # Lógica de Actualización (De suspendido a regular)
├── generador_word.py           # Generador de informes Word para Ampliaciones
├── generador_word_actualizacion.py # Generador de informes Word para Actualización
├── generador_excel.py          # Generador de notificaciones Excel para Ampliaciones
├── generador_excel_actualizacion.py # Generador de resoluciones Excel para Actualización
├── generador_oficio.py         # Generador de oficios Word (python-docx)
├── requirements.txt            # Dependencias del proyecto
├── actualizar_plantilla_no_viable.bat # Script para adecuar marcadores en plantilla No Viable
├── generar_ejecutable.bat      # Script para compilar el ejecutable portable
└── subir_a_github.bat          # Script para sincronizar cambios con GitHub
```

---

## 🚀 Requisitos e Instalación

### 1. Requisitos
- Windows 10 o Windows 11
- Python 3.10 o superior instalado (marcar la casilla *"Add Python to PATH"* durante la instalación).

### 2. Instalación de dependencias
Abre una terminal (PowerShell o CMD) en esta carpeta y ejecuta:
```bash
pip install -r requirements.txt
```

---

## 💻 Ejecución del Programa

### Modo Desarrollo / Directo:
```bash
python main.py
```

### Compilar Ejecutable (.exe) Portable:
Haz doble clic sobre el archivo:
```
generar_ejecutable.bat
```
El ejecutable resultante se encontrará en `dist\main\main.exe`.
