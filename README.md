# Sigedo Lima - Generador de Informes Offline

Aplicación de escritorio para la generación de informes técnicos, oficios y notificaciones oficiales (Individual y Múltiple) para PRONABEC - Macro Región Lima.

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
│   └── plantilla_notificacion_casilla_multiple.xlsx
├── Informes_Generados/         # Carpeta donde se guardan los archivos generados
├── main.py                     # Punto de entrada principal
├── gui.py                      # Interfaz gráfica (CustomTkinter)
├── procesador.py               # Lógica de extracción, validación y procesamiento
├── generador_word.py           # Generador de informes Word (docxtpl)
├── generador_excel.py          # Generador de notificaciones Excel (openpyxl)
├── generador_oficio.py         # Generador de oficios Word (python-docx)
├── requirements.txt            # Dependencias del proyecto
└── generar_ejecutable.bat      # Script para compilar el ejecutable portable
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
