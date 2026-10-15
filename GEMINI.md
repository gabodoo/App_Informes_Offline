# POLÍTICA ESTRICTA DE PRIVACIDAD, PROTECCIÓN DE DATOS Y ENTORNO SIN PRIVILEGIOS DE ADMINISTRADOR

Esta regla es de obligatorio cumplimiento en todas las sesiones y proyectos:

1. RESTRICCIÓN ESTRICTA DE ARCHIVOS DE DATOS:
   - Está terminantemente PROHIBIDO que el agente abra, lea o inspeccione archivos con extensiones `.xlsx`, `.xls`, `.pdf`, `.csv`, `.db`, `.sqlite` o carpetas de documentos institucionales (padrones, solicitudes, sustentos).
   
2. RESTRICCIÓN DE EJECUCIÓN EN TERMINAL:
   - Está terminantemente PROHIBIDO que el agente ejecute comandos en la terminal que procesen datos reales o que impriman datos personales (nombres, apellidos, DNIs, números de expedientes, notas, cursos) en la consola.
   - El agente no debe utilizar scripts de prueba que lean bases de datos o expedientes reales.
   - Toda validación de compilación debe hacerse a nivel de código o con datos ficticios/mocks creados expresamente para la prueba.

3. ROL DEL AGENTE:
   - El agente debe limitarse a analizar, corregir y escribir código fuente (`.py`, etc.).
   - Toda prueba de ejecución con datos reales la realizará exclusivamente el USUARIO de forma local en su propia ventana externa de Windows (PowerShell / interfaz gráfica).

4. ENTORNO SIN PRIVILEGIOS DE ADMINISTRADOR (USUARIO ESTÁNDAR / DOMINIO INSTITUCIONAL):
   - Los equipos de trabajo pertenecen a un entorno institucional/corporativo (Dominio PRONABEC) donde el usuario **NO cuenta con usuario ni contraseña de administrador**.
   - Queda terminantemente PROHIBIDO proponer, ejecutar o requerir acciones que disparen el Control de Cuentas de Usuario (UAC) o soliciten credenciales de administrador (tales como binarios con elevación, ejecutables como `language_server_windows_x64.exe`, instaladores a nivel de máquina, cambios en `C:\Program Files`, `C:\Windows`, registro de Windows o variables de sistema).
   - Toda solución, script batch (`.bat`), entorno de Python (`venv`), compilación portable con PyInstaller y lectura/escritura de archivos debe operar **100% en espacio de usuario** (rutas locales del proyecto, carpetas de usuario `%USERPROFILE%`, Desktop, etc.).
   - Ante cualquier necesidad de acceso o modificación que el programa requiera, el agente debe emplear siempre alternativas portables sin privilegios elevados.

