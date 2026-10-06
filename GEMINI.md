# POLÍTICA ESTRICTA DE PRIVACIDAD Y PROTECCIÓN DE DATOS PERSONALES

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
