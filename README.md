<!-- filepath: c:\Users\CCT\projects\docx2pdf\README.md -->

# Convertidor de Word a PDF

Esta es una aplicación de escritorio para Windows que convierte archivos .doc y .docx a formato PDF. Utiliza Microsoft Word para la conversión, por lo que requiere tener Word instalado.

## Características

- Convierte archivos .doc y .docx a PDF.
- Interfaz gráfica de usuario (GUI) fácil de usar.
- Permite seleccionar carpetas de entrada y salida.
- Muestra un log detallado del proceso de conversión.
- Barra de progreso para seguir el estado de la conversión.
- Manejo de errores y timeouts para evitar bloqueos.
- Limpieza automática de procesos de Word colgados.

## Requisitos Previos

- **Sistema Operativo:** Windows
- **Microsoft Word:** Debe estar instalado y activado en el sistema.
- **Python:** Versión 3.6 o superior.
- **Dependencias de Python:** `customtkinter`, `docx2pdf`, `pywin32`

## Instalación

1. **Clona el repositorio (o descarga los archivos):**

   ```bash
   git clone https://github.com/tu_usuario/word_to_pdf.git
   cd word_to_pdf
   ```

2. **Crea un entorno virtual (recomendado):**

   ```bash
   python -m venv venv
   venv\Scripts\activate  # En Windows
   # source venv/bin/activate  # En macOS/Linux
   ```

3. **Instala las dependencias:**

   ```bash
   pip install customtkinter docx2pdf pywin32
   ```

## Creación de Ejecutable Portable

Para crear un archivo ejecutable portable (.exe) que no requiera tener Python instalado:

1. **Instala PyInstaller:**

   ```bash
   pip install pyinstaller
   ```

2. **Genera el ejecutable:**

   ```bash
   pyinstaller --onefile --windowed --name "ConvertidorWordToPDF" app_convertidor_gui.py
   ```

   **Parámetros explicados:**
   - `--onefile`: Crea un único archivo ejecutable
   - `--windowed`: Evita que aparezca la ventana de consola (solo la GUI)
   - `--name`: Especifica el nombre del archivo ejecutable

3. **Ubicación del ejecutable:**
   El archivo ejecutable se generará en la carpeta `dist/ConvertidorWordToPDF.exe`

4. **Distribución:**
   Puedes copiar el archivo `ConvertidorWordToPDF.exe` a cualquier computadora con Windows que tenga Microsoft Word instalado, sin necesidad de instalar Python ni dependencias adicionales.

**Nota:** El archivo ejecutable será de mayor tamaño (aproximadamente 15-20 MB) ya que incluye el intérprete de Python y todas las dependencias necesarias.

## Uso

1. Ejecuta la aplicación:

   ```bash
   python app_convertidor_gui.py
   ```

2. **Selecciona la Carpeta Word (Entrada):** Haz clic en "Seleccionar" y elige la carpeta que contiene los archivos .doc o .docx que deseas convertir.
3. **Selecciona la Carpeta PDF (Salida):** Haz clic en "Seleccionar" y elige la carpeta donde se guardarán los archivos PDF convertidos. Si no se especifica, se creará automáticamente una subcarpeta llamada `PDF_Convertidos` dentro de la carpeta de entrada.
4. **Inicia la Conversión:** Haz clic en el botón "Iniciar Conversión".
5. **Monitorea el Progreso:** La barra de progreso mostrará el avance y el área de log mostrará mensajes detallados sobre cada archivo.
6. **Revisa los Resultados:** Una vez finalizada la conversión, aparecerá un mensaje. Puedes revisar el log para ver el resumen y los detalles de cualquier error.

## Funcionamiento Interno (Windows)

- La aplicación utiliza la biblioteca `customtkinter` para la interfaz gráfica.
- Para los archivos `.docx`, se utiliza la biblioteca `docx2pdf` que a su vez interactúa con Microsoft Word mediante COM para realizar la conversión a PDF.
- Para los archivos `.doc` (formato antiguo de Word):
  1. Primero se convierten a formato `.docx` utilizando la automatización COM de Microsoft Word. Esta conversión se realiza en un archivo temporal.
  2. Luego, el archivo `.docx` temporal se convierte a PDF utilizando `docx2pdf`.
  3. El archivo temporal `.docx` se elimina después de la conversión.
- Se implementan timeouts para la conversión de `.doc` a `.docx` para evitar que la aplicación se bloquee si Word no responde.
- Se incluye una función para intentar terminar procesos de `WINWORD.EXE` que puedan haber quedado colgados.

## Solución de Problemas

- **Error "Faltan bibliotecas 'docx2pdf' y/o 'pywin32'":**
  Asegúrate de haber instalado las dependencias correctamente: `pip install docx2pdf pywin32`.
- **La conversión falla para algunos archivos:**
  - El archivo podría estar corrupto o protegido por contraseña.
  - El contenido del archivo podría ser demasiado complejo o tener características que Word no puede procesar correctamente durante la automatización.
  - Puede haber problemas de permisos de archivo o carpeta.
  - Asegúrate de que Microsoft Word esté activado y funcionando correctamente.
- **La aplicación se congela o Word no responde:**
  La aplicación intenta manejar timeouts, pero en algunos casos, Word puede no responder. La función `kill_word_processes` intenta mitigar esto, pero puede ser necesario cerrar Word manualmente a través del Administrador de Tareas si persiste el problema.

## Contribuir

Las contribuciones son bienvenidas. Por favor, abre un issue o un pull request para discutir los cambios.

## Licencia

Este proyecto es de código abierto. Siéntete libre de usarlo y modificarlo.
