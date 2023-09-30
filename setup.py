# pip install cx_Freeze
import sys
import os
import urllib.request
import subprocess
from cx_Freeze import setup, Executable
import sys

if sys.platform == 'win32' and sys.getwindowsversion()[0] >= 6:
    from ctypes import windll
    windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, 0)
    sys.exit(0)

# Función para descargar y ejecutar el archivo de instalación del controlador
def instalar_controlador():
    url = 'https://download.microsoft.com/download/2/4/3/24375141-E08D-4803-AB0E-10F2E3A07AAA/AccessDatabaseEngine.exe'
    nombre_archivo = 'AccessDatabaseEngine.exe'
    urllib.request.urlretrieve(url, nombre_archivo)
    subprocess.call([nombre_archivo, '/passive'])

# Directorio donde se encuentra el archivo de configuración
base_dir = os.path.dirname(os.path.abspath(__file__))

# Opciones para generar el paquete de instalación
build_exe_options = {
    "packages": ["os", "tkinter", "pypyodbc", "pyodbc"],  # Asegúrate de incluir todos los paquetes que tu aplicación utiliza #"packages": ["os"],
    "include_files": [os.path.join(base_dir, "config.ini")],
}

# Configuración del paquete de instalación
setup(
    name="CopyDB",
    version="1.1",
    description="Copy database from one location to another",
    options={"build_exe": build_exe_options},
    executables=[Executable("copyDB.py", base="Win32GUI")],
)

# Instalar el controlador después de construir el paquete de instalación
instalar_controlador()
