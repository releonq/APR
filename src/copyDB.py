# pip install pypyodbc
# pip install pyodbc
# pip install pyinstaller (Para crear un ejecutable)
# Para Crear el Ejecutable
# 1.- pyinstaller --name copydb copydb.py (pyinstaller --name nombre_ejecutable archivo.py)
# 2.- pyinstaller copydb.spec (pyinstaller nombre_ejecutable.spec)
# 3.- python setup.py bdist_msi


import os
import pyodbc
#from pyodbc import compact
import pypyodbc
import configparser
import logging
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

# Crea una ventana
window = tk.Tk()
window.title("Progreso de la copia")

# Crea una barra de progreso
progress = ttk.Progressbar(window, length=500, mode='determinate')
progress.pack()

# Crea una etiqueta para el porcentaje
percentage_label = tk.Label(window, text="0%")
percentage_label.pack()

logging.basicConfig(filename='registro.log', level=logging.INFO)

# Función para actualizar la barra de progreso y la etiqueta del porcentaje
def actualizar_progreso(valor):
    progress['value'] = valor
    percentage_label['text'] = "{:.2f}%".format(valor)  # formatea el valor a dos decimales
    window.update_idletasks()

def mostrar_mensaje(mensaje):
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana de Tkinter
    messagebox.showinfo("Información", mensaje)
    root.destroy()

def solicitar_ruta(mensaje):
    ruta = input(mensaje)
    while not os.path.isfile(ruta):
        mostrar_mensaje(f"La ruta ingresada no es válida: {ruta}")
        logging.info(f"La ruta no es valida {ruta}")
        ruta = input(mensaje)
    return ruta

def obtener_columnas(tabla):
    cursor = conexion.cursor()
    cursor.execute(f"SELECT TOP 1 * FROM {tabla}")
    #num_col = [column[0] for column in cursor.description]
    logging.info(f"Se obtienen numero de columnas")
    return [column[0] for column in cursor.description]

def verificar_columnas(tabla_origen, tabla_destino):
    columnas_origen = obtener_columnas(tabla_origen)
    columnas_destino = obtener_columnas(tabla_destino)
    if columnas_origen == columnas_destino:
        logging.info(f"Se obtienen las columnas desde el origen y destino")
        return True
    else:
        logging.info("Las columnas de las tablas no coinciden")
        mostrar_mensaje("Las columnas de las tablas no coinciden")
        mostrar_mensaje(f"Tabla origen: {columnas_origen}")
        mostrar_mensaje(f"Tabla destino: {columnas_destino}")
        return False

def copiar_registros(origen, destino):
    logging.info("Se inicia el proceso de copia de registros")
    try:
        conexion_origen = pyodbc.connect(f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={origen};")
        conexion_destino = pyodbc.connect(f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={destino};")
        tabla_origen = "boletas"
        tabla_destino = "boletas"
        if not verificar_columnas(tabla_origen, tabla_destino):
            return

        cursor_origen = conexion_origen.cursor()
        cursor_origen.execute(f"SELECT * FROM {tabla_origen}")
        registros = cursor_origen.fetchall()
        total_registros = len(registros)
        logging.info(f"Total de registros a copiar: {total_registros}")
        mostrar_mensaje(f"Total de registros a copiar: {total_registros}")

        cursor_destino = conexion_destino.cursor()
        for index, registro in enumerate(registros):
            actualizar_progreso((index + 1) * 100 / total_registros)
            try:
                cursor_destino.execute(f"INSERT INTO {tabla_destino} VALUES ({', '.join(['?' for _ in range(len(registro))])})", registro)
                conexion_destino.commit()
                logging.info(f"OK  {index+1} de {total_registros}: Registro copiado correctamente")
                print(f"Registro {index+1} de {total_registros} copiado correctamente")
            except pyodbc.Error as e:
                if isinstance(e, pyodbc.IntegrityError) and e.args[0] == '23000':
                    logging.error(f"ERR {index+1} de {total_registros}: Registro Duplicado")
                    print(f"ERROR: Registro {index+1} de {total_registros} registro duplicado")
                else:
                    logging.error(f"Error al copiar registro {index+1}: {e}")
                    print(f"Error al copiar registro {index+1}: {e}")

        mostrar_mensaje("Copia de registros finalizada con exito")
    except pyodbc.Error as e:
        mostrar_mensaje(f"Error al conectar a la base de datos: {e}")

if __name__ == '__main__':
    # Configurar logger
    logging.basicConfig(filename='log.txt', level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')

    logging.info('Inicio del script')
    conexion = None
    # Cargar archivo de configuración
    config = configparser.ConfigParser()
    ruta_config = os.path.join(os.path.dirname(__file__), 'config.ini')
    config.read(ruta_config)

    # for key in config['PATH']:  
    #   print(key)
    # Obtener rutas de archivo origen y destino desde el archivo de configuración
    try:
        ruta_origen = config['PATH']['origen']
        ruta_destino = config['PATH']['destino']
        logging.info(f"Rutas de archivo de origen y destino leídas desde archivo de configuración: "
                        f"{ruta_origen}, {ruta_destino}")
    except KeyError:
        logging.error("No se pudo obtener las rutas de archivo de origen y destino desde archivo de configuración")
        ruta_origen = None
        ruta_destino = None
    # Leer rutas de origen y destino
    #config_origen = config['PATH']['origen']
    #config_destino = config['PATH']['destino']

    while True:

        if not ruta_origen:
            ruta_origen = solicitar_ruta("Por favor, ingrese la ruta de archivo de origen: ")
        else:
            if os.path.isfile(ruta_origen):
                logging.info(f"Ruta de archivo de origen válida: {ruta_origen}")
            else:
                logging.warning(f"La ruta de archivo de origen no es válida: {ruta_origen}")
                ruta_origen = solicitar_ruta("Por favor, ingrese la ruta de archivo de origen: ")

        if not ruta_destino:
            ruta_destino = solicitar_ruta("Por favor, ingrese la ruta de archivo de destino: ")
        else:
            if os.path.isfile(ruta_destino):
                logging.info(f"Ruta de archivo de destino válida: {ruta_destino}")
            else:
                logging.warning(f"La ruta de archivo de destino no es válida: {ruta_destino}")
                ruta_destino = solicitar_ruta("Por favor, ingrese la ruta de archivo de destino: ")

        # print(f"Ruta Origen: {ruta_origen} \n Ruta Destino {ruta_destino}")
        #logging.info(f"Ruta Origen: {ruta_origen}")
        #logging.info(f"Ruta Destino: {ruta_destino}")
        
        if ruta_origen != ruta_destino:
            try:
                conexion = pyodbc.connect(f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={ruta_origen};")
                copiar_registros(ruta_origen, ruta_destino)
                conexion.close()
                logging.info('OK  Fin del script')
                break
            except pyodbc.Error as e:
                logging.info(f"Error al conectar a la base de datos origen: {e}")
                mostrar_mensaje(f"Error al conectar a la base de datos origen: {e}")
                logging.info('ERR Fin del script')
                break
        else:
            logging.info(f"La ruta de origen y destino deben ser diferentes \n Origen: {ruta_origen} \n Destino: {ruta_destino}")
            mostrar_mensaje("La ruta de origen y destino deben ser diferentes")
