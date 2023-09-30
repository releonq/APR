import os

def es_archivo_access(ruta):
    extensiones_access = ['.accdb', '.mdb', '.accde', '.accdr', '.mde']
    _, ext = os.path.splitext(ruta)
    return ext.lower() in extensiones_access