import chardet


def detectar_encoding(ruta_archivo):
    """
    Detecta la codificación de un archivo de texto.
    
    Args:
        ruta_archivo (str): Ruta completa del archivo txt a leer
    
    Returns:
        str: Codificación detectada
    """
    with open(ruta_archivo, 'rb') as archivo:
        datos = archivo.read()
        return chardet.detect(datos)['encoding']

def leer_archivo_txt(ruta_archivo):
    """
    Lee el contenido de un archivo de texto.
    
    Args:
        ruta_archivo (str): Ruta completa del archivo txt a leer
    
    Returns:
        list: Lista de líneas del archivo
    """
    try:
        enc = detectar_encoding(ruta_archivo)
        with open(ruta_archivo, 'r', encoding=enc) as archivo:
            return [linea.strip() for linea in archivo.readlines()]
    except Exception as e:
        print(f'Error al leer el archivo: {e}')
        return []