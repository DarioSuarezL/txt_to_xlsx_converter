from charset_normalizer import detect

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
        encoding_detectado = detect(datos)['encoding']
        return encoding_detectado

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
        print("Codificación detectada:", enc)
        with open(ruta_archivo, 'r', encoding=enc) as archivo:
            return [linea.strip() for linea in archivo.readlines()]
    except Exception as e:
        print(f'Error al leer el archivo: {e}')
        return []