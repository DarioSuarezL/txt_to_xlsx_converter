import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter 
from openpyxl.styles import PatternFill, Alignment, Font
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
    
def colorear_fila(ws, counter, fill):
    ws['A'+str(counter)].fill = fill
    ws['B'+str(counter)].fill = fill
    ws['C'+str(counter)].fill = fill
    ws['D'+str(counter)].fill = fill
    ws['E'+str(counter)].fill = fill
    ws['F'+str(counter)].fill = fill
    ws['G'+str(counter)].fill = fill
    ws['H'+str(counter)].fill = fill
    ws['I'+str(counter)].fill = fill
    ws['J'+str(counter)].fill = fill
    ws['K'+str(counter)].fill = fill
    ws['L'+str(counter)].fill = fill
    ws['M'+str(counter)].fill = fill
    ws['N'+str(counter)].fill = fill
    ws['O'+str(counter)].fill = fill
    ws['P'+str(counter)].fill = fill
    ws['Q'+str(counter)].fill = fill
    ws['R'+str(counter)].fill = fill
    ws['S'+str(counter)].fill = fill
    ws['T'+str(counter)].fill = fill
    ws['U'+str(counter)].fill = fill
    ws['V'+str(counter)].fill = fill
    ws['W'+str(counter)].fill = fill
    ws['X'+str(counter)].fill = fill
    ws['Y'+str(counter)].fill = fill
    ws['Z'+str(counter)].fill = fill
    ws['AA'+str(counter)].fill = fill
    ws['AB'+str(counter)].fill = fill
    ws['AC'+str(counter)].fill = fill
    

def procesar_datos(lineas, ws, wb):
    """
    Procesa las líneas de un archivo de texto.
    
    Args:
        lineas (list): Lista de líneas del archivo

    Returns:
        pd.DataFrame: DataFrame con los datos procesados
    """

    ws['A1'] = "REGISTRO"
    ws['B1'] = "ALUMNO"
    ws['C1'] = "SIGLA"
    ws['D1'] = "CARRERA"
    ws['E1'] = "Ingreso"
    ws['F1'] = "PPAC"
    ws['G1'] = "PPACE"
    ws['H1'] = "S/Ano"
    ws['I1'] = "PPS"
    ws['J1'] = "MATERIA 1"
    ws['K1'] = "NOTA 1"
    ws['L1'] = "MATERIA 2"
    ws['M1'] = "NOTA 2"
    ws['N1'] = "MATERIA 3"
    ws['O1'] = "NOTA 3"
    ws['P1'] = "MATERIA 4"
    ws['Q1'] = "NOTA 4"
    ws['R1'] = "MATERIA 5"
    ws['S1'] = "NOTA 5"
    ws['T1'] = "MATERIA 6"
    ws['U1'] = "NOTA 6"
    ws['V1'] = "MATERIA 7"
    ws['W1'] = "NOTA 7"
    ws['X1'] = "MATERIA 8"
    ws['Y1'] = "NOTA 8"
    ws['Z1'] = "MATERIA 9"
    ws['AA1'] = "NOTA 9"
    ws['AB1'] = "MATERIA 10"
    ws['AC1'] = "NOTA 10"

    counter = 2 #fila que inicia

    nombre_contador = 0

    registro = None
    alumno = None
    sigla = None
    carrera = None
    ingreso = None
    ppac = None
    ppace = None

    lineas_a_omitir = [
        "----------------------------------------------------------------------------------------------------",
        "S/Ano  CARR  PPS |SIGLA-GR NOTA |SIGLA-GR NOTA |SIGLA-GR NOTA |SIGLA-GR NOTA |",
        "U.A.G.R.M.                  *  BORRADOR DEL HISTORICO ACADEMICO  *               Pag. : 448",
        "S.I.F.                            NO VALIDO PARA TRAMITES                      Fecha:18/Nov/2024",
        "SANTA CRUZ                                                                       Hora :09:39",
        "----------------------------------------------------------------------------------CPD-fapB50",
        "FORMA DE INGRESO :BACHILLERES DESTACADOS 2012",
        ""
    ]

    fill = PatternFill(start_color="FFDDEBF7", end_color="FFDDEBF7", fill_type = "solid")

    

    for linea in lineas:

        if linea in lineas_a_omitir:
            continue

        if counter % 2 == 0:
            colorear_fila(ws, counter, fill)

        #Extrae datos del estudiante
        match_estudiante = re.search(r'#\s+\d+:\s+Estudiante:\s*(\d+)\s+-\s+(.+)', linea)
        if match_estudiante:
            registro, alumno = match_estudiante.groups()
            nombre_contador += 1
            # print(f"# {nombre_contador}: Estudiante: {alumno}")
            continue
        else:
            #Caso especial de estudiante que tiene registro corto
            match_estudiante = re.search(r'#\s+\d+:\s+Estudiante:\s+(\d+)\s+-\s+(.+)', linea)
            if match_estudiante:
                registro, alumno = match_estudiante.groups()
                continue

        if registro and alumno:
            ws['A'+str(counter)] = registro
            ws['B'+str(counter)] = alumno
        
        #Extrae datos de la carrera
        match_carrera = re.search(r'Carrera Actual\s+:\s+(\S+)\s+(.+?)\s+INGRESO:\s+(\S+)\s+PPAC:\s+(\d+)\s+PPACE:\s+(\d+)', linea)
        if match_carrera:
            sigla, carrera, ingreso, ppac, ppace = match_carrera.groups()
            continue

        if sigla and carrera and ingreso and ppac and ppace:
            # ws['C'+str(counter)] = sigla    #SE OCUPARÁ LA CARRERA DE MATERIA Y NO LA DE ESTUDIANTE
            ws['D'+str(counter)] = carrera
            ws['E'+str(counter)] = ingreso
            ws['F'+str(counter)] = ppac
            ws['G'+str(counter)] = ppace


        #Extrae datos del histórico
        #Si tenia 10 inscritas
        match_historico = re.search(r"^([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            materia8, nota8, materia9, nota9, materia10, nota10 = match_historico.groups()
            ws['X'+str(counter-1)] = materia8
            ws['Y'+str(counter-1)] = nota8
            ws['Z'+str(counter-1)] = materia9
            ws['AA'+str(counter-1)] = nota9
            ws['AB'+str(counter-1)] = materia10
            ws['AC'+str(counter-1)] = nota10
            continue

        #Si tenia 9 inscritas
        match_historico = re.search(r"^([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            materia8, nota8, materia9, nota9 = match_historico.groups()
            ws['X'+str(counter-1)] = materia8
            ws['Y'+str(counter-1)] = nota8
            ws['Z'+str(counter-1)] = materia9
            ws['AA'+str(counter-1)] = nota9
            continue

        #Si tenia 8 inscritas
        match_historico = re.search(r"^([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            materia8, nota8 = match_historico.groups()
            ws['X'+str(counter-1)] = materia8
            ws['Y'+str(counter-1)] = nota8
            continue

        #Si tenia 7 inscritas
        match_historico = re.search(r"(\d{1}/\d{4})\s*(\d+-[A-Za-z0-9-@#%])\s*\(\s*(\d+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3, materia4, nota4, materia5, nota5, materia6, nota6, materia7, nota7 = match_historico.groups()
            ws['C'+str(counter)] = sigla_carrera

            ws['H'+str(counter)] = s_ano
            ws['I'+str(counter)] = pps
            ws['J'+str(counter)] = materia1
            ws['K'+str(counter)] = nota1
            ws['L'+str(counter)] = materia2
            ws['M'+str(counter)] = nota2
            ws['N'+str(counter)] = materia3
            ws['O'+str(counter)] = nota3
            ws['P'+str(counter)] = materia4
            ws['Q'+str(counter)] = nota4
            ws['R'+str(counter)] = materia5
            ws['S'+str(counter)] = nota5
            ws['T'+str(counter)] = materia6
            ws['U'+str(counter)] = nota6
            ws['V'+str(counter)] = materia7
            ws['W'+str(counter)] = nota7

            counter += 1
            continue
        
        #Si tenia 6 inscritas
        match_historico = re.search(r"(\d{1}/\d{4})\s*(\d+-[A-Za-z0-9-@#%])\s*\(\s*(\d+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3, materia4, nota4, materia5, nota5, materia6, nota6 = match_historico.groups()
            ws['C'+str(counter)] = sigla_carrera
            
            ws['H'+str(counter)] = s_ano
            ws['I'+str(counter)] = pps
            ws['J'+str(counter)] = materia1
            ws['K'+str(counter)] = nota1
            ws['L'+str(counter)] = materia2
            ws['M'+str(counter)] = nota2
            ws['N'+str(counter)] = materia3
            ws['O'+str(counter)] = nota3
            ws['P'+str(counter)] = materia4
            ws['Q'+str(counter)] = nota4
            ws['R'+str(counter)] = materia5
            ws['S'+str(counter)] = nota5
            ws['T'+str(counter)] = materia6
            ws['U'+str(counter)] = nota6

            counter += 1
            continue
        
        #Si tenia 5 inscritas
        match_historico = re.search(r"(\d{1}/\d{4})\s*(\d+-[A-Za-z0-9-@#%])\s*\(\s*(\d+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3, materia4, nota4, materia5, nota5 = match_historico.groups()
            ws['C'+str(counter)] = sigla_carrera
            
            ws['H'+str(counter)] = s_ano
            ws['I'+str(counter)] = pps
            ws['J'+str(counter)] = materia1
            ws['K'+str(counter)] = nota1
            ws['L'+str(counter)] = materia2
            ws['M'+str(counter)] = nota2
            ws['N'+str(counter)] = materia3
            ws['O'+str(counter)] = nota3
            ws['P'+str(counter)] = materia4
            ws['Q'+str(counter)] = nota4
            ws['R'+str(counter)] = materia5
            ws['S'+str(counter)] = nota5

            counter += 1
            continue

        #Si tenia 4 inscritas
        match_historico = re.search(r"(\d{1}/\d{4})\s*(\d+-[A-Za-z0-9-@#%])\s*\(\s*(\d+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3, materia4, nota4 = match_historico.groups()
            ws['C'+str(counter)] = sigla_carrera
            
            ws['H'+str(counter)] = s_ano
            ws['I'+str(counter)] = pps
            ws['J'+str(counter)] = materia1
            ws['K'+str(counter)] = nota1
            ws['L'+str(counter)] = materia2
            ws['M'+str(counter)] = nota2
            ws['N'+str(counter)] = materia3
            ws['O'+str(counter)] = nota3
            ws['P'+str(counter)] = materia4
            ws['Q'+str(counter)] = nota4

            counter += 1
            continue

        #Si tenia 3 inscritas
        match_historico = re.search(r"(\d{1}/\d{4})\s*(\d+-[A-Za-z0-9-@#%])\s*\(\s*(\d+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3 = match_historico.groups()
            ws['C'+str(counter)] = sigla_carrera
            
            ws['H'+str(counter)] = s_ano
            ws['I'+str(counter)] = pps
            ws['J'+str(counter)] = materia1
            ws['K'+str(counter)] = nota1
            ws['L'+str(counter)] = materia2
            ws['M'+str(counter)] = nota2
            ws['N'+str(counter)] = materia3
            ws['O'+str(counter)] = nota3

            counter += 1
            continue

        #Si tenia 2 inscritas
        match_historico = re.search(r"(\d{1}/\d{4})\s*(\d+-[A-Za-z0-9-@#%])\s*\(\s*(\d+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2 = match_historico.groups()
            ws['C'+str(counter)] = sigla_carrera
            
            ws['H'+str(counter)] = s_ano
            ws['I'+str(counter)] = pps
            ws['J'+str(counter)] = materia1
            ws['K'+str(counter)] = nota1
            ws['L'+str(counter)] = materia2
            ws['M'+str(counter)] = nota2

            counter += 1
            continue

        #Si tenia 1 inscritas
        match_historico = re.search(r"(\d{1}/\d{4})\s*(\d+-[A-Za-z0-9-@#%])\s*\(\s*(\d+)\s*\)[a-z]*\s*([A-Za-z0-9-@#%]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1 = match_historico.groups()
            ws['C'+str(counter)] = sigla_carrera
            
            ws['H'+str(counter)] = s_ano
            ws['I'+str(counter)] = pps
            ws['J'+str(counter)] = materia1
            ws['K'+str(counter)] = nota1

            counter += 1
            continue

    print(f"El documento cuenta con {counter - 1} filas hechas")
    print(f"El documento cuenta con {nombre_contador} estudiantes")

def ajustar_tamanio_columnas(ws):
    """
    Ajusta el tamaño de las columnas de un DataFrame en un archivo Excel.
    
    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): Hoja de cálculo de Excel
    """
    for col in ws.columns:
        max_length = 0
        column_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = (max_length + 4)
        ws.column_dimensions[column_letter].width = adjusted_width
    print("Tamaño de columnas ajustado correctamente.")

def formatear_celdas(ws):
    """
    Colorea las celdas de un DataFrame en un archivo Excel.
    
    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): Hoja de cálculo de Excel
    """
    fill = PatternFill(start_color="FF5B9BD5", end_color="FF5B9BD5", fill_type = "solid")
    font = Font(color="FFFFFFFF")
    alignment = Alignment(horizontal='center', vertical='center')

    ws['A1'].fill = fill
    ws['B1'].fill = fill
    ws['C1'].fill = fill
    ws['D1'].fill = fill
    ws['E1'].fill = fill
    ws['F1'].fill = fill
    ws['G1'].fill = fill
    ws['H1'].fill = fill
    ws['I1'].fill = fill
    ws['J1'].fill = fill
    ws['K1'].fill = fill
    ws['L1'].fill = fill
    ws['M1'].fill = fill
    ws['N1'].fill = fill
    ws['O1'].fill = fill
    ws['P1'].fill = fill
    ws['Q1'].fill = fill
    ws['R1'].fill = fill
    ws['S1'].fill = fill
    ws['T1'].fill = fill
    ws['U1'].fill = fill
    ws['V1'].fill = fill
    ws['W1'].fill = fill
    ws['X1'].fill = fill
    ws['Y1'].fill = fill
    ws['Z1'].fill = fill
    ws['AA1'].fill = fill
    ws['AB1'].fill = fill
    ws['AC1'].fill = fill

    ws['A1'].font = font
    ws['B1'].font = font
    ws['C1'].font = font
    ws['D1'].font = font
    ws['E1'].font = font
    ws['F1'].font = font
    ws['G1'].font = font
    ws['H1'].font = font
    ws['I1'].font = font
    ws['J1'].font = font
    ws['K1'].font = font
    ws['L1'].font = font
    ws['M1'].font = font
    ws['N1'].font = font
    ws['O1'].font = font
    ws['P1'].font = font
    ws['Q1'].font = font
    ws['R1'].font = font
    ws['S1'].font = font 
    ws['T1'].font = font  
    ws['U1'].font = font 
    ws['V1'].font = font  
    ws['W1'].font = font 
    ws['X1'].font = font  
    ws['Y1'].font = font 
    ws['Z1'].font = font  
    ws['AA1'].font = font 
    ws['AB1'].font = font  
    ws['AC1'].font = font

    print("Celdas header coloreadas correctamente.")

    ws['A1'].alignment = alignment
    ws['B1'].alignment = alignment
    ws['C1'].alignment = alignment
    ws['D1'].alignment = alignment
    ws['E1'].alignment = alignment
    ws['F1'].alignment = alignment
    ws['G1'].alignment = alignment
    ws['H1'].alignment = alignment
    ws['I1'].alignment = alignment
    ws['J1'].alignment = alignment
    ws['K1'].alignment = alignment
    ws['L1'].alignment = alignment
    ws['M1'].alignment = alignment
    ws['N1'].alignment = alignment
    ws['O1'].alignment = alignment
    ws['P1'].alignment = alignment
    ws['Q1'].alignment = alignment
    ws['R1'].alignment = alignment
    ws['S1'].alignment = alignment
    ws['T1'].alignment = alignment
    ws['U1'].alignment = alignment
    ws['V1'].alignment = alignment
    ws['W1'].alignment = alignment
    ws['X1'].alignment = alignment
    ws['Y1'].alignment = alignment
    ws['Z1'].alignment = alignment
    ws['AA1'].alignment = alignment
    ws['AB1'].alignment = alignment
    ws['AC1'].alignment = alignment

    print("Celdas header alineadas correctamente.")
    


def main():
    wb = Workbook()
    ws = wb.active

    ruta_archivo = './data/input.txt'
    lineas = leer_archivo_txt(ruta_archivo)

    procesar_datos(lineas, ws, wb)
    
    ajustar_tamanio_columnas(ws)

    formatear_celdas(ws)
    
    wb.save("output.xlsx")
    print("Documento output.xlsx generado correctamente.")



if __name__ == "__main__":
    main()