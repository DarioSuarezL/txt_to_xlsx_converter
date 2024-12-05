import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

def leer_archivo_txt(ruta_archivo):
    """
    Lee el contenido de un archivo de texto.
    
    Args:
        ruta_archivo (str): Ruta completa del archivo txt a leer
    
    Returns:
        list: Lista de líneas del archivo
    """
    try:
        with open(ruta_archivo, 'r') as archivo:
            return [linea.strip() for linea in archivo.readlines()]
    except Exception as e:
        print(f'Error al leer el archivo: {e}')
        return []

def procesar_datos(lineas):
    """
    Procesa las líneas de un archivo de texto.
    
    Args:
        lineas (list): Lista de líneas del archivo

    Returns:
        pd.DataFrame: DataFrame con los datos procesados
    """

    wb = Workbook()
    ws = wb.active

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

    counter = 2 #fila que inicia

    registro = None
    alumno = None
    sigla = None
    carrera = None
    ingreso = None
    ppac = None
    ppace = None
    ended_line = False

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

    for linea in lineas:

        if linea in lineas_a_omitir:
            continue

        #Extrae datos del estudiante
        match_estudiante = re.search(r'#\s+\d+:\s+Estudiante:(\d+)\s+-\s+(.+)', linea)
        if match_estudiante:
            registro, alumno = match_estudiante.groups()
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
            ws['C'+str(counter)] = sigla
            ws['D'+str(counter)] = carrera
            ws['E'+str(counter)] = ingreso
            ws['F'+str(counter)] = ppac
            ws['G'+str(counter)] = ppace


        #Extrae datos del histórico
        #Si tenia 7 inscritas
        match_historico = re.search(r"(\d{1,2}/\d{4})\s*(\d+-\d+)\s*\(\s*(\d+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-])\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3, materia4, nota4, materia5, nota5, materia6, nota6, materia7, nota7 = match_historico.groups()
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
        match_historico = re.search(r"(\d{1,2}/\d{4})\s*(\d+-\d+)\s*\(\s*(\d+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3, materia4, nota4, materia5, nota5, materia6, nota6 = match_historico.groups()
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
        match_historico = re.search(r"(\d{1,2}/\d{4})\s*(\d+-\d+)\s*\(\s*(\d+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3, materia4, nota4, materia5, nota5 = match_historico.groups()
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
        match_historico = re.search(r"(\d{1,2}/\d{4})\s*(\d+-\d+)\s*\(\s*(\d+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3, materia4, nota4 = match_historico.groups()
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
        match_historico = re.search(r"(\d{1,2}/\d{4})\s*(\d+-\d+)\s*\(\s*(\d+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2, materia3, nota3 = match_historico.groups()
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
        match_historico = re.search(r"(\d{1,2}/\d{4})\s*(\d+-\d+)\s*\(\s*(\d+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1, materia2, nota2 = match_historico.groups()
            ws['H'+str(counter)] = s_ano
            ws['I'+str(counter)] = pps
            ws['J'+str(counter)] = materia1
            ws['K'+str(counter)] = nota1
            ws['L'+str(counter)] = materia2
            ws['M'+str(counter)] = nota2

            counter += 1
            continue

        #Si tenia 1 inscritas
        match_historico = re.search(r"(\d{1,2}/\d{4})\s*(\d+-\d+)\s*\(\s*(\d+)\s*\)\s*([A-Za-z0-9-]+)\s*\(\s*([A-Za-z0-9-]+)\s*\)", linea)
        if match_historico:
            s_ano, sigla_carrera, pps, materia1, nota1 = match_historico.groups()
            ws['H'+str(counter)] = s_ano
            ws['I'+str(counter)] = pps
            ws['J'+str(counter)] = materia1
            ws['K'+str(counter)] = nota1

            counter += 1
            continue


    
    wb.save("output.xlsx")
    print(f"El documento cuenta con {counter - 1} registros realizados")
    print("Documento output.xlsx generado correctamente.")






ruta_archivo = './data/input.txt'
lineas = leer_archivo_txt(ruta_archivo)
procesar_datos(lineas)