# -*- coding: utf-8 -*-
"""
Created on Thu Mar 16 16:21:10 2023

@author: rafael.cozmar
"""
import datetime as dt
import pandas as pd
# Nombres de columnas que generan errores normalmente

# Diccionario para poder combinar A3 con A4
dicc_a3xa4 = {'Capacidad': 'CAPACIDAD (plazas)',
              'TIPO_BUS': 'Tipologia',
              'PERIODO_INICIO': 'MH',
              'TIPO_DIA': 'TIPO DIA',
              'DISTANCIA (KM)': 'DISTANCIA'}
    
# Diccionario que corrige tipico error en columnas del A4  TH
dicc_a4th = {'TIPO DIA': 'TIPO_DIA',
             'DISTANCIA (KM)': 'DISTANCIA',
             'DISTANCIA (km)': 'DISTANCIA'}

# Funcion para pasar a string el código TS si es que no lo está
def codts_tostring(CTS):
    if CTS.__class__.__name__ == 'str':
        return CTS
    else:
        return str(CTS)

    
def isNowInTimePeriod(startTime, endTime, nowTime): 
    if startTime < endTime: 
        return nowTime>=startTime and nowTime<=endTime
    else: 
        #Over midnight: 
        return nowTime >= startTime or nowTime <= endTime

da1 = 'C:/Users/rafael.cozmar/OneDrive - Directorio de Transportes Público Metropolitano/Escritorio/PO 2023(27Feb al 31Dic) U8 - Anexo 1.xlsx'
da3 = 'C:/Users/rafael.cozmar/OneDrive - Directorio de Transportes Público Metropolitano/Escritorio/PO 2023(27Feb al 31Dic) U8 - Anexo 3.xlsx'
da4 = 'C:/Users/rafael.cozmar/OneDrive - Directorio de Transportes Público Metropolitano/Escritorio/PO 2023(27Feb al 31Dic) U8 - Anexo 4.xlsx'


def fix_string_as_time(ST):
    # Si por alguna razón está como string
    if ST.__class__.__name__ == 'str':
        return (dt.datetime.strptime(ST,'%H:%M:%S')).time()
    elif ST.__class__.__name__ == 'time':
        return ST
    # Si por alguna razon está como fechahora, hay que extraer solamente la hora
    elif ST.__class__.__name__ == 'datetime':
        return ST.time()
    elif ST.__class__.__name__ == 'Timestamp':
        return ST.time()
    else:
        print('Existe hora corrupta')

# def fix_agrupacionA3(MHA):
#     agrupaciones_posibles=list(range(1,49)).append('--')
#     if MHA in agrupaciones_posibles:
#         return MHA
#     else:
#         print('Existe una agrupación MH corrupta en anexo 3')
#         return int(MHA)

def rev_tipodiaA3(tpo_dia):
    tipos_dia = ['Laboral','Sábado','Domingo']
    if tpo_dia in tipos_dia:
        return tpo_dia
    else:
        print('Existe un tipo día diferente de: Laboral, Sábado o Domingo o está mal escrito')
        return tpo_dia


def rev_sentido(sentido):
    sentidos = ['Ida','Ret']
    if sentido in sentidos:
        return sentido
    else: #En algun momento seria bueno corregirlo..
        print('Existe un sentido mal escrito en A3')
        return sentido
    
    
# Funcion para truncar, actualmente no se usa pero puede servir
def truncate(number: float, digits: int) -> float:
    pow10 = 10 ** digits
    return number * pow10 // 1 / pow10

## en algun momento una funcion para detectar 2 decimales...
def detectar_decimales(num):
    if num != round(num, 2):
        print('Existen datos no redondeados a 2 dígitos (igualmente sólo consideraremos los 2 dígitos')
        return round(num, 2)
    else:
        return num
    
def rev_tipo_evento(tipo_evento):
    tipos_evento = ['I01','C01','V01','V02','V03','V04','V05','V06','L01','R01','R02','D01','E01','F01']
    if tipo_evento in tipos_evento:
        return tipo_evento
    else:
        print('Existe un tipo de evento diferente a los permitidos o está mal escrito')
        return tipo_evento
    
# Esta funcion se debe aplicar en cada celda del A1
def convertir_horario_dia_siguiente(HI, HT):
    # Requiere diferente tratamiento si es un datetime o un time...
    if HI.__class__.__name__ == 'time' and HT.__class__.__name__ == 'time':
        if HI < HT:
            HT_ARREG = dt.datetime(1995, 6, 13, HT.hour, HT.minute, HT.second)
            return (HT_ARREG)
        elif HT < HI and dt.time(0, 0) <= HT: #Condicion que implica que las expediciones siguientes pasan al otro dia
            HT_ARREG = dt.datetime(1995, 6, 14, HT.hour, HT.minute, HT.second)
            return (HT_ARREG)
    elif HI.__class__.__name__ == 'time' and HT.__class__.__name__ != 'time':
        HI_mod = HI
        HT_mod = HT.time()
        convertir_horario_dia_siguiente(HI_mod,HT_mod)
    elif HI.__class__.__name__ != 'time' and HT.__class__.__name__ == 'time':
        HI_mod = HI.time()
        HT_mod = HT
        convertir_horario_dia_siguiente(HI_mod, HT_mod)
    else:
        HI_mod = HI.time()
        HT_mod = HT.time()
        convertir_horario_dia_siguiente(HI_mod, HT_mod)
        
##Recibe un horario de inicio y término en formato datetime.time
def convertir_horario_dia_siguiente_func(HI, HT):
    # Requiere diferente tratamiento si es un datetime o un time...
    HI = fix_string_as_time(HI)
    HT = fix_string_as_time(HT)
    if HI < HT:
        HT_ARREG = dt.datetime(1995, 6, 13, HT.hour, HT.minute, HT.second)
        return (HT_ARREG)
    elif HT < HI and dt.time(0, 0) <= HT: # Condicion que implica que las expediciones siguientes pasan al otro dia
        HT_ARREG = dt.datetime(1995, 6, 14, HT.hour, HT.minute, HT.second)
        return (HT_ARREG)

##Una función para tratar el A3 como se debe, con los formatos necesarios para su correcto procesamiento

#Recibe una dirección de dónde está el A3, y entrega un anexo 3 (la parte útil) con los formatos correctos
def formatear_a3(A3_dir):
    A3_sf = pd.read_excel(A3_dir, sheet_name="Parámetros", header=6)
    A3_sf['CODIGO TS SERVICIO'] = A3_sf.apply(lambda row: codts_tostring(row['CODIGO TS SERVICIO']), axis=1)
    A3_sf['CODIGO USUARIO SERVICIO'] = A3_sf.apply(lambda row: codts_tostring(row['CODIGO USUARIO SERVICIO']), axis=1)
    A3_sf['SENTIDO'] = A3_sf.apply(lambda row: rev_sentido(row['SENTIDO']), axis=1)
    A3_sf['TIPO DIA'] = A3_sf.apply(lambda row: rev_tipodiaA3(row['TIPO DIA']), axis=1)
    # Sentido y Tipo día deben siempre decir las mismas cosas
    A3_sf['MH'] = A3_sf.apply(lambda row: fix_string_as_time(row['MH']), axis=1)
    A3_sf['VELOCIDAD (Km/hra)'] = A3_sf.apply(lambda row: detectar_decimales(row['VELOCIDAD (Km/hra)']), axis=1)
    A3_sf['DISTANCIA BASE (Km)'] = A3_sf.apply(lambda row: detectar_decimales(row['DISTANCIA BASE (Km)']), axis=1)
    A3_sf['DISTANCIA TOTAL (POB+POI) (Km)'] = A3_sf.apply(lambda row: detectar_decimales(
                                                                    row['DISTANCIA TOTAL (POB+POI) (Km)']), axis=1)
    A3_sf['N° SALIDAS'] = A3_sf.apply(lambda row: int(row['N° SALIDAS']), axis=1)
    A3_sf['CAPACIDAD (plazas)'] = A3_sf.apply(lambda row: int(row['CAPACIDAD (plazas)']), axis=1)
    # A3_sf['AGRUPACIÓN MH']=A3_sf.apply(lambda row: fix_agrupacionA3(row['AGRUPACIÓN MH']),axis=1)
    
    # A3_sf['FAMILIA DE SERVICIOS']= ##no se usa por ahora..
    # A3_sf['INDICADOR TIEMPO DE ESPERA']= ##no se usa por ahora..
    A3_formateado = A3_sf
    return A3_formateado

def formatear_a1(A1_dir):
    A1_sf = pd.read_excel(A1_dir, sheet_name="Horarios", header=6)
    A1_sf.rename(columns={'CÓDIGO Usuario servicio': 'CODIGO USUARIO SERVICIO',
                          'Sentido': 'SENTIDO',
                          'TIPO_DIA':'TIPO DIA'}, inplace=True)
    
    A1_sf['CODIGO TS SERVICIO'] = A1_sf.apply(lambda row: codts_tostring(row['CODIGO TS SERVICIO']), axis=1)
    A1_sf['CODIGO USUARIO SERVICIO'] = A1_sf.apply(lambda row: codts_tostring(row['CODIGO USUARIO SERVICIO']), axis=1)
    A1_sf['SENTIDO'] = A1_sf.apply(lambda row: rev_sentido(row['SENTIDO']), axis=1)
    A1_sf['TIPO DIA'] = A1_sf.apply(lambda row: rev_tipodiaA3(row['TIPO DIA']), axis=1)
    A1_sf['TRAMO HORARIO'] = A1_sf.apply(lambda row: int(row['TRAMO HORARIO']), axis=1)
    A1_sf['HORA INICIO'] = A1_sf.apply(lambda row: fix_string_as_time(row['HORA INICIO']), axis=1)
    A1_sf['HORA INICIO'] = A1_sf.apply(lambda row: dt.datetime(1995, 6, 13, row['HORA INICIO'].hour,
                                                               row['HORA INICIO'].minute, row['HORA INICIO'].second),
                                       axis=1)
    A1_sf['HORA TERMINO'] = A1_sf.apply(lambda row: fix_string_as_time(row['HORA TERMINO']), axis=1)
    A1_sf['HORA TERMINO'] = A1_sf.apply(lambda row: convertir_horario_dia_siguiente_func(row['HORA INICIO'],
                                                                                         row['HORA TERMINO']), axis=1)
    #A1_sf['HORA TERMINO']=A1_sf.apply(lambda row: dt.datetime(1995,6,13,row['HORA TERMINO'].hour,row['HORA TERMINO'].minute,row['HORA TERMINO'].second).time(),axis=1)
    A1_formateado = A1_sf
    return A1_formateado


def formatear_a1_vRH(A1_dir):
    A1_sf = pd.read_excel(A1_dir, sheet_name="Horarios", header=6)
    A1_sf.rename(columns={'CÓDIGO Usuario servicio': 'CODIGO USUARIO SERVICIO',
                          'Sentido': 'SENTIDO',
                          'TIPO_DIA': 'TIPO DIA'}, inplace=True)
    
    A1_sf['CODIGO TS SERVICIO'] = A1_sf.apply(lambda row: codts_tostring(row['CODIGO TS SERVICIO']), axis=1)
    A1_sf['CODIGO USUARIO SERVICIO'] = A1_sf.apply(lambda row: codts_tostring(row['CODIGO USUARIO SERVICIO']), axis=1)
    A1_sf['SENTIDO'] = A1_sf.apply(lambda row: rev_sentido(row['SENTIDO']), axis=1)
    A1_sf['TIPO DIA'] = A1_sf.apply(lambda row: rev_tipodiaA3(row['TIPO DIA']), axis=1)
    A1_sf['TRAMO HORARIO'] = A1_sf.apply(lambda row: int(row['TRAMO HORARIO']), axis=1)
    A1_sf['HORA INICIO'] = A1_sf.apply(lambda row: fix_string_as_time(row['HORA INICIO']), axis=1)
    A1_sf['HORA INICIO'] = A1_sf.apply(lambda row: dt.datetime(1995, 6, 13, row['HORA INICIO'].hour,
                                                               row['HORA INICIO'].minute,
                                                               row['HORA INICIO'].second).time(), axis=1)
    A1_sf['HORA TERMINO'] = A1_sf.apply(lambda row: fix_string_as_time(row['HORA TERMINO']), axis=1)
    A1_sf['HORA TERMINO'] = A1_sf.apply(lambda row: convertir_horario_dia_siguiente_func(row['HORA INICIO'],
                                                                                         row['HORA TERMINO']).time(),
                                        axis=1)
    #A1_sf['HORA TERMINO']=A1_sf.apply(lambda row: dt.datetime(1995,6,13,row['HORA TERMINO'].hour,row['HORA TERMINO'].minute,row['HORA TERMINO'].second).time(),axis=1)
    A1_formateado = A1_sf
    return A1_formateado
# UNIDAD DE SERVICIO	 ESCENARIO	BUS_LOGICO	CODIGO TS SERVICIO	SENTIDO	TIPO DIA	TIPO_EVENTO	HORA_INICIO	PERIODO_INICIO	HORA_FIN	PERIODO_FIN	DURACION	PUNTO_INICIO	PUNTO_FIN	DISTANCIA	TIPO_BUS

def formatear_a4th(A4_dir):
    A4_sf = pd.read_excel(A4_dir, sheet_name="Tabla Horaria", header=6)
    A4_sf = A4_sf.rename(columns=dicc_a4th)
    
    A4_sf['CODIGO TS SERVICIO'] = A4_sf.apply(lambda row: codts_tostring(row['CODIGO TS SERVICIO']), axis=1)
    A4_sf['SENTIDO'] = A4_sf.apply(lambda row: rev_sentido(row['SENTIDO']), axis=1)
    A4_sf['TIPO_DIA'] = A4_sf.apply(lambda row: rev_tipodiaA3(row['TIPO_DIA']), axis=1)
    A4_sf['TIPO_EVENTO'] = A4_sf.apply(lambda row: rev_tipo_evento(row['TIPO_EVENTO']), axis=1)
    A4_sf['HORA_INICIO'] = A4_sf.apply(lambda row: fix_string_as_time(row['HORA_INICIO']), axis=1)
    A4_sf['PERIODO_INICIO'] = A4_sf.apply(lambda row: fix_string_as_time(row['PERIODO_INICIO']), axis=1)
    A4_sf['HORA_FIN'] = A4_sf.apply(lambda row: fix_string_as_time(row['HORA_FIN']), axis=1)
    A4_sf['PERIODO_FIN'] = A4_sf.apply(lambda row: fix_string_as_time(row['PERIODO_FIN']), axis=1)
    A4_sf['DURACION'] = A4_sf.apply(lambda row: fix_string_as_time(row['DURACION']), axis=1)
    A4_sf['DISTANCIA'] = A4_sf.apply(lambda row: detectar_decimales(row['DISTANCIA']), axis=1)
    A4_sf['TIPO_BUS'] = A4_sf.apply(lambda row: codts_tostring(row['TIPO_BUS']), axis=1)
    A4_formateado = A4_sf.copy()
    return A4_formateado

def obtener_dicc_cap(A4_dir):
    Dicc_cap_bus = pd.read_excel(A4_dir,sheet_name="Diccionario",header=1)
    indices_salto = Dicc_cap_bus.loc[Dicc_cap_bus["Tipo Bus"].isnull()].index
    return Dicc_cap_bus.head(indices_salto[0]) ##Obtiene los parametros hasta el primer salto jeje
    

##Función que dado: intervalo programado,  Intervalo mínimo de la media hora agrupada (MHA) e Intervalo máximo de la MHA, 
# indica si el intervalo programado respeta las condiciones del manual de programación
def intervalo_aceptable(Imin, Imax, Iprog):
    diferencia = Imax-Imin
    if diferencia < dt.timedelta(minutes=2) and Iprog <= dt.timedelta(minutes=10):
        return True
    elif diferencia < dt.timedelta(minutes=4) and (dt.timedelta(minutes=10) < Iprog <= dt.timedelta(minutes=20)):
        return True
    elif diferencia < dt.timedelta(minutes=6) and Iprog > dt.timedelta(minutes=20):
        return True
    else:
        return False
    

##Esta función permite imputar el día en que inicia la expedición, así las expediciones que comienzan después de las 00 horas del día siguiente pueden ser ordenadas y se les imputa el día siguiente
def convertir_horasalida_nextday(hora, minutos, segundos, horario_ini, horario_fin):
#Termina despues de las 0 horas y la expedicion empieza despues de las 00 horas y esas expediciones son antes de la hora de inicio si no fuera el contexto de un reloj
    if horario_fin.time() >= dt.time(hour=0) and dt.time(hour=0) <= dt.time(hora, minutos,
                                                                            segundos) < horario_ini.time():
        return dt.datetime(1995, 6, 14, hora, minutos, segundos)
    else:
        return dt.datetime(1995, 6, 13, hora, minutos, segundos)

