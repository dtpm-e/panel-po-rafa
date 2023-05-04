# -*- coding: utf-8 -*-
"""
Created on Mon Apr 10 10:55:41 2023

@author: rafael.cozmar
"""

import pandas as pd
import os as os
import numpy as np
np.set_printoptions(threshold=np.inf)


from parches import formatear_a1_vRH
from parches import formatear_a1
from parches import formatear_a4th
from parches import isNowInTimePeriod
from parches import codts_tostring
from parches import formatear_a3
import datetime as dt

def param_faltantes_A3(A3_direccion,ruta_OUTPUT):
    A3=pd.read_excel(A3_direccion,sheet_name="Parámetros",header=6) # Parametros de operacion  
    print("Buscando parametros faltantes en anexo 3....")
    #Revisar ceros o numeros negativos
    Datos_con_ceros=((A3[( A3["N° SALIDAS"]>0) & ((A3["VELOCIDAD (Km/hra)"]<=0 )|(A3["DISTANCIA BASE (Km)"]<=0 )|(A3["DISTANCIA TOTAL (POB+POI) (Km)"]<=0) | (A3["CAPACIDAD (plazas)"]<=0))]))
                     
    #Revisar datos faltantes si es que hay salidas
    datos_nan_velocidad=((A3[( A3["N° SALIDAS"]>0) & (A3["VELOCIDAD (Km/hra)"].isnull().values)]))
    datos_nan_dist_base=((A3[( A3["N° SALIDAS"]>0) & (A3["DISTANCIA BASE (Km)"].isnull().values)]))
    datos_nan_dist_total=((A3[( A3["N° SALIDAS"]>0) & (A3["DISTANCIA TOTAL (POB+POI) (Km)"].isnull().values)]))
    datos_nan_capacidad=((A3[( A3["N° SALIDAS"]>0) & (A3["CAPACIDAD (plazas)"].isnull().values)]))
    ##Ningun SSTD debe tener celdas vacias..
    datos_nan=A3.loc[A3.isnull().any(axis=1)]
    filas_con_formulas=pd.DataFrame()
    ###ARCHIVO DE SALIDA QUE INDIQUE LO QUE ESTÁ MALO
    ##Moverse a salida
    os.chdir(ruta_OUTPUT)
    ##GUARDAR COMO EXCEL
    #Creamos planilla donde se guardará
    
    planilla_param_faltantes=pd.ExcelWriter('Parametros faltantes.xlsx')
    
    #agregamos a las hojas
    Datos_con_ceros.to_excel(planilla_param_faltantes,sheet_name="Serv con ceros ")
    datos_nan_velocidad.to_excel(planilla_param_faltantes,sheet_name="Serv sin datos en velocidad")
    datos_nan_dist_base.to_excel(planilla_param_faltantes,sheet_name="Serv sin datos en dist.base")
    datos_nan_dist_total.to_excel(planilla_param_faltantes,sheet_name="Serv sin datos en distancia")
    datos_nan_capacidad.to_excel(planilla_param_faltantes,sheet_name="Serv sin datos en capacidad")
    datos_nan.to_excel(planilla_param_faltantes,sheet_name="Fila con espacios vacios")
    filas_con_formulas.to_excel(planilla_param_faltantes,sheet_name="Fila con fórmulas")
    planilla_param_faltantes.close()

    print('Proceso Finalizado')
    print('Revise archivo "Parametros faltantes.xlsx", si están vacíos son anexos consistentes')
    print("Ubicado en:")
    print(ruta_OUTPUT)


def param_faltantes_A4(A4_direccion,ruta_OUTPUT):
    print("Buscando filas con celdas vacias en Tabla Horaria Anexo 4....")
    ##Hay que eliminar las columnas que tienen otras cosas

    #Abriendo Anexos Nuevos
    A4th=pd.read_excel(A4_direccion,sheet_name="Tabla Horaria",header=6)
    datos_nan=A4th.loc[A4th.isnull().any(axis=1)]

    os.chdir(ruta_OUTPUT)
    planilla_param_faltantes=pd.ExcelWriter('Parametros faltantes A4.xlsx')

    #agregamos a las hojas
    datos_nan.to_excel(planilla_param_faltantes,sheet_name="Expediciones falta dato")

    #guardamos
    planilla_param_faltantes.close()

    print('Proceso Finalizado')
    print('Revise archivo "Parametros faltantes A4.xlsx", si está vacíos significa que la Tabla Horaria tiene todos los datos')
    print("Ubicado en:")
    print(ruta_OUTPUT)
    
def revisar_horarios_a1_a4(A1_direccion,A4_direccion,ruta_OUTPUT):

    print("Revisando horarios.... ")



    #Abriendo Anexos Nuevos

    A1=formatear_a1_vRH(A1_direccion)
    A4th=formatear_a4th(A4_direccion)

    ##Rename porque no están con la misma notación para hacer merge despues
    A1=A1.rename(columns={'Sentido':'SENTIDO'})
    A4th=A4th.rename(columns={'TIPO_DIA':'TIPO DIA'})



    ##Serv Sentido TD con 2 o mas horarios de operación discontinuos (serv operado por tramos horarios) manipulamos las tablas para poder colocar el segundo horario a la derecha en un mismo sstd
    A1_tramos_horarios=A1[A1.duplicated(['CODIGO TS SERVICIO','SENTIDO','TIPO DIA'],keep=False)]
    A1_tramos_horarios=pd.merge(A1_tramos_horarios,A1_tramos_horarios,on=['CODIGO TS SERVICIO','SENTIDO','TIPO DIA'])
    A1_tramos_horarios=A1_tramos_horarios[A1_tramos_horarios['TRAMO HORARIO_x']!=A1_tramos_horarios['TRAMO HORARIO_y']]
    A1_tramos_horarios=A1_tramos_horarios[A1_tramos_horarios.duplicated(['CODIGO TS SERVICIO','SENTIDO','TIPO DIA'])]
     ##Los sstd que si operan por tramos horarios
    A1=A1[~A1.duplicated(['CODIGO TS SERVICIO','SENTIDO','TIPO DIA'],keep=False)] ###Sacamos los que operan por tramos de horarios

    ##
    a4th_com=A4th[(A4th["TIPO_EVENTO"]=="C01") | (A4th["TIPO_EVENTO"]=="C01 ") ] ##Filtramos expediciones comerciales
    a4a1=pd.merge(a4th_com,A1,on=['CODIGO TS SERVICIO','SENTIDO','TIPO DIA'],suffixes=['_a4','_a1'])
    ###Expediciones que salen antes de la hora de inicio
    ifh=a4a1[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA',"HORA INICIO", "HORA TERMINO" , "HORA_INICIO"]] #IFH= INICIO FUERA DE HORARIO #HORA_INICIO ES LA HORA DE SALIDA DE LA EXPEDICION HORA INICIO ES EL HORARIO DESDE QUE SE REALIZA EL SERV
    ###De las anteriores, sacamos las que se debe a que pasaron al otro día
    #IFH=IFH[IFH['HORA INICIO']<=IFH['HORA_FIN']]
    ifh["¿Respeta horarios?"] = ifh.apply(lambda row: isNowInTimePeriod(row['HORA INICIO'], row['HORA TERMINO'], row['HORA_INICIO']),axis=1)
    ##Los servicios con horario de operación discontinuo saldrán igual en esta tabla
    ##BUSCAR SERV CON HORARIO DISCONTINUO PQ TIENEN REPETIDO EL HORARIO DE OPERACION
    a4a1_tramos_horarios=pd.merge(a4th_com,A1_tramos_horarios,on=['CODIGO TS SERVICIO','SENTIDO','TIPO DIA'],suffixes=['_a4','_a1_tramos_horario'])
    ifh_tramos_horarios=a4a1_tramos_horarios[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA',"HORA INICIO_x", "HORA TERMINO_x","HORA INICIO_y", "HORA TERMINO_y" , "HORA_INICIO"]]
    if len(ifh_tramos_horarios)>0:
        ifh_tramos_horarios["¿Respeta horario_x?"] = ifh_tramos_horarios.apply(lambda row: isNowInTimePeriod(row['HORA INICIO_x'], row['HORA TERMINO_x'], row['HORA_INICIO']),axis=1)
        ifh_tramos_horarios["¿Respeta horario_y?"] = ifh_tramos_horarios.apply(lambda row: isNowInTimePeriod(row['HORA INICIO_y'], row['HORA TERMINO_y'], row['HORA_INICIO']),axis=1)
        ifh_tramos_horarios["¿Respeta alguno de sus horarios?"]=ifh_tramos_horarios.apply(lambda row: row["¿Respeta horario_x?"] | row["¿Respeta horario_y?"],axis=1)
        ifh_tramos_horarios=ifh_tramos_horarios[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA',"HORA INICIO_x", "HORA TERMINO_x","HORA INICIO_y", "HORA TERMINO_y" , "HORA_INICIO","¿Respeta alguno de sus horarios?"]]
    else:
        ifh_tramos_horarios["¿Respeta alguno de sus horarios?"]=True

    ##Entregamos solamente sstd con problemas
    ifh=ifh[~ifh["¿Respeta horarios?"]]
    ifh_tramos_horarios=ifh_tramos_horarios[~ifh_tramos_horarios["¿Respeta alguno de sus horarios?"]]

    #Cambiamos nombres de columnas para que sean más claros
    cambios_nombre_output={'HORA_INICIO':'HORA EXPEDICION (A4)','HORA INICIO_x':'Horario 1 inicio servicio (A1)','HORA INICIO_y':'Horario 2 inicio servicio (A1)','HORA TERMINO_x':'Horario 1 termino servicio (A1)','HORA TERMINO_y':'Horario 2 termino servicio (A1)','HORA INICIO':'Horario inicio servicio (A1)','HORA TERMINO':'Horario termino servicio (A1)'}

    ifh.rename(columns=cambios_nombre_output,inplace=True)
    ifh_tramos_horarios.rename(columns=cambios_nombre_output,inplace=True)

    #Todos_sstd[~Todos_sstd.duplicated(['CODIGO TS SERVICIO'])]['CODIGO TS SERVICIO']
    ###Encontrar que la ultima salida sea la primera y ultima MH
    sstdexp=A4th[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','PERIODO_INICIO']]
    sstdexp2=A4th[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','PERIODO_INICIO']]
    sstdexp.rename(columns={'PERIODO_INICIO':'HORA INICIO'},inplace=True)
    # sstdexp=sstdexp[~sstdexp.duplicated(['CODIGO TS SERVICIO', 'SENTIDO', 'TIPO DIA', 'HORA INICIO'])]
    hini_sin_salida=[]
    hfin_sin_salida=[]

    for i,row in A1.iterrows():
        servicio=row['CODIGO TS SERVICIO']
        sentido=row['SENTIDO']
        td=row['TIPO DIA']
        hi=row['HORA INICIO']
        ht=row['HORA TERMINO']
        expedicion_hora_inicio=A4th[(A4th['CODIGO TS SERVICIO']==servicio) & (A4th['SENTIDO']==sentido) & (A4th['TIPO DIA']==td) & (A4th['HORA_INICIO']==hi)]
        expedicion_hora_termino=A4th[(A4th['CODIGO TS SERVICIO']==servicio) & (A4th['SENTIDO']==sentido) & (A4th['TIPO DIA']==td) & (A4th['HORA_INICIO']==ht)]
        if len(expedicion_hora_inicio.axes[0])==0:
            sin_hi=pd.DataFrame({'CODIGO TS SERVICIO':[row['CODIGO TS SERVICIO']],'SENTIDO':[row['SENTIDO']],'TIPO DIA':[row['TIPO DIA']],'No hay expedición en esta Hora Inicio':[hi]})
            hini_sin_salida.append(sin_hi)
        if len(expedicion_hora_termino.axes[0])==0:
            sin_ht=pd.DataFrame({'CODIGO TS SERVICIO':[row['CODIGO TS SERVICIO']],'SENTIDO':[row['SENTIDO']],'TIPO DIA':[row['TIPO DIA']],'No hay expedición en esta Hora Termino':[ht]})
            hfin_sin_salida.append(sin_ht)
        
    if len(hini_sin_salida)>0:
        reportar_sin_hi=pd.concat(hini_sin_salida)
    else:
        reportar_sin_hi=pd.DataFrame()
        
    if len(hfin_sin_salida)>0:
        reportar_sin_ht=pd.concat(hfin_sin_salida)
    else:
        reportar_sin_ht=pd.DataFrame()
    ##Revisar que esté sstd de la tabla horaria en el A1..
    ##Moverse a salida
    os.chdir(ruta_OUTPUT)
    ##GUARDAR COMO EXCEL
    #Creamos planilla donde se guardará
    planilla_diferencias_horarios=pd.ExcelWriter('Diferencia Horarios.xlsx')

    ifh.to_excel(planilla_diferencias_horarios,sheet_name="Serv inconsist SIN Tramos Hor")
    reportar_sin_hi.to_excel(planilla_diferencias_horarios,sheet_name="Revisar HoraIni con primera exp")
    reportar_sin_ht.to_excel(planilla_diferencias_horarios,sheet_name="Revisar HoraTer con ultima exp")
    ifh_tramos_horarios.to_excel(planilla_diferencias_horarios,sheet_name="Serv inconsist CON Tramos Hor")
    planilla_diferencias_horarios.close()

    print('Proceso Finalizado')
    print('Revise archivo "Diferencia Horarios.xlsx", si están vacíos son anexos consistentes')
    print("Ubicado en:")
    print(ruta_OUTPUT)




def contar_salidas(A3_direccion,A4_direccion,ruta_OUTPUT):

    print("Contando salidas....")

    ####LECTURA ANEXOS
    A3=formatear_a3(A3_direccion) # Parametros de operacion
    A4TH=formatear_a4th(A4_direccion)  # Tabla Horaria
    
    
    A4TH=A4TH[(A4TH['TIPO_EVENTO']=='C01') | (A4TH['TIPO_EVENTO']=='C01 ')] ##fILTRAR EXPEDICIONES COMERCIALES
    # agrupacionmhxsstpo=A3[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','MH','AGRUPACIÓN MH']] #####Obtenemos diccionario con las agrupaciones de MH por sstd
    salidas_xmh_a3=A3[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','MH','AGRUPACIÓN MH','N° SALIDAS']]#####Filtramos para no tener los demas parametros en el diccionario
    salidas_xmh_th=A4TH.groupby(['CODIGO TS SERVICIO','SENTIDO','TIPO_DIA','PERIODO_INICIO']).size().reset_index() #CONTAMOS LAS SALIDAS POR PERIODO DE INICIO
    salidas_xmh_th=salidas_xmh_th.rename(columns={0:'N° SALIDAS','PERIODO_INICIO':'MH','TIPO_DIA':'TIPO DIA'})

    comparacionSalidasXMH=pd.merge(salidas_xmh_a3,salidas_xmh_th,on=['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','MH'],suffixes=['_A3','_TH'],how='outer',indicator=True)
    salidas_sin_th_xmh=comparacionSalidasXMH[(comparacionSalidasXMH['_merge']=='left_only') & (comparacionSalidasXMH['N° SALIDAS_A3']>0)]

    salidas_sin_a3_xmh=comparacionSalidasXMH[comparacionSalidasXMH['_merge']=='right_only']
    comparacionSalidasXMH=comparacionSalidasXMH[comparacionSalidasXMH['_merge']=='both']
    
    comparacionSalidasXMH["Diferencia entre salidas"]=comparacionSalidasXMH.apply(lambda row: row['N° SALIDAS_A3'] - row['N° SALIDAS_TH'],axis=1)


    comparacionSalidasXMHA=comparacionSalidasXMH[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','AGRUPACIÓN MH','N° SALIDAS_A3','N° SALIDAS_TH']]
    comparacionSalidasXMHA=comparacionSalidasXMHA.groupby(['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','AGRUPACIÓN MH']).sum().reset_index()
    comparacionSalidasXMHA["Diferencia entre salidas"]=comparacionSalidasXMHA.apply(lambda row: row['N° SALIDAS_A3'] - row['N° SALIDAS_TH'],axis=1)
    
    Datos_dif_salidas_XMHA=comparacionSalidasXMHA[comparacionSalidasXMHA["Diferencia entre salidas"]!=0]
    Datos_dif_salidas_XMH=comparacionSalidasXMH[comparacionSalidasXMH["Diferencia entre salidas"]!=0]
    
    #Falta entregar un archivo con los sstd con  
    ##Moverse a salida
    os.chdir(ruta_OUTPUT)
    ##GUARDAR COMO EXCEL
    #Creamos planilla donde se guardará
    planilla_diferencias_salidas=pd.ExcelWriter('Diferencia Salidas a3 y a4.xlsx')
    ###Redundante
    dif_salidas_XMHA,dif_salidas_XMH,dif_salidas_sin_th_xmh,dif_salidas_sin_a3_xmh=Datos_dif_salidas_XMHA,Datos_dif_salidas_XMH,salidas_sin_th_xmh,salidas_sin_a3_xmh
    dif_salidas_XMH.to_excel(planilla_diferencias_salidas,sheet_name="Diferencia Salidas XMH")
    dif_salidas_sin_th_xmh.to_excel(planilla_diferencias_salidas,sheet_name="SSTD EN A3 sin TH")
    dif_salidas_sin_a3_xmh.to_excel(planilla_diferencias_salidas,sheet_name="SSTD EN TH sin salidas A3")
    dif_salidas_XMHA.to_excel(planilla_diferencias_salidas,sheet_name="Diferencia Salidas XMHA")
    
    planilla_diferencias_salidas.close()
    
    print('Proceso Finalizado')
    print('Revise archivo "Diferencia Salidas a3 y a4.xlsx", si están vacíos son anexos consistentes')
    print("Ubicado en:")
    print(ruta_OUTPUT)


def calcular_capacidades(A3_direccion,A4_direccion,ruta_OUTPUT):
    #Abriendo Anexos Nuevos
    A3=formatear_a3(A3_direccion)    
    # Parametros de operacion
    a4th=formatear_a4th(A4_direccion) 
    
     # Tabla Horaria
    a4th=a4th[(a4th['TIPO_EVENTO']=='C01') | (a4th['TIPO_EVENTO']=='C01 ')] ##fILTRAR EXPEDICIONES COMERCIALES
    ##Obtener Diccionario de capacidades de buses
    Dicc_cap_bus=pd.read_excel(A4_direccion,sheet_name="Diccionario",header=1)
    indices_salto=Dicc_cap_bus.loc[Dicc_cap_bus["Tipo Bus"].isnull()].index
    Dicc_cap_bus=Dicc_cap_bus.head(indices_salto[0]) ##Obtiene los parametros hasta el primer salto jeje
    
    
    
    
    agrupacionmhxsstpo=A3[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','MH','AGRUPACIÓN MH']]
    capacidad_xmh_a3=A3[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','MH','AGRUPACIÓN MH','CAPACIDAD (plazas)']]
    a4th=a4th.rename(columns={'TIPO_BUS':'Tipologia','TIPO_DIA':'TIPO DIA'})
    a4th=pd.merge(a4th,Dicc_cap_bus,on=['Tipologia'])
    
    
    
    capacidad_xmh_th=a4th.groupby(['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','PERIODO_INICIO'])['Capacidad'].sum().reset_index() #Sumamos capacidad
    capacidad_xmh_th=capacidad_xmh_th.rename(columns={'PERIODO_INICIO':'MH','TIPO_DIA':'TIPO DIA','Capacidad':'CAPACIDAD (plazas)'})
    
    comparacionCapacidadXMH=pd.merge(capacidad_xmh_a3,capacidad_xmh_th,on=['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','MH'],suffixes=['_A3','_TH'])
    comparacionCapacidadXMH["Diferencia entre Capacidad [Plazas]"]=comparacionCapacidadXMH.apply(lambda row: row['CAPACIDAD (plazas)_A3'] - row['CAPACIDAD (plazas)_TH'],axis=1)
    
    
    comparacionCapacidadXMHA=comparacionCapacidadXMH[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','AGRUPACIÓN MH','CAPACIDAD (plazas)_A3','CAPACIDAD (plazas)_TH']]
    comparacionCapacidadXMHA['CAPACIDAD (plazas)_TH']=comparacionCapacidadXMHA[['CAPACIDAD (plazas)_TH']].apply(pd.to_numeric)
    comparacionCapacidadXMHA=comparacionCapacidadXMHA.groupby(['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','AGRUPACIÓN MH'])['CAPACIDAD (plazas)_A3','CAPACIDAD (plazas)_TH'].sum().reset_index()
    
    comparacionCapacidadXMHA["Diferencia entre Capacidad [Plazas]"]=comparacionCapacidadXMHA.apply(lambda row: row['CAPACIDAD (plazas)_A3'] - row['CAPACIDAD (plazas)_TH'],axis=1)
    Datos_dif_capacidad_XMHA=comparacionCapacidadXMHA[comparacionCapacidadXMHA["Diferencia entre Capacidad [Plazas]"]!=0]
    Datos_dif_capacidad_XMH=comparacionCapacidadXMH[comparacionCapacidadXMH["Diferencia entre Capacidad [Plazas]"]!=0]
    #Falta entregar un archivo con los sstd con  
    ##Moverse a salida
    os.chdir(ruta_OUTPUT)
    ##GUARDAR COMO EXCEL
    #Creamos planilla donde se guardará
    planilla_diferencias_capacidad=pd.ExcelWriter('Diferencia Capacidad.xlsx')
    
    Datos_dif_capacidad_XMH.to_excel(planilla_diferencias_capacidad,sheet_name="Diferencia Capacidad XMH")
    Datos_dif_capacidad_XMHA.to_excel(planilla_diferencias_capacidad,sheet_name="Diferencia Capacidad XMHA")
    capacidad_xmh_th.to_excel(planilla_diferencias_capacidad,sheet_name="Cap TH")
    planilla_diferencias_capacidad.close()
    
    print('Proceso Finalizado')
    print('Revise archivo "Diferencia Capacidad.xlsx", si están vacíos son anexos consistentes')
    print("Ubicado en:")
    print(ruta_OUTPUT)

def revisar_dist(A3_direccion,A4_direccion,ruta_OUTPUT):
    
    #Abriendo Anexos Nuevos
    A3=formatear_a3(A3_direccion) 
    # Parametros de operacion
    A4TH=formatear_a4th(A4_direccion)

    ##fILTRAR EXPEDICIONES COMERCIALES
    A4TH=A4TH[(A4TH['TIPO_EVENTO']=='C01') | (A4TH['TIPO_EVENTO']=='C01 ')]
    ##Se renombra columnas para poder unirlas con anexo 3
    A4TH=A4TH.rename(columns={'PERIODO_INICIO':'MH','TIPO_DIA':'TIPO DIA','Capacidad':'CAPACIDAD (plazas)'})
    ##Union de ambos anexos
    A4THcdist=pd.merge(A3,A4TH,on=['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','MH'],suffixes=['_A3','_TH'],how='right')
    #Se calcula la resta de la distancia reportada en A3 con la del A4
    A4THcdist["Diferencia entre Distancia [Km]"]=A4THcdist.apply(lambda row: row['DISTANCIA TOTAL (POB+POI) (Km)'] - row['DISTANCIA'],axis=1)
    dif_dist=A4THcdist[A4THcdist["Diferencia entre Distancia [Km]"]!=0]
    dif_dist=dif_dist[['CODIGO TS SERVICIO','CODIGO USUARIO SERVICIO','SENTIDO','TIPO DIA','MH','HORA_INICIO','DISTANCIA TOTAL (POB+POI) (Km)','DISTANCIA','Diferencia entre Distancia [Km]']]
    dif_dist=dif_dist.rename(columns={'DISTANCIA TOTAL (POB+POI) (Km)':'DISTANCIA TOTAL A3','DISTANCIA':'DISTANCIA TOTAL A4'})
    
    os.chdir(ruta_OUTPUT)

    planilla_diferencias_dist=pd.ExcelWriter('Dif Dist A3 A4TH.xlsx')
    dif_dist.to_excel(planilla_diferencias_dist,sheet_name="Datos con dif dist")
    planilla_diferencias_dist.close()
    
    print('Proceso Finalizado')
    print('Revise archivo Dif Dist A3 A4TH.xlsx, si están vacíos son anexos consistentes')
    print("Ubicado en:")
    print(ruta_OUTPUT)


def revisar_intervalos(A1_direccion,A3_direccion,A4_direccion,ruta_OUTPUT):
    
    
    #Abriendo Anexos Nuevos
    A1=pd.read_excel(A1_direccion,sheet_name="Horarios",header=6)

    A1['HORARIO INICIO'] = A1.apply(lambda row: dt.datetime(1995,6,13,row['HORA INICIO'].hour,row['HORA INICIO'].minute,row['HORA INICIO'].second),axis=1)
    A1['HORARIO TERMINO'] = A1.apply(lambda row: convertir_horario_dia_siguiente(row['HORA INICIO'],row['HORA TERMINO']),axis=1)
    A1=A1.rename(columns={'Sentido':'SENTIDO'})
    #Hora ini mayor a 00 horas y menor a hora TERMINO -- Mover al dia siguiente

    A1=A1[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','HORARIO INICIO','HORARIO TERMINO']]
    A1['CODIGO TS SERVICIO']=A1.apply(lambda row: codts_tostring(row['CODIGO TS SERVICIO']),axis=1)
    A3=pd.read_excel(A3_direccion,sheet_name="Parámetros",header=6)  # Parametros de operacion
    A3['CODIGO TS SERVICIO']=A3.apply(lambda row: codts_tostring(row['CODIGO TS SERVICIO']),axis=1)
    DiccMHxMHA=A3[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','MH','AGRUPACIÓN MH']] ##Diccionario de agrupaciones de MHa x MH en A3
    a4th=pd.read_excel(A4_direccion,sheet_name="Tabla Horaria",header=6)
    a4th['CODIGO TS SERVICIO']=a4th.apply(lambda row: codts_tostring(row['CODIGO TS SERVICIO']),axis=1) #tabla horaria A4
    a4th=a4th.rename(columns={'PERIODO_INICIO':'MH','TIPO_DIA':'TIPO DIA'})
    a4tha1=pd.merge(a4th,A1,on=['CODIGO TS SERVICIO','SENTIDO','TIPO DIA'])
    ##Falta mover la hora de inicio al siguiente dia si es necesario..

    #convertir_horasalida_nextday(hora,minutos,segundos,horario_ini,horario_fin):

    #HORARIO DE INICIO DE EXPEDICION (IE) COMO FECHEHORA
    a4tha1['FH IE']=a4tha1.apply(lambda row: convertir_horasalida_nextday(row['HORA_INICIO'].hour,row['HORA_INICIO'].minute,row['HORA_INICIO'].second,row['HORARIO INICIO'],row['HORARIO TERMINO']),axis=1)
    a4tha1['DIA']=a4tha1.apply(lambda row: row['FH IE'].day,axis=1)

    a4thMHA=pd.merge(DiccMHxMHA,a4tha1,on=['CODIGO TS SERVICIO','SENTIDO','TIPO DIA','MH']) ##colocamos las MHA agrupadas que corresponde cada expedicion

    Todos_sstd=DiccMHxMHA[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA']]

    ##Todos los servicios sentidos sacando los duplicados para no hacer lo mismo infinitas veces jejeje
    Todos_sstd=Todos_sstd[~Todos_sstd.duplicated(['CODIGO TS SERVICIO','SENTIDO','TIPO DIA'])].reset_index()
    ##Obtenemos series para que sea mas simple el analisis, el programa lo hace como si una persona estuviera haciendolo en excel, asi es mas intuitivo (?)
    Servicios=Todos_sstd[~Todos_sstd.duplicated(['CODIGO TS SERVICIO'])]['CODIGO TS SERVICIO']
    Sentidos=Todos_sstd[~Todos_sstd.duplicated(['SENTIDO'])]['SENTIDO']
    Tipos_de_dias=Todos_sstd[~Todos_sstd.duplicated(['TIPO DIA'])]['TIPO DIA']

    Todas_mhs=DiccMHxMHA[['MH']].reset_index()
    mhs=Todas_mhs[~Todas_mhs.duplicated(['MH'])]['MH']




    res_list=[]
    for servicio in Servicios:
        for sentido in Sentidos:
            for td in Tipos_de_dias:
                ##Hasta acá ya se el sstd que tengo que ver
                ##Debo agarrar la tabla horaria solo de ese servicio
                th_sstd=a4thMHA[(a4thMHA['CODIGO TS SERVICIO']==servicio) & (a4thMHA['SENTIDO']==sentido) & (a4thMHA['TIPO DIA']==td)]

                if len(th_sstd.axes[0])==0:
                    pass
                else:
                    ###Hay que arreglar el timetable
                    th_sstd=th_sstd.sort_values(['DIA','FH IE']).reset_index(drop=True) #Ordenar por hora de salidas
                    ###Agarrar la agrupacion del sstd???--> es mas simple tomar el máximo de las agrupaciones, y construir una lista con la funcion range..
                    max_agrupacion=th_sstd['AGRUPACIÓN MH'].max()
                    if  max_agrupacion=='--' or np.isnan(max_agrupacion):
                        pass
                    else:
                    ###Iteramos sobre cada MHA
                        for PerAgr in range(1,max_agrupacion+1):  #Periodos Agrupados = PerAgr
                            # print([servicio,sentido,td,PerAgr])
                            
                            #Filtramos Tabla Horaria ahora en los serv agrupados
                            th_sstdMHA=th_sstd[th_sstd['AGRUPACIÓN MH']==PerAgr].reset_index(drop=True)
                            th_sstdMHA['Intervalo con expedicion anterior']= th_sstdMHA['FH IE'].diff()
                            intervalo_promedio=th_sstdMHA['Intervalo con expedicion anterior'].mean()
                            intervalo_minimo=th_sstdMHA['Intervalo con expedicion anterior'].min()
                            intervalo_maximo=th_sstdMHA['Intervalo con expedicion anterior'].max()
                            #falta revisar intervalo con intervalo...
                            for i,row in th_sstdMHA.iterrows():
                                intervalo_prog=row['Intervalo con expedicion anterior']
                                if pd.isnull(intervalo_prog):
                                    resultado='Inicio MHA'
                                else:
                                    resultado=(intervalo_aceptable(intervalo_minimo,intervalo_maximo,intervalo_prog))
                                #etiquetamos a la expedicion del SSTD
                                res_x_exp=pd.DataFrame({'CODIGO TS SERVICIO':[row['CODIGO TS SERVICIO']],'SENTIDO':[row['SENTIDO']],'TIPO DIA':[row['TIPO DIA']],'AGRUPACIÓN MH':[row['AGRUPACIÓN MH']],'Hora inicio exp':[row['FH IE']],'Cumplimiento Intervalo ManualPPO':[resultado],'Intervalo prom MHA':[intervalo_promedio]})
                                res_list.append(res_x_exp)

    resultados=pd.concat(res_list).reset_index()

    ###Esto agrupa primero y luego obtiene intervalo prom

    # ##Primero hay que obtener el intervalo promedio por MH
    # res_list2=[]
    # for servicio in Servicios:
    #     for sentido in Sentidos:
    #         for td in Tipos_de_dias:
    #             for mh in mhs:
    #             ##Hasta acá ya se el sstd que tengo que ver
    #             ##Debo agarrar la tabla horaria solo de ese servicio
    #                 th_sstdmh=a4thMHA[(a4thMHA['CODIGO TS SERVICIO']==servicio) & (a4thMHA['SENTIDO']==sentido) & (a4thMHA['TIPO DIA']==td) & (a4thMHA['MH']==mh)].reset_index(drop=True)
    #                 if len(th_sstdmh.axes[0])==0:
    #                     pass
    #                 else:
    #                     th_sstdmh['Intervalo con expedicion anterior']=th_sstdmh['FH IE'].diff()
                        
    #                     intervalo_promedio=th_sstdmh['Intervalo con expedicion anterior'].mean()
    #                     res_mh=pd.DataFrame({'CODIGO TS SERVICIO':[servicio],'SENTIDO':[sentido],'TIPO DIA':[td],'AGRUPACIÓN MH':[mh],'Iprom MH':[intervalo_promedio]})
    #                     res_list2.append(res_mh)
                        
    # IpromMH=pd.concat(res_list2).reset_index()                 


    ##Primero hay que obtener el intervalo promedio por MH
    res_list2=[]
    for servicio in Servicios:
        for sentido in Sentidos:
            for td in Tipos_de_dias:
            ##Hasta acá ya se el sstd que tengo que ver
            ##Debo agarrar la tabla horaria solo de ese servicio
                th_sstd=a4thMHA[(a4thMHA['CODIGO TS SERVICIO']==servicio) & (a4thMHA['SENTIDO']==sentido) & (a4thMHA['TIPO DIA']==td)].reset_index(drop=True)
                if len(th_sstd.axes[0])==0:
                    pass
                else:
                    th_sstd=th_sstd.sort_values(['DIA','FH IE']).reset_index(drop=True)
                    th_sstd['Intervalo con expedicion anterior']=th_sstd['FH IE'].diff() 
                    th_sstdmh=th_sstd[['MH','Intervalo con expedicion anterior']].groupby(['MH']).mean().reset_index()
                    th_sstdmh.rename(columns={'Intervalo con expedicion anterior':'Iprom'},inplace=True)
                    th_sstdmh['CODIGO TS SERVICIO']=servicio
                    th_sstdmh['SENTIDO']=sentido
                    th_sstdmh['TIPO DIA']=td
                    res_list2.append(th_sstdmh)
    IpromMH=pd.concat(res_list2).reset_index()

    IpromMH['Iprom_en_min']=IpromMH.apply(lambda row: row['Iprom'].seconds/60,axis=1)


    dif_intervalos=[]
    for servicio in Servicios:
        for sentido in Sentidos:
            for td in Tipos_de_dias:
                I_sstd=IpromMH[(IpromMH['CODIGO TS SERVICIO']==servicio) & (IpromMH['SENTIDO']==sentido) & (IpromMH['TIPO DIA']==td)].reset_index(drop=True)
                if len(I_sstd.axes[0])==0:
                    pass
                else:
                    # I_sstd=I_sstd.sort_values(['MH']).reset_index(drop=True)
                    I_sstd['Diferencia entre intervalos']=I_sstd['Iprom_en_min'].diff()
                    dif_intervalos.append(I_sstd)

    Idifprom=pd.concat(dif_intervalos).reset_index()
    ##Luego sacar el promedio en la MH
    ##Comparar ese con el de la siguiente MH, no puede subir bajar subir bajar en o bajar subir bajar subir (tener dientes)
    ##Revisar dientes





    ###ARCHIVO DE SALIDA
    ##Moverse a salida
    os.chdir(ruta_OUTPUT)
    ##GUARDAR COMO EXCEL
    #Creamos planilla donde se guardará

    planilla_intervalos=pd.ExcelWriter('Revision intervalos.xlsx')

    #agregamos a las hojas
    resultados.to_excel(planilla_intervalos,sheet_name="Expediciones XMHA",index=False)
    IpromMH.to_excel(planilla_intervalos,sheet_name="Prom Intervalo MH",index=False)
    Idifprom.to_excel(planilla_intervalos,sheet_name="Dif prom intervalos",index=False)
    #guardamos
    planilla_intervalos.close()

    print('Proceso Finalizado')
    print('Revise archivo "Revision intervalos.xlsx", si están vacíos son anexos consistentes')
    print("Ubicado en:")
    print(ruta_OUTPUT)


def intervalo_aceptable(Imin,Imax,Iprog):
    diferencia=Imax-Imin
    if diferencia<dt.timedelta(minutes=2) and Iprog<=dt.timedelta(minutes=10):
        return True
    elif diferencia<dt.timedelta(minutes=4) and (Iprog>dt.timedelta(minutes=10) and Iprog<=dt.timedelta(minutes=20)) :
        return True
    elif diferencia<dt.timedelta(minutes=6) and Iprog>dt.timedelta(minutes=20):
        return True
    else:
        return False
    

##Convertir celda en datetime y asignar al día de hoy o al siguiente 
##Como en los aviones
def convertir_horasalida_nextday(hora,minutos,segundos,horario_ini,horario_fin):
#Termina despues de las 0 horas y la expedicion empieza despues de las 00 horas y esas expediciones son antes de la hora de inicio si no fuera el contexto de un reloj
    if horario_fin.time()>=dt.time(hour=0) and dt.time(hora,minutos,segundos)>=dt.time(hour=0) and dt.time(hora,minutos,segundos)<horario_ini.time():
        return dt.datetime(1995,6,14,hora,minutos,segundos)
    else:
        return dt.datetime(1995,6,13,hora,minutos,segundos)
def convertir_horario_dia_siguiente(HI,HT):
    ##Requiere diferente tratamiento si es un datetime o un time...
    
    if HI.__class__.__name__=='time' and HT.__class__.__name__=='time':      
        if HI<HT:
            HT_ARREG=dt.datetime(1995,6,13,HT.hour,HT.minute,HT.second)
            return (HT_ARREG)
        elif HT<HI and dt.time(0,0)<=HT: #Condicion que implica que las expediciones siguientes pasan al otro dia
            HT_ARREG=dt.datetime(1995,6,14,HT.hour,HT.minute,HT.second)
            return (HT_ARREG)
    else:
        HI_mod=HI.time()
        HT_mod=HT.time()
        convertir_horario_dia_siguiente(HI_mod,HT_mod)
        
        
def calcular_a1_desde_a4(A4_direccion,ruta_OUTPUT):
    
    A4TH=formatear_a4th(A4_direccion)
    ##fILTRAR EXPEDICIONES COMERCIALES
    A4TH=A4TH[(A4TH['TIPO_EVENTO']=='C01') | (A4TH['TIPO_EVENTO']=='C01 ')]
    Todos_sstd=A4TH[['CODIGO TS SERVICIO','SENTIDO','TIPO DIA']]

    ##Todos los servicios sentidos sacando los duplicados para no hacer lo mismo infinitas veces jejeje
    Todos_sstd=Todos_sstd[~Todos_sstd.duplicated(['CODIGO TS SERVICIO','SENTIDO','TIPO DIA'])].reset_index()
    ##Obtenemos series para que sea mas simple el analisis, el programa lo hace como si una persona estuviera haciendolo en excel, asi es mas intuitivo (?)
    Servicios=Todos_sstd[~Todos_sstd.duplicated(['CODIGO TS SERVICIO'])]['CODIGO TS SERVICIO']
    Sentidos=Todos_sstd[~Todos_sstd.duplicated(['SENTIDO'])]['SENTIDO']
    Tipos_de_dias=Todos_sstd[~Todos_sstd.duplicated(['TIPO DIA'])]['TIPO DIA']

    
    
    
    return
