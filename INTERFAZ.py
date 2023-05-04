# -*- coding: utf-8 -*-
"""
Created on Wed Apr  5 12:24:59 2023

@author: rafael.cozmar
"""

import tkinter as tk
from tkinter import filedialog
import os
from core import param_faltantes_A3,param_faltantes_A4,revisar_horarios_a1_a4,contar_salidas,calcular_capacidades,revisar_dist,revisar_intervalos

root=tk.Tk()
root.title('Panel PO US')

# fondo=ImageTk.PhotoImage(file = 'citaro.jpg')
# label_fondo=tk.Label(root,image=fondo)
# label_fondo.place(x=0,y=0)
root.geometry("1200x300")
titulo = tk.Label(root,text= 'Panel PO US Beta',font='Arial 30',bg='red')
output_text=tk.Label(root,text = '')
NOMBRE_OUTPUT=tk.Entry(root, font = 'Arial 12')

rout=os.getcwd()

def ejecutar_param_faltantes_a3():
    param_faltantes_A3(rA3,rout)
    
def ejecutar_param_faltantes_a4():
    param_faltantes_A4(rA4,rout)

def ejecutar_contar_salidas():
    contar_salidas(rA3,rA4,rout)

def ejecutar_revisar_horarios():
    revisar_horarios_a1_a4(rA1,rA4,rout)
    
def ejecutar_revisar_cap():
    calcular_capacidades(rA3,rA4,rout)

def ejecutar_revisar_dist():
    revisar_dist(rA3,rA4,rout)

def ejecutar_revisar_intervalos():
    revisar_intervalos(rA1,rA3,rA4,rout)
    
def obtener_ruta_a1():
    global rA1
    rA1=filedialog.askopenfilename()
    lbl_a1.config(text=str(rA1),bg='white')
def obtener_ruta_a3():
    global rA3
    rA3=filedialog.askopenfilename()
    lbl_a3.config(text=str(rA3),bg='white')
def obtener_ruta_a4():
    global rA4
    rA4=filedialog.askopenfilename()
    lbl_a4.config(text=str(rA4),bg='white')
def obtener_ruta_out():
    global rout
    rout=filedialog.askdirectory()
    lbl_rout.config(text=str(rout),bg='white')

# def guardar_directorio():
    


###Botones para cargar anexos##################
btna1 = tk.Button(root, text = 'Cargar Anexo 1',command = obtener_ruta_a1)
btna1.grid(row=1,column=1)
btna3 = tk.Button(root, text = 'Cargar Anexo 3',command = obtener_ruta_a3)
btna3.grid(row=3,column=1)
btna4 = tk.Button(root, text = 'Cargar Anexo 4',command = obtener_ruta_a4)
btna4.grid(row=4,column=1)
###
##Boton ruta output
btna5 = tk.Button(root, text = 'Seleccionar Carpeta Salidas',command = obtener_ruta_out)
btna5.grid(row=5,column=1)


lbl_a1=tk.Label(root,text='No cargado',bg='red',anchor='w',font='Arial 8')
lbl_a1.grid(row=1,column=2)
lbl_a3=tk.Label(root,text='No cargado',bg='red',anchor='w',font='Arial 8')
lbl_a3.grid(row=3,column=2)
lbl_a4=tk.Label(root,text='No cargado',bg='red',anchor='w',font='Arial 8')
lbl_a4.grid(row=4,column=2)



##Label ruta output
lbl_rout=tk.Label(root,text=str(rout),bg='skyblue',anchor='w',font='Arial 8')
lbl_rout.grid(row=5,column=2)

###############################################


######Botones de acciones a ejecutar#####
boton_param_faltantes_a3=tk.Button(root,text='Buscar Parámetros faltantes A3',command = ejecutar_param_faltantes_a3)
boton_param_faltantes_a4=tk.Button(root,text='Buscar Parámetros faltantes Tabla Horaria A4',command = ejecutar_param_faltantes_a4)
boton_revisar_horarios=tk.Button(root,text='Revisar horarios de Tabla Horaria y Anexo 1',command = ejecutar_revisar_horarios)
comparar_salidas_a3a4=tk.Button(root,text='Comparar salidas entre Anexo 3 y Anexo 4 (Tabla horaria)',command = ejecutar_contar_salidas)
comparar_capacidades_a3a4=tk.Button(root,text='Comparar capacidad entre Anexo 3 y Anexo 4 (Tabla horaria y Diccionario)',command = ejecutar_revisar_cap)
boton_revisar_intervalos=tk.Button(root,text='Revisar intervalos Tabla Horaria',command= ejecutar_revisar_intervalos) 
boton_revisar_dist=tk.Button(root,text='Revisar Distancias A3 y A4',command= ejecutar_revisar_dist) 
titulo.grid(row=0,column=0)
boton_param_faltantes_a3.grid(row=1,column=0)
boton_param_faltantes_a4.grid(row=2,column=0)
boton_revisar_horarios.grid(row=3,column=0)
comparar_salidas_a3a4.grid(row=4,column=0)
comparar_capacidades_a3a4.grid(row=5,column=0)
boton_revisar_intervalos.grid(row=6,column=0)
boton_revisar_dist.grid(row=7,column=0)

# NOMBRE_OUTPUT.grid(row=1,column=1)
root.mainloop()