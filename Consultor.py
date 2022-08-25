from cgitb import text
import inspect, urllib3, paramiko, sys, io, random, xlwt, pyodbc, os, shutil, tkinter, pyperclip, time, openpyxl, webbrowser, smtplib, ssl, subprocess, csv, xlrd, math, PIL, pytesseract, comtypes.client, threading, time, pyperclip, sys, openpyxl, tkinter, mysql.connector, pyodbc
from tkinter.tix import CheckList
from typing import Any
from pickle import TRUE
from tabnanny import check
from multiprocessing.sharedctypes import Value
from turtle import goto, onclick
from win32 import win32gui, win32process, win32api
from win32.lib import win32con
from tkinter.filedialog import askopenfilename
import pyscreenshot as img
import cv2 as cv
from tkinter import Button, ttk, messagebox
from openpyxl.styles import colors
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.cell import Cell
from datetime import date, datetime, timedelta
from datetime import timedelta
from fpdf import FPDF
from xlutils.copy import copy
from win32com.client import Dispatch
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
import fitz, warnings
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from email.mime.multipart import MIMEMultipart#
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from PIL import Image
from time import sleep
import numpy as np
import win32com
from win32com.client import DispatchEx,Dispatch
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
import exchangelib, requests
from collections import Counter as Contador
import locale,stat,decimal
from libreria import *
from datetime import datetime as dt, timedelta as td
def Cerrar(): 
    root.destroy()

class ModalWindow(object):
    global but
    def __init__(self, root, Imagen):
        self.entry_id = tkinter.StringVar()
        self.modalWindow = tkinter.Toplevel(root)
        self.modalWindow.title('Captcha')
        self.modalWindow.resizable(False, False)
        self.modalWindow.wm_attributes("-topmost", True)
        self.modalWindow.geometry('330x100+'+str(int((win32api.GetSystemMetrics(0)-420)/2))+'+'+str(int((win32api.GetSystemMetrics(1)-140)/2)))
        self.modalWindow.overrideredirect(1)
        Imagg = tkinter.PhotoImage(file=os.getcwd() + "\\" + Imagen)
        label = tkinter.ttk.Label(self.modalWindow, image=Imagg)
        label.image = Imagg
        label.place(x=20, y=10)
        labeled = tkinter.ttk.Label(self.modalWindow, text='Ingrese Captcha')
        labeled.place(x=180, y=10)
        but = tkinter.ttk.Entry(self.modalWindow, textvariable=self.entry_id)
        but.place(x=180, y=30)
        Aceptar = tkinter.ttk.Button(self.modalWindow, text='Aceptar', command=self.Adios)
        Aceptar.place(x=180, y=60)
        self.modalWindow.rowconfigure(1, minsize = 500)

    def Adios(self):
        self.modalWindow.destroy()

    def GetCaptcha(self):
        text = self.entry_id.get()
        return str(text)

# def colnum_string(n):

# def Void(folder):
  
# def DiccionarioSQL(Select, Modo=2, OnHeader=False):

# def Execute(Update, Modo=2):

# def DiccionarioExcelXLSX(Excel, Filass=1):

# def DiccionarioExcelXLS(Excel):

# def ResumenPrevio(Lista, Nombre):    

# def ResumenPrevio_2(Lista, Nombre):

# def ResumenPrevio_3(Lista, Nombre):

# def TablaHTML(Lista):

# def Consolidado(Listado):

# def GuiaMax():

# def CreaPDF(x):

# def UnirPDFs(Listado):

# def InicioPortal():

# def DoRecibos(x):

# def Buscar_Columna(a):

# def Cambio(a):

# def SMTPLib(Subject, Recipients, Body, attachments=None):

# def CorreoFaltantes(Faltantes=[]):
    
# def Correo_No_Despacho():

# def make_excel(resultado,dst,name):# Mejorar funcion

# def CorreoDespacho():


# def carga_input():


# def Activate():

# #Validaciones
# def AgruparTODO(Todo):


# def Actualizar():


# def RutaSSH(sftp, Ruta):
#     Ruteando = [a for a in Ruta.split("/") if a!='']
#     for idx, x in enumerate(Ruteando):
#         try: sftp.chdir("".join(["/{0}".format(str(b)) for b in Ruteando[:idx+1]]))
#         except IOError:
#             sftp.mkdir("".join(["/{0}".format(str(b)) for b in Ruteando[:idx+1]]))
#             sftp.chdir("".join(["/{0}".format(str(b)) for b in Ruteando[:idx+1]]))

# def SubirSSH():
#     def Reconectar(ssh,Ruta):
#         ssh.close()
#         ssh.connect('172.28.11.228', username="usr_ychavez", password="Password$31",port=22, timeout=100)
#         sftp = ssh.open_sftp()
#         sftp.chdir(Ruta)
#         return sftp
#     def put(ssh,sftp,src,dst,times=8):
#         time=1
#         while True:
#             if time==2**times:break
#             try:
#                 sftp.put(src,dst)
#                 break
#             except Exception as e:
#                 if not(ssh.get_transport().is_active()) or sftp.get_channel().closed:
#                     try:
#                         sftp=Reconectar(ssh,dst)
#                     except Exception as e:
#                         print("Error1:",e)
#                         print(f"se espera {time} segundos ....")
#                         sleep(time)
#                         time*=2
#                     continue
#                 print("Error2:",e)
#                 print(f"se espera {time} segundos ....")
#                 sleep(time)
#                 time*=2
                
#     ssh = paramiko.SSHClient()
#     ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
#     User = DiccionarioSQL("SELECT APLICATIVO, USUARIO, CLAVE FROM robot_accesibilidad WHERE Aplicativo = 'SSH'", 4)[0]
#     print(User)
#     #ssh.connect('172.28.11.228', username=User['USUARIO'], password=User['CLAVE'])
# ##    ssh.connect('172.28.11.228', username="lurcuhua", password="Brasil54+",port=22, timeout=100)
#     ssh.connect('172.28.11.228', username="usr_ychavez", password="Password$28",port=22, timeout=100)
# ##    ssh.connect('172.28.11.228', username="ssh_jsantiagoa", password="Jhon%123456789023")
#     sftp = ssh.open_sftp()
#     locale.setlocale(locale.LC_ALL, 'es-ES')
#     mes= date.today().strftime("%m")+"."+date.today().strftime("%B").upper()
#     locale.setlocale(locale.LC_ALL, 'en_US')    
#     print(mes)

#     for x in ['Fisico', 'Digital']:
        
#         if not os.path.isdir(os.getcwd() + "//Resultado//{0}".format(x)): continue
#         Ruta = '/SSH/Despacho UE/DESPACHO_ROBOT/{0}/{1}/{2}'.format(mes,x,date.today().strftime("%d.%m.%y"))
#         RutaSSH(sftp, Ruta)
#         Excel = '{0} {1}.xls'.format(x, date.today().strftime("%d.%m.%Y"))
#         put(ssh,sftp,os.getcwd() + "\\Resultado\\{0}".format(Excel),"{0}/{1}".format(Ruta, Excel))
# ##        sftp.put(os.getcwd() + "\\Resultado\\{0}".format(Excel), "{0}/{1}".format(Ruta, Excel)) #Mover Excel
#         Lista = os.listdir(os.getcwd() + "/Resultado/{0}".format(x))
#         pb.configure(maximum=len(Lista))
#         progress.set(1)
#         Carpeta = '{0} {1}'.format(x, date.today().strftime("%d.%m.%Y"))
#         Carpeta = '{0}'.format(x)
        
#         for idy, y in enumerate(Lista):
#             progress.set(idy+1)
#             Cuenta.config(text=str(idy+1)+"/"+str(len(Lista)))
#             root.wm_attributes("-topmost", True)
#             root.update_idletasks()
#             root.update()
#             if idy==0: RutaSSH(sftp, Ruta + "/" + Carpeta);
#             print (Ruta)
#             print (Carpeta)
#             print (y)
#             put(ssh,sftp,os.getcwd() + "/Resultado/{0}/{1}".format(x, y), "{0}/{1}/{2}".format(Ruta, Carpeta, y))
# ##            sftp.put(os.getcwd() + "/Resultado/{0}/{1}".format(x, y), "{0}/{1}/{2}".format(Ruta, Carpeta, y)) #Mover Archivos
#     sftp.close()
#     ssh.close()
#     messagebox.showinfo("Mensaje", "Terminó")

# def AnularValores():


# def LecturaPDF():

def ShowP():
    a=[CierreP, ReversionesP, AseguramientoP, RegistroP]
    for x in a:
        print(str(x))
        myLabel=tkinter.Label(root, text=x.get()).pack()
    # myLabel=tkinter.Label(root, text=ReversionesP.get()).pack()
    # myLabel=tkinter.Label(root, text=CierreP.get()).pack()
    # myLabel=tkinter.Label(root, text=RegistroP.get()).pack()
    if Pendiente.get()==1:
        print('PONEMOS LA BASE DE DATOS')


def BasesP():
    global CierreP, ReversionesP, AseguramientoP, RegistroP

    CierreP= tkinter.IntVar()
    Marca1 = tkinter.Checkbutton (root, text='Pendientes de Cierre', variable=CierreP)
    Marca1.place(x=130, y=40)

    ReversionesP= tkinter.IntVar()
    Marca2 = tkinter.Checkbutton (root, text='Pendientes de Reversiones', variable=ReversionesP)
    Marca2.place(x=130, y=70)

    AseguramientoP= tkinter.IntVar()
    Marca3 = tkinter.Checkbutton (root, text='Pendientes de Aseguramiento', variable=AseguramientoP)
    Marca3.place(x=130, y=100)

    RegistroP= tkinter.IntVar()
    Marca4 = tkinter.Checkbutton (root, text='Pendientes de Registro Web Publica', variable=RegistroP)
    Marca4.place(x=130, y=130)

    myButton=Button(root, text='Show selection', command=ShowP).pack()

def Pendientes():
    global Pendiente
    Pendiente= tkinter.IntVar(root,value=1)
    # Pendiente= tkinter.IntVar()
    Pe = tkinter.Checkbutton(root, text='PENDIENTES', variable=Pendiente)
    
    # Pe = tkinter.Checkbutton (root, text='PENDIENTES', variable=Pendiente)
    # Pe.place(x=130, y=5)
    # Pe.select()
    Pe.pack()

    # myLabel=tkinter.Label(root, text=Pendiente.get()).place(x=250, y=250)
    BasesP()






if __name__ == '__main__':
#     NoExcel = []
#     Robot_Despacho = """
#         SELECT ID,NEGOCIO,NUMERO_RESOLUCION,FECHA_RESOLUCION,COD_RECLAMO,SERVICIO,NOMBRE,
#         ANALISTA,FECHA_DESPACHO,DIRECCION,DISTRITO,PROVINCIA,DEPARTAMENTO,ANEXO_INSCRIPCION,CLIENTE,SEGMENTO,FECHA_RECLAMO,
#         MONTO_RECLAMO,TIPO_ENVIO,RESULTADO,CORREO,CANAL,INSTANCIA,CODIGO_RECLAMO_1RA,FECHA_CARGA,GUIA,NOMBRE_ARCHIVO
#         FROM robot_despacho
#         WHERE GUIA IS NULL AND NOMBRE_ARCHIVO IS NULL
#         ORDER BY RESULTADO ASC
#     """
#     Robot_No_Despachado = "SELECT NEGOCIO,ANALISTA,NOMBRE_EXCEL,NUMERO_RESOLUCION,COD_RECLAMO,FECHA_RECLAMO,NOMBRE_ARCHIVO AS OBSERVACION FROM ROBOT_DESPACHO WHERE GUIA IS NULL AND FECHA_CARGA=CONVERT(VARCHAR,GETDATE(),103)"

#     locale.setlocale(locale.LC_ALL, 'es-ES')
#     mes= date.today().strftime("%m")+"_"+date.today().strftime("%B").upper()+"_"+date.today().strftime("%Y")
# ##    Ruta_Despacho = "\\\\10.4.40.191\\Informacion Primera y Segunda Instancia Soluciones\\DESPACHO_ROBOT\\Input\\{1}\\{0}".format(date.today().strftime("%Y-%m-%d"),mes)
#     Ruta_Despacho = os.getcwd()+"\\Input\\{1}\\{0}".format(date.today().strftime("%Y-%m-%d"),mes)
# ##def prueba():
#     if not(os.path.isdir(os.getcwd()+"\\Input\\{0}".format(mes))):os.mkdir(os.getcwd()+"\\Input\\{0}".format(mes))
#     if not(os.path.isdir(Ruta_Despacho)):os.mkdir(Ruta_Despacho)
#     #Ruta_Despacho ="D:\\2020-08-26"

#     locale.setlocale(locale.LC_ALL, 'en_US')
    
#     print(Ruta_Despacho)


#     Ruta_Pdfs = Ruta_Despacho
#     Ruta_descargas = os.getcwd()+ "\\Descargas"
#     Ruta_Eliminados = os.getcwd()+ "\\Eliminados"
#     Change = {"NUMERO_RESOLUCION":"NUMERO_RESOLUCIONssssssssssssssss","FECHA_RESOLUCION":"FECHA_RESOLUCION","CODIGO_RECLAMO":"COD_RECLAMO","SERVICIO":"SERVICIO","NOMBRE DEL RECLAMANTE":"NOMBRE","ANALISTA":"ANALISTA","DIRECCION":"DIRECCION","DISTRITO":"DISTRITO","PROVINCIA":"PROVINCIA","DEPARTAMENTO":"DEPARTAMENTO","ANEXO/INSCRIPCION":"ANEXO_INSCRIPCION","FECHA_RECLAMO":"FECHA_RECLAMO","MONTO_RECLAMO":"MONTO_RECLAMO","TIPO_ENVIO":"TIPO_ENVIO","RESULTADO":"RESULTADO","CORREO":"CORREO","CANAL":"CANAL","INSTANCIA":"INSTANCIA","NEGOCIO":"NEGOCIO","CODIGODERECLAMO1RAINSTANCIA":"CODIGO_RECLAMO_1RA"}
#     Variables = ['GUIA','NUMERO_RESOLUCION','FECHA_RESOLUCION','COD_RECLAMO','SERVICIO','NOMBRE','ANALISTA','FECHA_DESPACHO','DIRECCION','DISTRITO','PROVINCIA','DEPARTAMENTO','ANEXO_INSCRIPCION','SEGMENTO','FECHA_RECLAMO','CANAL','INSTANCIA','CORREO']
#     Diccionario_columnas = DiccionarioSQL("select  Columna_variable,Columna_correcta from sa_diccionario_columnas  ", 4)
    
    #Tkinter
    
    warnings.filterwarnings("ignore")
    root = tkinter.Tk()
    root.geometry('1200x600+'+str(0)+'+'+str(win32api.GetSystemMetrics(1)-180))
    root.title('CONSULTOR')
    root.wm_attributes("-topmost", True)
    #root.overrideredirect(1)
    root.resizable(False, False)
    print('Inicia los graficos')
    importar = tkinter.Button(root, text='PENDIENTES',command=Pendientes, height = 2, width = 15,bg = "blue", fg = "white")
    importar.place(x=10, y=10)

    Casilla = tkinter.Label(root, text="Ingrese Data requerida")
    Casilla.place(x=300, y=5)

    Input = tkinter.Entry(width=80)
    Input.place(x=130, y=25)

    Recibos = tkinter.Button(root, text='BASE PROCESADOS',  height = 2, width = 15,bg = "green", fg = "white")
    Recibos.place(x=10, y=60)

    ActRespuesta = tkinter.Button(root, text='ESTADO DE LA BASE',  height = 2, width = 15,bg = "orange", fg = "white")
    ActRespuesta.place(x=10, y=110)

    Faltan = tkinter.Button(root, text='Casos no Despachados',  height = 2, width = 15,bg = "brown", fg = "white")
    Faltan.place(x=10, y=160)

    Despacha = tkinter.Button(root, text='Correo Despacho',  height = 2, width = 15,bg = "gray", fg = "white")
    Despacha.place(x=10, y=210)

    SSH = tkinter.Button(root, text='Subir SSH',  height = 2, width = 15,bg = "black", fg = "white")
    SSH.place(x=10, y=260)

    Anular = tkinter.Button(root, text='Anular',  height = 2, width = 15,bg = "turquoise", fg = "white")
    Anular.place(x=10, y=310)

    Arreglar = tkinter.Button(root, text='Arreglar Recibos',  height = 2, width = 15,bg = "purple", fg = "white")
    Arreglar.place(x=10, y=360)

    descarga = tkinter.Button(root, text='Descarga ssh',  height = 2, width = 15,bg = "yellow", fg = "red")
    descarga.place(x=10, y=410)


    cerrar = tkinter.Button(root, text='Cerrar',  height = 2, width = 15,bg = "red", fg = "white")
    cerrar.place(x=10, y=460)

    Cuenta = tkinter.Label(root, text="??? / ???")
    Cuenta.place(x=160, y=550)

    progress = tkinter.IntVar()
    pb = ttk.Progressbar(root, orient="horizontal", length=120, mode="determinate", variable=progress)
    pb.place(x=10, y=550)
    root.mainloop()
