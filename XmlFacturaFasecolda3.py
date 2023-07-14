import os
import csv
import re
import http.client
import json
import base64
import urllib.parse
import fileinput
from zipfile import ZipFile
from tkinter import *
import tkinter as tk
#import PyPDF2
from PIL import Image, ImageTk
#from tkinter.filedialog import askopenfile
import tkinter.filedialog as filedialog
from datetime import datetime

root=tk.Tk()

# extracting all the files  
# funtion to Read text File

#obtiene la ruta+
def input_source(path1):
    input_path = tk.filedialog.askdirectory()
    #print(input_path)
    path1.insert(0, input_path)  # Insert the 'path'

def extraezip(file_path):
    # opening the zip file in READ mode
    with ZipFile(file_path, 'r') as zip:
	    #zip.extractall("C:/Users/jcarl/Proyectospython/OCR16012022/OCR23012022/Facturaszip/")
        zip.extractall(path)
        
# function to extract information

def hacerrx(file_path,archivo):
    
    x15=0
    x11=0
    x21=0
    z11=0
    x15=0
    fecha=''
    factura=''
    descripcion=''
    nit2=''
    valor=''

    for filename in fileinput.input([file_path],openhook=fileinput.hook_encoded("utf-8")):
        filename = filename.strip()
        #print(filename)
        #for lines in fileinput.input([bt]): esto solo sirve para ciertos documentos que no tienen utf-8
        #Hace el regex dentro del for para hacerlo por línea
        
        x9= re.findall("ParentDocumentID.\w+\d+", filename)
        str1=''
        x9a=str1.join(x9)

        x9b= re.findall("\w+\d+", x9a)
        x9c=str1.join(x9b)
        if x9c!='':
            #print(x9c)
            factura=x9c
            #print('la dificil',factura)

        # busqueda larga de due date <cbc:PaymentMeansCode>

        x3= re.findall("PaymentDueDate.\d{4}-\d{2}-\d{2}", filename)
        str1=''
        x3a=str1.join(x3)
        x3b = re.findall("\d{4}-\d{2}-\d{2}", x3a)
        x3c=str1.join(x3b)
        if x3c!='':
            #print(x3c)
            fecha=x3c

        x5= re.findall("PayableAmount currencyID=.COP..\d+", filename)
        str1=''
        x5a=str1.join(x5)
        x5b= re.findall("\d+", x5a)
        x5c=str1.join(x5b)
        if x5c!='':
            #print(x5c)
            valor=x5c

# Sacando la informacion para la descripcion

        x13= re.findall("<cac:Item>.*", filename)
        str2=''
        x13a=str2.join(x13)
        # si esta en un texto largo
        if x13a!='':
            if len(x13a)>10:
                x15=1
                #print('x15 primera si es 1')

        if x15==1:
            #print('filename cuando x15 es 1',filename)
            q = re.findall("<cbc:Description>.*</cbc:Description>", x13a)
            str=""
            q1=str.join(q)
            q2 = q1.replace("<cbc:Description>","",1)
            #print('q2 es',q2)
            sub_str = "</cbc:Description>"
            # slicing off after length computation
            res = q2[:q2.index(sub_str)]  
            #print('res es ',res)
            descripcion=res
            x15=0

        if x11==1:    
            qa = re.findall("<cbc:Description>.*</cbc:Description>", filename)
            str=""
            qa1=str.join(qa)
            #print('x11 segunda si es 1',filename)
            qa2 = qa1.replace("<cbc:Description>","",1)
            descripcion = qa2.replace("</cbc:Description>","",1)
            #print (descripcion)
            x11=0
        
        if x13a!='':
            if len(x13a)<=10:
                x11=1
                #print('Se encuentra item en linea')

        if z11==1:
            #print('filename cuando z11 para empresa es 1',filename)
            az = re.findall("\d+</cbc:CompanyID>", filename)
            str=""
            az1=str.join(az)
            #print('este es el nit2: ',az1)
            if az1!="":
                #bz = az1.replace("</cbc:CompanyID>","",1)
                nit2 = az1.replace("</cbc:CompanyID>","",1)
                z11=0

        z13= re.findall("<cac:AccountingSupplierParty>", filename)
        str2=''
        z13a=str2.join(z13)

        if z13a!='':
            z11=1
            #print('si encuentra supplier', z13a)

    listaind=[factura,fecha, valor, nit2, descripcion,archivo]
    date_obj = datetime.strptime(fecha, '%Y-%m-%d')
    fechacorrecta = date_obj.strftime('%d/%m/%Y')


    create_form(token,factura,valor,nit2,fechacorrecta,descripcion)
    print ("La lista es:", listaind)
    with open(r'C:\Users\David.Rodriguez\Desktop\FACTURAS PY\datos.csv', mode='a', newline='') as datosentra:
        emq = csv.writer(datosentra)
        emq.writerow(listaind)
        datosentra.close()
             

def sacainfo():
    
    # crea el archivo CSV, pone los titulos
    with open(r'C:\Users\David.Rodriguez\Desktop\FACTURAS PY\datos.csv', mode='w', newline='') as datosentra:
        emq = csv.writer(datosentra, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        emq.writerow(['No Factura', 'Valor a Pagar', 'Fecha de Pago', 'NIT Proveedor','Descripcion','Archivo de Factura'])

    #extrae info de zip
    files = os.listdir(path1.get())
    for file in files:
#   for file in os.listdir():
        # Check whether file is in text format or not
        if file.endswith(".zip"):
            file_path = f"{path}\{file}"
            # call read text file function
            extraezip(file_path)
    archivo=''
    for file in files:
        # Check whether file is in text format or not
        if file.endswith(".xml"):
            file_path = f"{path}\{file}"
            # call read text file function
            hacerrx(file_path,file)
            #filename = Path.GetFileName(file);
        archivo=archivo+' , '+file

           
    #text box in root window
    text1 = Text(root)
    text1.grid( columns=4, rows=10)
    #text input area at index 1 in text window
    text1.insert('1.0', archivo)
    text1.insert('1.0','''Los archivos que han sido gestionados son>\:''')

# global variables
environment = 'sa1'

# function to get access token
def get_auth_token():
    userName = '82d4c35a-320a-47e5-8661-80b9de2eafe8'
    password = '+vyQTfM32ULSwLmnn6KytGDCmCtdzhP6WiJhrNY8xok='

    postData = {
        'username': userName,
        'password': password,
        'grant_type': 'password'
    }
    postDataEncoded = urllib.parse.urlencode(postData)

    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Host': f'{environment}.visualvault.com',
        'Authorization': 'Basic ' + base64.b64encode((userName + ':' + password).encode('utf-8')).decode('utf-8')
    }

    conn = http.client.HTTPSConnection(f'{environment}.visualvault.com')
    conn.request('POST', '/oauth/token', postDataEncoded, headers)
    response = conn.getresponse()
    data = response.read().decode('utf-8')
    conn.close()

    token = json.loads(data)['access_token']
    return token


# function to create a form instance and fill out the form fields
# listaind=[factura,fecha, valor, nit2, descripcion,archivo]

def create_form(token,fact,val,nit,fechac,descr):
    postData = {
        'txt_idfactura': fact,
        'txt_numerodocumento': fact,
        'txt_numvalorf': val,
        'Nit proveedor': nit,
        'Fecha del documento' : fechac,
        'Observaciones': descr,
        'ddl_tipo de documetno':"Factura",
        'Tipo de radicación':"Electrónica",
        'Estado General':"cargue realizado",        
    }
    postDataEncoded = json.dumps(postData)

    customeralias = 'Fasecolda'
    databasealias = 'Default'
    #formTemplateId = '96eb77f0-fee9-ed11-b988-0eb5b24fbfab'
    #formTemplateId = '8c2f85e8-8b0c-ee11-b994-0eb5b24fbfab'
    #formTemplateId = '2fd72b2e-880d-ee11-b995-0eb5b24fbfab'
    formTemplateId = '2e14155a-5111-ee11-b995-0eb5b24fbfab'

    #Número del Documento = txt_idfactura
    #Número del Documento = txt_numerodocumento
    #Fecha de Vigencia = Fecha del documento
    #Valor = txt_facturavalor
    #descripción =  Observaciones
    #nombre de la empresa = Seleccionar proveedor
    #nombre del documento de la factura ?
    #Formu = FS-Correspondencia
    #idform= 2e14155a-5111-ee11-b995-0eb5b24fbfab

    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {token}'
    }

    conn = http.client.HTTPSConnection(f'{environment}.visualvault.com')
    conn.request('POST', f'/api/v1/{customeralias}/{databasealias}/formtemplates/{formTemplateId}/forms', postDataEncoded, headers)
    response = conn.getresponse()
    data = response.read().decode('utf-8')
    conn.close()

    print(data)


#declara variables temporales
# 
x11=0
x21=0
z11=0
x15=0
descripcion=''    

#Inicio Programa pone el logo

token = get_auth_token()

logo = Image.open('logogrm.jpg')

logo = ImageTk.PhotoImage(logo)
logo_label = tk.Label(image=logo)

logo_label.image = logo
logo_label.grid(column=1,row=0)


#tratando de sacar la ruta de un texto que se escriba
pt=Label(root, text='Ruta de los archivos de las facturas: ')
pt.grid(column=1,row=3)
path1 = tk.Entry(root, text="", width=40)
path1.grid(column=3,row=3)
#path = "C:/Users/jcarl/Proyectospython/OCR16012022/OCR23012022/Facturaszip/"
browse1 = tk.Button(root, text="Browse", command=input_source(path1))
browse1.grid(column=3,row=2)
path=(path1.get())
print(path)

browse_btn= tk.Button(root, text="Extraer Información", command=lambda:sacainfo(), font="Raleway", bg="#C70039" , fg="white", height=2, width=15)
browse_btn.grid(column=1, row=2)

#Label(root, text='Palabra que busca: ').grid(column=1,row=4)
#e2 = Entry(root)
#e2.grid(row=4, column=2)
#os.chdir(path)

# call get_auth_token function and chain the create_form function using the access token returned from get_auth_token
#token = get_auth_token()
#create_form(token)

root.mainloop()
