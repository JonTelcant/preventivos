from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from difflib import SequenceMatcher as IndiceCoincidencia
import sys
import re
import csv
import msvcrt

def cargar_excel(nombre):
    try:
        return load_workbook(nombre)
    except Exception:
        print(f"No se ha encontrado el fichero {nombre}")
        sys.exit(1)

def FormatearEmplazamiento(emplazamiento):


	emplazamiento_fomateado=re.sub('(2G)','',emplazamiento) # Borro 2G

	emplazamiento_fomateado=re.sub('(3G)','',emplazamiento_fomateado) # Borro 3G

	emplazamiento_fomateado=re.sub('(4G)','',emplazamiento_fomateado) # Borro 4G

	emplazamiento_fomateado=re.sub('(C.T.)','',emplazamiento_fomateado)

	emplazamiento_fomateado=re.sub('(CT)','',emplazamiento_fomateado)

	emplazamiento_fomateado=re.sub('(REP-INT-GSM)','',emplazamiento_fomateado) # Borro TSM

	emplazamiento_fomateado=re.sub('(--)','',emplazamiento_fomateado) # Borro --

	emplazamiento_fomateado=re.sub('(ATW-T)','',emplazamiento_fomateado) # Borro ATW-T

	emplazamiento_fomateado=re.sub('(ATW-V)','',emplazamiento_fomateado)

	emplazamiento_fomateado=re.sub('(ATW)','',emplazamiento_fomateado)

	emplazamiento_fomateado=re.sub('[-]','',emplazamiento_fomateado) # Borro guiones

	emplazamiento_fomateado=re.sub('(TSM)','',emplazamiento_fomateado) # Borro TSM

	emplazamiento_fomateado=re.sub('(T.S.M.)','',emplazamiento_fomateado)

	emplazamiento_fomateado=re.sub('(E.B.)','',emplazamiento_fomateado)

	#emplazamiento_fomateado=re.sub('(EB)','',emplazamiento_fomateado)

	emplazamiento_fomateado=re.sub('\s?\d', '',emplazamiento_fomateado) # Borro numeros y espacion en blanco inecesarios

	emplazamiento_fomateado=emplazamiento_fomateado.lstrip()

	return(emplazamiento_fomateado)

def LeerExcelEmplazamientos(archivo: Workbook):
    hojaExcel = archivo.worksheets[0]
    lista = []
    for fila in hojaExcel.rows:
        emplazamiento = fila[4].value
        emplazamientoFormateado = FormatearEmplazamiento(emplazamiento)
        latitud = fila[2].value
        altitud = fila[3].value
        lista.append([emplazamiento, emplazamientoFormateado, latitud, altitud])
    return lista

def LeerExcelPreventivos(archivo: Workbook):
     
    hojaExcel = archivo.worksheets[0]
    lista = []
    cont = 0
    emplazamientoA = "SALA ELEMENTO"
    accion = ""
    cantidad = 0
    for fila in hojaExcel.rows:
        cont += 1
        Depurar(emplazamientoA)
        Depurar(cont)
        if emplazamientoA != fila[5].value:
            if cont != 2: 
                lista.append ([emplazamientoA, emplazamientoFormateadoA, emplazamientoB, \
                                         emplazamientoFormateadoB, accion])
                cantidad += 1
            emplazamientoA = fila[5].value
            emplazamientoB = fila[6].value
            emplazamientoFormateadoA = FormatearEmplazamiento(emplazamientoA)
            emplazamientoFormateadoB = FormatearEmplazamiento(emplazamientoB)
            accion = str(fila[8].value)
        else:
            if cont != 1: accion = accion + "\n" + str(fila[8].value)
        emplazamientoA = fila[5].value
    lista.append ([emplazamientoA, emplazamientoFormateadoA, emplazamientoB, \
                                         emplazamientoFormateadoB, accion])
    print(emplazamientoA)
    
    cantidad += 1
    print("Se han encontrado " + str(cantidad) + " preventivos")
    return lista

def ImprimirLista(listaPreventivo, emplazamientoCoincidencia):
    print("El preventivo " + listaPreventivo[0][4] + " ha tenido varias coincidencias: ")
    cont=0
    for i in emplazamientoCoincidencia:
        cont=cont+1
        print(str(cont) + ": " + i[0])
    
    while True:
        try:
            opcion=int(input("Elige una opcion: "))
            if opcion <= cont and opcion>0:
                return listaPreventivo[opcion-1]
                break
            else:
                print("La opcion elegida es incorrecta...")
        except:
             print("La opcion elegida es incorrecta...")  

def BuscarCoincidencia(listaPreventivos: list, listaEmplazamientos: list, indice: int):

    listaCoincidencia = []
    listaSinCoincidencia = []
    for preventivo in listaPreventivos:    
        contadorCoincidencias = 0
        lista = []
        nombrePreventivo = preventivo[0]
        Depurar(nombrePreventivo)
        accion = preventivo[4]
        Depurar(accion)
        emplazamientoCoincidencia = []
        for emplazamiento in listaEmplazamientos:
            if IndiceCoincidencia(None, preventivo[indice], emplazamiento[1]).ratio() > 0.7:
                Depurar("Coincidencia en el preventivo" + nombrePreventivo)                
                latitud = emplazamiento[2]
                altitud = emplazamiento[3]                
                lista.append(["Preventivos", "#FF0000", latitud, altitud, nombrePreventivo,accion,"#71b300"])
                contadorCoincidencias += 1
                emplazamientoCoincidencia.append(emplazamiento)
        if contadorCoincidencias == 0:
            Depurar("Sin coincidencia " + nombrePreventivo)
            listaSinCoincidencia.append(preventivo)
        elif contadorCoincidencias == 1:
            listaCoincidencia.append(lista[0])
        else:
            lista = ImprimirLista(lista, emplazamientoCoincidencia)
            listaCoincidencia.append(lista)
    lista = []
    lista.append(listaCoincidencia)
    lista.append(listaSinCoincidencia)
    return lista

def Coincidencias(listaPreventivos: list, listaEmplazamientos: list):
    listaPreventivos=BuscarCoincidencia(listaPreventivos, listaEmplazamientos, 1)
    listaCoincidencia = listaPreventivos[0]
    listaSinCoincidencia = listaPreventivos[1]
    print(str(len(listaCoincidencia)) + " coincidencias y " + \
          str(len(listaSinCoincidencia)) + " sin coincidencia en la primera vuelta.")
    listaPreventivos=BuscarCoincidencia(listaSinCoincidencia, listaEmplazamientos, 3)
    print(str(len(listaPreventivos[0])) + " coincidencias y " + \
          str(len(listaPreventivos[1])) + " sin coincidencia en la primera vuelta.")
    for item in listaPreventivos[0]:
        listaCoincidencia.append(item)
    for item in listaPreventivos[1]:
        listaCoincidencia.append(["Preventivos", "#FF0000", 0, 0, item[0], item[4], "#FF0000"])
    print("Un total de " + str(len(listaCoincidencia)) + " preventivos y " + str(len(listaPreventivos[1])) + " sin coincidencia")
    return listaCoincidencia

def Depurar(texto):
    #print(texto)
    pass

preventivosExcel = cargar_excel("preventivos.xlsx")
Depurar("Excel preventivos cargado")
emplazamientosExcel = cargar_excel("emplazamientos.xlsx")
Depurar("Excel emplazamiento cargado")
listaEmplazamientos = LeerExcelEmplazamientos(emplazamientosExcel)
Depurar("Lista emplazamientos cargado")
Depurar(listaEmplazamientos)
listaPreventivos = LeerExcelPreventivos(preventivosExcel)
Depurar("Lista preventivos cargado")
Depurar(listaPreventivos)
preventivosExcel.close
emplazamientosExcel.close
lista = Coincidencias(listaPreventivos, listaEmplazamientos)

with open("preventivos.csv", mode="w", newline="", encoding="utf-8") as archivo:
    escritor = csv.writer(archivo)    
    escritor.writerow(["Folder name", "Folder color", "Latitude", "Longitude","Title","Description", "Color"])
    escritor.writerows(lista)

print("Pulsa cualquier tecla para salir")
msvcrt.getch()

#