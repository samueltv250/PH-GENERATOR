import sys
import pandas as pd
import os



TablaPrevia = False



def CheckTab(state):
    global TablaPrevia
    TablaPrevia = state



def listOfDat():
    workbook = pd.read_excel('Input/EXCEL/DATA-GENERAL.xlsx')

    listofdatt = workbook.values.tolist()   
    return listofdatt      
def getVals(pnumb):
    workbook = pd.read_excel('Input/EXCEL/TIPO-MODELO.xlsx')

    workbook = workbook[workbook.columns[::-1]]
    workbook = workbook.T

    dictionaryObject = workbook.to_dict();
    tipomod = {}
    dicdeff = {}
    dicTOT = {}
    tttopi = {}
    for i in dictionaryObject:
        modd = dictionaryObject[i]["MODELO"]
        if modd =="na" or str(modd) == "nan":
            modd = ""
        else:
            modd = " " + modd
        dicdeff[(dictionaryObject[i]["TIPO"]+modd)] = str(dictionaryObject[i]["DESCRIPCION"])
        dicTOT[(dictionaryObject[i]["TIPO"]+modd)] = 0
        tipomod[(dictionaryObject[i]["TIPO"]+modd)] = []
        tttopi[(dictionaryObject[i]["TIPO"]+modd)] =[]
    
    

    manzana = ""

    for file in os.listdir("Input/EXCEL/Alcance"):
        if file.endswith(".xlsx"):
            workbook = pd.read_excel("Input/EXCEL/Alcance/"+file)
            workbook = workbook[workbook.columns[::-1]]
            workbook = workbook.T

            dictionaryObject = workbook.to_dict();


            for p in dictionaryObject:
                manzana = str(dictionaryObject[p]["Manzana"]).replace("0","")
                unimanz = ""
                unimanz = manzana+"-"+str(int(dictionaryObject[p]["Unidad"]))
                
                for i in tttopi:
            
                    if i in dictionaryObject[p]["Modelo"]:
                        tttopi[i].append(unimanz)
    dictTemp = {}               
    for p in tttopi:
        if len(tttopi[p]) != 0:
            dictTemp[p] = len(tttopi[p])
    dicTOT = dictTemp
    


    workbook = pd.read_excel('Input/EXCEL/DATA-GENERAL.xlsx')

    listofdatt = workbook.values.tolist()

    faSe = listofdatt[0][5]




    try:
   
        workbook = pd.read_excel('Input/EXCEL/Alcance/Etapa-'+str(int(pnumb))+'.xlsx')
    

    except Exception as e:
        raise Exception("Etapa-"+str(int(pnumb))+".xlsx no se encuentra en Alcance y se estan incorporando "+str(int(pnumb))+" Etapas en esta corrida")
    workbook = workbook[workbook.columns[::-1]]
    workbook = workbook.T

    dictionaryObject = workbook.to_dict();

    
    


    for p in dictionaryObject:
        unimanz = ""
        unimanz = str(dictionaryObject[p]["Manzana"]).replace("0","")+"-"+str(int(dictionaryObject[p]["Unidad"]))
        
        for i in tipomod:
      
            if i in dictionaryObject[p]["Modelo"]:
                tipomod[i].append(unimanz)
     
 
    
    
    

    return listofdatt , tipomod , dicdeff, dicTOT, manzana


def extractTableDat():
    dictAreasCerr = {}
    dictAreasAbiert = {}
    dictAreasPav = {}


    workbook = pd.read_excel('Input/EXCEL/TIPO-MODELO.xlsx')
    workbook = workbook[workbook.columns[::-1]]
    workbook = workbook.T
    dictionaryObject = workbook.to_dict();

    for i in dictionaryObject:
        modd = dictionaryObject[i]["MODELO"]
        if modd =="na" or str(modd) == "nan":
            modd = ""
        else:
            modd = " " + modd
        dictAreasCerr[(dictionaryObject[i]["TIPO"]+modd)] = dictionaryObject[i]["AREA CERRADA"]
        dictAreasAbiert[(dictionaryObject[i]["TIPO"]+modd)] = dictionaryObject[i]["AREA ABIERTA"]
        dictAreasPav[(dictionaryObject[i]["TIPO"]+modd)] = dictionaryObject[i]["AREA PAVIMENTO"]
    return dictAreasCerr,dictAreasAbiert,dictAreasPav

def extractMaterialDat():

    dictAreasPav = {}


    workbook = pd.read_excel('Input/EXCEL/TIPO-MODELO.xlsx')
    workbook = workbook[workbook.columns[::-1]]
    workbook = workbook.T
    dictionaryObject = workbook.to_dict();

    for i in dictionaryObject:
        modd = dictionaryObject[i]["MODELO"]
        if modd =="na" or str(modd) == "nan":
            modd = ""
        else:
            modd = " " + modd

        dictAreasPav[(dictionaryObject[i]["TIPO"]+modd)] = dictionaryObject[i]["MATERIALES"]
    return dictAreasPav

def extractDescrip():
    
    dictAreasPav = {}


    workbook = pd.read_excel('Input/EXCEL/TIPO-MODELO.xlsx')
    workbook = workbook[workbook.columns[::-1]]
    workbook = workbook.T
    dictionaryObject = workbook.to_dict();

    for i in dictionaryObject:
        modd = dictionaryObject[i]["MODELO"]
        if modd =="na" or str(modd) == "nan":
            modd = ""
        else:
            modd = " " + modd

        dictAreasPav[(dictionaryObject[i]["TIPO"]+modd)] = dictionaryObject[i]["DESCRIPCION"]
    return dictAreasPav
    
def extractAlcDictsCerr():
    dictAreasCerr,dictAreasAbiert,dictAreasPav = extractTableDat()
    arr =  os.listdir('Input/EXCEL/Alcance/')
    NUM = 0
    for i in arr:
        if i[0] == "E":
            NUM +=1

    arr =NUM
    dictEtaps = {}
    for pnumb in range(arr):

        workbook = pd.read_excel('Input/EXCEL/Alcance/Etapa-'+str(int(pnumb+1))+'.xlsx')

        workbook = workbook[workbook.columns[::-1]]
        workbook = workbook.T

        dictionaryObject = workbook.to_dict();
        listwitAreas = []
        
        
        

        for p in dictionaryObject:
            unimanz = ""
            unimanz = str(dictionaryObject[p]["Manzana"]).replace("0","")+"-"+str(int(dictionaryObject[p]["Unidad"]))
            lln = len(dictAreasCerr)
            countr = 0
            hins = False
            for i in dictAreasCerr:
        
                if i in dictionaryObject[p]["Modelo"]:
                    listwitAreas.append([unimanz,dictAreasCerr[i]])
                    hins = True
                else:
                    countr+=1
                if hins == False and countr == lln:
                    raise Exception(dictionaryObject[p]["Modelo"]+" no se encuentra en TIPO-MODELO.xlsx")


        
        dictEtaps[pnumb+1] = listwitAreas
    return dictEtaps
import collections
def extractDictEt_mod():
    dictAreasCerr,dictAreasAbiert,dictAreasPav = extractTableDat()
    liis = os.listdir('Input/EXCEL/Alcance/')
    arr = 0

    for i in liis:
        if i[0] == 'E':
            arr +=1
    

    dictEtaps = {}
    for pnumb in range(arr):
        listofdatt , tipomod , dicdeff, dicTOT, manzana = getVals(pnumb+1)
        dictEtaps[pnumb+1] = tipomod
    dictout = {}
    for i in dictEtaps:

        dictout[i] = {}
        for mod in dictEtaps[i]:
            if len(dictEtaps[i][mod]) !=0:
                for p in range(len(dictEtaps[i][mod])):
                    dictout[i][dictEtaps[i][mod][p]]=mod

        dictout[i] = sorted(dictout[i].items(), key=lambda x: int(x[0][2:]), reverse=False)
        tempdict = {}

        for p in dictout[i]:
            tempdict[p[0]] = p[1]
        dictout[i] = tempdict
        
    



    return dictout

def DictManzTerr():
    workbook = pd.read_excel('Input/EXCEL/Areas-Parcela.xlsx')

    workbook = workbook[workbook.columns[::-1]]
    workbook = workbook.T
    dictOut = {}
    dictionaryObject = workbook.to_dict();
    for p in dictionaryObject:
        unimanz = str(dictionaryObject[p]["Parcel Name"]).replace(" ","")
        dictOut[unimanz] = dictionaryObject[p]["Square Meters"]
    return dictOut

def VariablesCalcCostInc():
    workbook = pd.read_excel('Input/EXCEL/DATA-GENERAL.xlsx')

    listofdatt = workbook.values.tolist()
    valor = listofdatt[0][2]
    areaTerreno = listofdatt[0][7]
    areaConstruccion = listofdatt[0][8]
    areaComerc = listofdatt[3][0]
    costounitario = round(valor/areaTerreno,2)

    costoporM2 = round((int(areaTerreno)-int(areaComerc))*costounitario,2)

    
    costIncorp = round(costoporM2/areaConstruccion,2)
    return costIncorp,costounitario

def Variabl():
    workbook = pd.read_excel('Input/EXCEL/DATA-GENERAL.xlsx')
    costIncorp,costounitario = VariablesCalcCostInc()
    listofdatt = workbook.values.tolist()
    valor = listofdatt[0][2]
    restolibre = round(listofdatt[3][0]*costounitario,2)
    AREAconst = round(listofdatt[0][8],2)
    totmnrest = valor - restolibre

    return totmnrest,restolibre,AREAconst


def valorTerreno():
    costINCORP,unss = VariablesCalcCostInc()
    dictout = DictManzTerr()
    AreaTotal = 0
    for i in dictout:
        temp = dictout[i]
        dictout[i] = temp
        if '-' in i:
            AreaTotal += dictout[i]

    return dictout,AreaTotal,costINCORP



def ValMejoras(areaCerr):
    if areaCerr <= 60:
        return areaCerr*360
    elif areaCerr >=121:
        return areaCerr*900
    else:
        return areaCerr*460
import math
from decimal import Decimal

isFinal = False

def declareFinal(isFin):
    global isFinal
    isFinal = isFin
import re
def getOldTable():
    workbook = pd.read_excel('Input/VerifyTable/Table-Old.xlsx')
    
    workbook = workbook[workbook.columns[::-1]]
    workbook = workbook.T
    dictOut = {}
    dictionaryObject = workbook.to_dict();

    for i in dictionaryObject:
        modd = dictionaryObject[i]["UNIDAD "]
        if "-" not in modd:
            lett = re.sub(r'[^a-zA-Z]', '', modd)
            Numb = ""
            # iterate over characters in s
            for ch in modd:
                if ch.isdigit():
                    Numb += ch
            modd = lett +"-"+ Numb
        dictOut[modd] = [dictionaryObject[i]["VALOR DE TERRENO "],dictionaryObject[i]["VALOR DE MEJORAS "],dictionaryObject[i]["%"]]

    return dictOut


    



def TablaFinal():
    oldTable = getOldTable()
    initdict = extractAlcDictsCerr()
    dictValTerr,AreaTotal,costINCORP = valorTerreno()
    totMenRest, restoLibre,AREAconst = Variabl()
    # print(totMenRest)
    reserv = totMenRest
    areaRes = AREAconst
    terRem = 0
    percRem = 0
    remmP = 100
    remmA = reserv
    AreaVIz = AREAconst
    ValPorM2 = Decimal(reserv)/Decimal(areaRes)
    Notin = False
    for i in initdict:
 
        for p in range(len(initdict[i])):  
            
            if initdict[i][p][0] in dictValTerr:
                # edited function
                # valTerr = math.floor(dictValTerr[initdict[i][p][0]]* ValPorM2 * 100)/100.0
                # areaRes -= dictValTerr[initdict[i][p][0]]
                # reserv -= valTerr,2
          
                # unrounded
                # valTerr = round(dictValTerr[initdict[i][p][0]]* ValPorM2,2)

                if initdict[i][p][0] in oldTable and TablaPrevia == True:
                    initdict[i][p][1] = oldTable[initdict[i][p][0]][1]
                    valTerr = oldTable[initdict[i][p][0]][0]
                    areaporcRoun = oldTable[initdict[i][p][0]][2]
                    Notin = False
                else:
                    if Notin == False:
                        Notin = True
                        ValPorM2 = Decimal(remmA)/Decimal(AreaVIz)
                    
                    initdict[i][p][1] = ValMejoras(initdict[i][p][1])

                    

                    # Og function
                    valTotTerr = Decimal(dictValTerr[initdict[i][p][0]])* Decimal(ValPorM2)
                    valTerr = math.floor(valTotTerr * 100)/100
            
                    terRem += Decimal(valTotTerr)-Decimal(valTerr)

                    if terRem >= 0.01:
                        terRem -= Decimal(0.01)
                        valTerr += 0.01

                    

                    areaPorc = (float(valTotTerr)/float(totMenRest)) * 100
                    areaporcRoun = math.floor(areaPorc*100)/100
                    percRem +=  areaPorc-areaporcRoun
                    if percRem >= 0.01:
                        percRem -= 0.01
                        areaporcRoun += 0.01

                
                if i == len(initdict) and p+1 == len(initdict[i]):
                    valTerr = remmA
                    areaporcRoun = remmP
                else:
                    AreaVIz -= dictValTerr[initdict[i][p][0]]
                    remmP -= areaporcRoun
                    remmA -= valTerr
                initdict[i][p].append(valTerr)

            else:
                raise Exception(initdict[i][p][0]+" no se encuentra en Areas-Parcela.xlsx")
            initdict[i][p].append(round(round(initdict[i][p][2],2) + round(initdict[i][p][1],2),2))
        
            initdict[i][p].append(areaporcRoun)

    return initdict
