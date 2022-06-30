import sys
import pandas as pd
           
           
def getVals():
    workbook = pd.read_excel('TIPO-MODELO.xlsx')

    workbook = workbook[workbook.columns[::-1]]
    workbook = workbook.T

    dictionaryObject = workbook.to_dict();
    tipomod = {}
    dicdeff = {}
    dicTOT = {}
    for i in dictionaryObject:
        modd = dictionaryObject[i]["MODELO"]
        if modd =="na":
            modd = ""
        else:
            modd = " " + modd
        dicdeff[(dictionaryObject[i]["TIPO"]+modd)] = str(dictionaryObject[i]["DESCRIPCION"])
        dicTOT[(dictionaryObject[i]["TIPO"]+modd)] = str(dictionaryObject[i]["TOTAL"])
        tipomod[(dictionaryObject[i]["TIPO"]+modd)] = []
                










    workbook = pd.read_excel('MM5. Alcance de PH.xlsx')

    workbook = workbook[workbook.columns[::-1]]
    workbook = workbook.T

    dictionaryObject = workbook.to_dict();





    for p in dictionaryObject:
        unimanz = ""
        unimanz = str(dictionaryObject[p]["Manzana"])[-1]+"-"+str(int(dictionaryObject[p]["Unidad"]))
        
        for i in tipomod:
      
            if i in dictionaryObject[p]["Modelo"]:
                tipomod[i].append(unimanz)
     


    workbook = pd.read_excel('DATA-GENERAL.xlsx')

    listofdatt = workbook.values.tolist()

    return listofdatt , tipomod , dicdeff, dicTOT


