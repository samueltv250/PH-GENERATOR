from striprtf.striprtf import rtf_to_text
from numpy import sqrt 
from extract_Alcance_PH import *
from WordWrite import *
def ExtrMed(strr):
    a_file = "Input/TXT/"+strr
    llist = []
    with open(a_file, encoding="utf8", errors='ignore') as f:
        for line in f:
            line = str(line.strip()).replace("\\'","")
            if line != "":
                llist.append(line)
    key = ""

    keyit = ""
    i = 0
    dictfin = {}
    listofSegs =[]
    while i < len(llist):
        if llist[i][0:4]== 'Name':
            dictfin[key] = listofSegs
            listofSegs =[]
            ind = llist[i].find("|")
            if "-" in llist[i][6:]:
                  
                key = llist[i][6:].replace(" ","")
                
                


            else:
              
                key = llist[i][6:]
                
            dictfin[key] = []
        if llist[i][0:7]== 'Segment':
            listforKey = []
            if "Curve" in llist[i]:
                keyit = "Curve" 
                listforKey.append(keyit)

                i+=1
                ind = llist[i].find("Radius:")
                keylen = float(llist[i][8:ind].replace(" ", "").replace("m",""))
                listforKey.append(keylen)
                ind += 7
                keyrad = float(llist[i][ind:].replace(" ", "").replace("m",""))
                listforKey.append(keyrad)
                i+=1
                ind = llist[i].find("Tangent:")
                keyDelt = float(llist[i][7:ind].replace(" ", "").replace("m","").replace("(d)",""))
                listforKey.append(keyDelt)
                

                i+=1
                ind = llist[i].find("Course:")
                keychord = float(llist[i][6:ind].replace(" ", "").replace("m",""))
                listforKey.append(keychord)
                
                ind += 8
                keycourse1 = llist[i][ind:ind+12]
              
                keycourse = keycourse1.replace("' ", "'-").replace(" ","deg").replace("'-","'")
                if len(keycourse1)!=12:
                    
                    keycourse = keycourse.replace(keycourse[0],keycourse[0]+"0")
                listforKey.append(keycourse)
                
                i+=3
                ind = llist[i].find("East:")
                keyNorth = float(llist[i][10:ind].replace(" ", "").replace("m",""))
                listforKey.append(keyNorth)
                # print(keyNorth)
                ind += 5
                keyEast = float(llist[i][ind:].replace(" ", "").replace("m",""))
                listforKey.append(keyEast)
                # print(keyEast)
                listofSegs.append(listforKey)




            elif "Line" in llist[i]:
                keyit = "Line" 
                listforKey.append(keyit)

                i+=1
                keycourse = llist[i][8:20].replace("' ", "'-").replace(" ","deg").replace("'-","'")
                if keycourse[-3:] == 'deg':
                    keycourse = keycourse[0:-3]
                    keycourse = keycourse.replace(keycourse[0],keycourse[0]+"0")

                    
                listforKey.append(keycourse)
                ind = llist[i].find("Length:") +8
                leng = float(llist[i][ind:].replace("m", ""))
                listforKey.append(leng)
                
                i+=1
                ind = llist[i].find("East:")
                keyNorth = float(llist[i][7:ind].replace(" ","").replace("m",""))
                ind += 5
                listforKey.append(keyNorth)
                keyEast = float(llist[i][ind:].replace(" ","").replace("m",""))
                listforKey.append(keyEast)
                


                listofSegs.append(listforKey)
        i += 1
    dictfin[key] = listofSegs


        
            
    for key in dictfin:

    
        listofsegs = dictfin[key]


        newlistofSeg = []

        MXnorth = 0
        MXind = 0
        for i in range(len(dictfin[key])):
            segment = listofsegs[i]
            if len(segment)>0:
                if segment[0] == 'Line':
                    if segment[3] > MXnorth:
                        MXnorth = segment[3]
                        MXind = i
                if segment[0] == 'Curve':
                    if segment[6] > MXnorth:
                        MXnorth = segment[6]
                        MXind = i
        
        for i in range(len(dictfin[key])-MXind):
            ind = i+MXind
            newlistofSeg.append(listofsegs[ind])
        for i in range(MXind):
            
            newlistofSeg.append(listofsegs[i])
        # first to last
        if len(newlistofSeg) >0:
            temp = newlistofSeg[0]
            for i in range(len(newlistofSeg)-1):
                newlistofSeg[i] = newlistofSeg[i+1]
            newlistofSeg[len(newlistofSeg)-1] = temp

        dictfin[key] = newlistofSeg
    return dictfin
def midpoint(x1, y1, x2, y2):
    return ((x1+x2)/2, (y1+y2)/2)
def SegToList(seg):
    if seg[0] == 'Line':
        return[seg[1],seg[3],seg[4]] 

    if seg[0] == 'Curve':
        return[seg[5],seg[6],seg[7]]
def cordToComp(Cord):
    deg = int(Cord[1:3])
    if deg < 45:
        return "WE"
    else:
        return "NS"

def getPoints(seg):
    if seg[0] == "Line":
        return[seg[3],seg[4]] 

    if seg[0] == 'Curve':
        return[seg[6],seg[7]]

def getcenteroid(listofSegs):
    point1 =[]
    pointy = []
    for seg in listofSegs:
        point = getPoints(seg)
        point1.append(point[1])
        pointy.append(point[0])


    x = [p for p in point1]
    y = [p for p in pointy]
    centroidMain = (sum(x) / len(point1), sum(y) / len(pointy))
    return centroidMain 
import math
def slope(x1, y1, x2, y2):
    
    x2 -= x1
    x1 -= x1
    y2 -= y1
    y1 -= y1

    return math.atan2(y2,x2)/math.pi*180



def GetDirr(centroidMain,centroid2):
    return slope(centroidMain[0], centroidMain[1],centroid2[0], centroid2[1])

def getNWSEMain(lissegmain,x1, y1, x2, y2,i):
    centmain = getcenteroid(lissegmain)
    
    cent2 = midpoint(x1, y1, x2, y2)
    segmain = lissegmain[i]
    if segmain[0] == 'Line':
        nS,deg,min,seg,wE = degtoDeg(segmain[1])
    else:
        nS,deg,min,seg,wE = degtoDeg(segmain[5])
    deg = deg.replace("deg","")
    slop = 0
    if nS == 'norte':
        slop += 90
        if wE == "este":
            slop -= int(deg)
        elif wE == "oeste":
            slop += int(deg)
    elif nS == 'sur':
        slop -= 90
        if wE == "este":
            slop += int(deg)
        elif wE == "oeste":
            slop -= int(deg)
    mainDIR = ""
    if slop>-45 and slop <= 45:
        mainDIR =  "NS"
    elif slop>45 and slop <= 135:
        mainDIR =  "EW"
    elif slop>-135 and slop <= -45:
        mainDIR =  "EW"
    elif slop>135 or slop <= -135:
        mainDIR =  "NS"
    
    if mainDIR =="NS":
        if cent2[1] >centmain[1]:
            return "NORTH"
        else:
            return "SOUTH"
    elif mainDIR =="EW":
        if cent2[0] >centmain[0]:
            return "EAST"
        else:
            return "WEST"


def getNWSE(lissegmain,x1, y1, x2, y2):
    centmain = getcenteroid(lissegmain)
    
    cent2 = midpoint(x1, y1, x2, y2)
    slop = GetDirr(centmain,cent2)
    if slop>-45 and slop <= 45:
        return "EAST"
    elif slop>45 and slop <= 135:
        return "NORTH"
    elif slop>-135 and slop <= -45:
        return "SOUTH"
    elif slop>135 or slop <= -135:
        return "WEST"

def getNWSE2(lissegmain,lisseg2):
    centmain = getcenteroid(lissegmain)
    
    cent2 = getcenteroid(lisseg2)
    slop = GetDirr(centmain,cent2)
    if slop>-45 and slop <= 45:
        return "EAST"
    elif slop>45 and slop <= 135:
        return "NORTH"
    elif slop>-135 and slop <= -45:
        return "SOUTH"
    elif slop>135 or slop <= -135:
        return "WEST"
def getNWSE3(lissegmain,lisseg2,dirr):
    centmain = getcenteroid(lissegmain)
    
    cent2 = getcenteroid(lisseg2)
    slop = GetDirr(centmain,cent2)
    if dirr == "EAST":
        if slop>-60 and slop <= 60:
            return "EAST"
    elif dirr == "NORTH":
        if slop>30 and slop <= 150:
            return "NORTH"
    elif dirr == "SOUTH":
        if slop>-150 and slop <= -30:
            return "SOUTH"
    elif dirr == "WEST":
        if slop>120 or slop <= -120:
            return "WEST"
    return None

def getNWSE4(lissegmain,x1, y1, x2, y2,dirr):
    centmain = getcenteroid(lissegmain)
    
    cent2 = midpoint(x1, y1, x2, y2)
    slop = GetDirr(centmain,cent2)
    if dirr == "EAST":
        if slop>-50 and slop <= 50:
            return "EAST"
    elif dirr == "NORTH":
        if slop>40 and slop <= 140:
            return "NORTH"
    elif dirr == "SOUTH":
        if slop>-140 and slop <= -40:
            return "SOUTH"
    elif dirr == "WEST":
        if slop>130 or slop <= -130:
            return "WEST"
    return None
def GetLindd(strr):
    out = ExtrMed(strr)

    for i in out:
        out[i] = []

    outMed = ExtrMed(strr)
    outMed.pop("")

    # print(outMed)
    for i in outMed:
        listOfLind = {'NORTH':[],'SOUTH':[],'EAST':[],'WEST':[]}

            
        for p in outMed:
            if p != i:
                if len(outMed[p]) !=0 and len(outMed[i]) !=0:
                    for itemmainind in range(len(outMed[i])):
                        for item2ind in reversed(range(len(outMed[p]))):
                    
                            indmain1 = itemmainind
                            indmain2 = itemmainind+1
                            if indmain2 not in range(len(outMed[i])):
                                indmain2 = 0
                            indsearch1 = item2ind
                            indsearch2 = item2ind+1
                            if indsearch2 not in range(len(outMed[p])):
                                indsearch2 = 0
                            segsearch1 = SegToList(outMed[p][indsearch1])
                            segsearch2 = SegToList(outMed[p][indsearch2])
                            segmain2 = SegToList(outMed[i][indmain1])
                            segmain1 = SegToList(outMed[i][indmain2])

                            if (abs((segsearch1[1] - segmain1[1])) < 1.05) and (abs((segsearch1[2] - segmain1[2])) < 1.05) and (abs((segsearch2[1] - segmain2[1])) < 1.05) and (abs((segsearch2[2] - segmain2[2])) < 1.05):
                                dist = round(sqrt( (segsearch2[1] - segsearch1[1])**2 + (segsearch2[2] - segsearch1[2])**2 ),2)
                                ddir= getNWSEMain(outMed[i],segsearch1[2],segsearch1[1],segsearch2[2],segsearch2[1],indmain2)

                                if "-" == outMed[p][1]:
                                    ddir2 = getNWSE3(outMed[i],outMed[p],ddir)
                                else:
                                    ddir2 = getNWSE4(outMed[i],segsearch1[2],segsearch1[1],segsearch2[2],segsearch2[1],ddir)
                                if len(listOfLind[ddir]) ==0:
                                    listOfLind[ddir].append(p)
                                else:
                                    if ddir2 != None:
                                        listOfLind[ddir].append(p)
        out[i] = listOfLind
    return out


def degtoDeg(degg):
    dictCords = {'N':"norte",'S':'sur','E': 'este','W':'oeste'}
    nS = dictCords[degg[0]]
    deg = degg[1:6]
    min = degg[6:9]
    seg = degg[9:12]
    wE = dictCords[degg[-1]]
    return nS,deg,min,seg,wE

from numero_letras import *
def listToDEg(ExtrMe):
    strr = "partiendo del punto ubicado en la esquina más hacia el norte del lote a describir,"
    for item in ExtrMe:
        strr = "partiendo del punto ubicado en la esquina más hacia el norte del lote a describir,"
        thsttr = strr
        
        for i in range(len(ExtrMe[item])):
            segment = ExtrMe[item][i]
            if segment[0] == 'Curve':
                
                rad = round(segment[2],2)
                delt = round(segment[3],2)
                len1 = round(segment[1],2)

                nS,deg,min,seg,wE = degtoDeg(segment[5])
             
                len2 = round(segment[4],2)
                thsttr2 = " de este punto con un segmento circular,cuyo radio es de "+check_content(str(rad)+'m')+" ("+str(rad)+"m"+"), un delta de "+numero_a_Grads(delt)+" ("+str(delt)+'\u00B0'+") y un largo de curva de "+check_content(str(len1)+"m")+" ("+round2(len1)+"m)"+", con rumbo "+nS+" "+check_content(deg)+", "+check_content(min)+", "+check_content(seg)+" "+wE+", se miden "+check_content(str(len2)+'m')+" ("+round2(len2)+'m' +") hasta llegar al punto "
                thsttr += thsttr2
            elif segment[0] == 'Line':
                
                len1 = round(segment[2],2)

                nS,deg,min,seg,wE = degtoDeg(segment[1])
             
                
                thsttr2 = " de este punto con rumbo "+nS+" "+check_content(deg)+", "+check_content(min)+", "+check_content(seg)+" "+wE+", se miden "+check_content(str(len1)+'m')+" ("+str(round2(len1))+'m' +") hasta llegar al punto "
                thsttr += thsttr2


            if i+1 == len(ExtrMe[item]):
                thsttr2 = " que sirvió de partida de esta descripción. ---------------"
                thsttr += thsttr2
            else:
                thsttr2 = numero_a_letras(i+2)+" ("+str(i+2)+"),"
                thsttr += thsttr2

        ExtrMe[item] = thsttr

    return ExtrMe


def listToDEgMejor(ExtrMe):
    strr = "partiendo del punto más al norte,"
    for item in ExtrMe:
        strr = "partiendo del punto ubicado en la esquina más hacia el norte del lote a describir,"
        thsttr = strr
        
        for i in range(len(ExtrMe[item])):
            segment = ExtrMe[item][i]
            if segment[0] == 'Curve':
                
                rad = round(segment[2],2)
                delt = round(segment[3],2)
                len1 = round(segment[1],2)

                nS,deg,min,seg,wE = degtoDeg(segment[5])
             
                len2 = round(segment[4],2)
                thsttr2 = " de este punto con un segmento circular,cuyo radio es de "+check_content(str(rad)+'m')+" ("+str(rad)+"m"+"), un delta de "+numero_a_Grads(delt)+" ("+str(delt)+'\u00B0'+") y un largo de curva de "+check_content(str(len1)+"m")+" ("+round2(len1)+"m)"+", con rumbo "+nS+" "+check_content(deg)+", "+check_content(min)+", "+check_content(seg)+" "+wE+", se miden "+check_content(str(len2)+'m')+" ("+round2(len2)+'m' +") hasta llegar al punto "
                thsttr += thsttr2
            elif segment[0] == 'Line':
                
                len1 = round(segment[2],2)

                nS,deg,min,seg,wE = degtoDeg(segment[1])
             
                
                thsttr2 = " de este punto con rumbo "+nS+wE+", se miden "+check_content(str(len1)+'m')+" ("+str(round2(len1))+'m' +") hasta llegar al siguiente punto "
                thsttr += thsttr2


            if i+1 == len(ExtrMe[item]):
                thsttr2 = " que sirvió de partida de esta descripción. ---------------"
                thsttr += thsttr2
            else:
                thsttr2 = numero_a_letras(i+2)+" ("+str(i+2)+"),"
                thsttr += thsttr2

        ExtrMe[item] = thsttr

    return ExtrMe

def merge_two_dicts(x, y):
    z = x.copy()   # start with keys and values of x
    z.update(y)    # modifies z with keys and values of y
    return z
def aregmed(medidasMain,medidasMejoras,mainDICT,linder):
    dictCure = {}

    z = {}
    for i in mainDICT:
        z = merge_two_dicts(mainDICT[i], z) 
    
    for i in medidasMejoras:
        id = ""
        if i in z:
            model = z[i]
        else:
            model = ""
        if len(medidasMain[i]) > 0 and len(medidasMain[i][0])>1:
            id = createCod(medidasMain[i], model)
        dictCure[id] = [medidasMejoras[i],linder[i]]

    for i in medidasMain:
        if i in z:
            model = z[i]
            id = createCod(medidasMain[i], model)
            if id in dictCure and i not in medidasMejoras:
                lind1 = dictCure[id][1]
                lind2 = linder[i]
                if lind2['NORTH'] == lind1['NORTH'] or lind2['SOUTH'] == lind1['SOUTH']or lind2['EAST'] == lind1['EAST']or lind2['WEST'] == lind1['WEST']:
                    medidasMejoras[i] = dictCure[id][0]

    return medidasMejoras

    


    
def createCod(segList, model):
    id = ""
    for seg in segList:
        if seg[0].lower() == 'line':
            id += str(seg[1])+str(seg[2])+str(model)
    return id



def GetAllText(mainDICT):
    linder = GetLindd("Medidas.txt")
    areaTerr = DictManzTerr()


    medidasMain1 = ExtrMed("Medidas.txt")
    medidasMejoras1 = ExtrMed("Medidascerr.txt")
    medidasMejoras = aregmed(medidasMain1,medidasMejoras1,mainDICT,linder)

    medidasMain = listToDEg(medidasMain1)
    medidasMejoras = listToDEgMejor(medidasMejoras)

    superfMain = {}
    linderosfMain = {}

    
    for i in linder:
        for cord in linder[i]:
            if len(linder[i][cord]) == 0:
                linder[i][cord].append("Not Found")
                print("Warning: No se encontro linero para "+i+ " en la direccion "+cord)

    
    for i in areaTerr:
        superfMain[i] = "La superficie total del lote que acabamos de describir es de "+check_content(str(areaTerr[i])+'mts2')+" ("+round2(areaTerr[i])+'m\u00b2'+"). ----------------------------------------------"
    for i in linder:
        for cord in linder[i]:
            linder[i][cord] = list(dict.fromkeys(linder[i][cord]))
        linderosfMain[i] = parcelTOTXT(linder[i])
       

    
        
    return  linderosfMain, superfMain,medidasMain,medidasMejoras
    
    # print(linderosfMain["E-166"])
    # print(superfMain["E-166"])
    # print(medidasMain["E-166"])
    # print(medidasMejoras['E-166'])



def WRITEPRot(numb):
    dictAreasCerr,dictAreasAbiert,dictAreasPav = extractTableDat()
    dictAreasTotal = {}
    for i in dictAreasCerr:
        dictAreasTotal[i] = round(dictAreasCerr[i]+dictAreasAbiert[i]+dictAreasPav[i],2)
    mainDICT = extractDictEt_mod()
    mattDict = extractMaterialDat()
    descDict = extractDescrip()
    tabla = TablaFinal()
    for i in tabla:
        temp = tabla[i]
        tabla[i] = {}
        for p in range(len(temp)):
            ttemp = temp[p].pop(0)
            tabla[i][ttemp] = temp[p]
    
    linderosfMain, superfMain,medidasMain,medidasMejoras = GetAllText(mainDICT)

    document = docx.Document()
    document.styles['Normal'].font.name = 'Arial'

    style = document.styles.add_style('bca', WD_STYLE_TYPE.PARAGRAPH)
    style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    font = style.font
    font.bold = True
    font.name = 'Garamond'

    style = document.styles.add_style('bcal', WD_STYLE_TYPE.PARAGRAPH)
    style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    font = style.font
    font.underline = True
    font.bold = True
    font.name = 'Garamond'

    style = document.styles.add_style('ba', WD_STYLE_TYPE.PARAGRAPH)
    style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    font = style.font
    font.bold = True
    font.name = 'Garamond'

    style = document.styles.add_style('maiNN', WD_STYLE_TYPE.PARAGRAPH)
    style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    font = style.font
    font.name = 'Garamond'

    for i in mainDICT[numb]:
        uniman = i
        
        mod = mainDICT[numb][i]
        tipo = mod
        model = mod
        mod = mod.split()
        model = "VIVIENDA "+model.upper()
        if len(mod) == 2:
            tipo = mod[0]
            model = mod[1]
            model = "VIVIENDA "+tipo.upper()+" TIPO "+model.upper()


        # must currect
        # ----------------
        ind = uniman.find("-")
        manz = uniman[:ind]
        uni = uniman[ind+1:]
        paragrap = document.add_paragraph("DESCRIPCIÓN DEL LOTE "+manz+"-"+numero_a_letras(int(uni)).upper()+" ("+uniman+") ---------------------------------------------------------",style='ba')
        paragrap = document.add_paragraph("MEDIDAS: ",style='maiNN')
        paragrap.add_run(medidasMain[i])
        paragrap.add_run("---------")
        paragrap = document.add_paragraph("SUPERFICIE: ",style='maiNN')
        paragrap.add_run(superfMain[i])

        paragrap = document.add_paragraph("LINDEROS: ",style='maiNN')
        paragrap.add_run(linderosfMain[i])
        paragrap = document.add_paragraph("DECLARACIÓN DE MEJORAS DE LOTE: "+uniman+" --------------------------",style='maiNN')
        paragrap = document.add_paragraph(model +": ",style='maiNN')
        paragrap.add_run(mattDict[mainDICT[numb][i]])
        paragrap.add_run(" -------------------------------------------------------------------------------------------------------- ")
        paragrap.add_run(descDict[mainDICT[numb][i]])
        paragrap = document.add_paragraph("MEDIDAS DE LA UNIDAD INMOBILIARIA: ",style='maiNN')
        if i in medidasMejoras:
            paragrap.add_run(medidasMejoras[i])
        else:
            print("Warning: EL AREA CERRADO PARA PARCELA "+i+" NO SE ENCUENTRA EN CIVIL 3D")
            paragrap.add_run("<EL AREA CERRADO NO SE ENCUENTRA EN CIVIL 3D>")
        paragrap = document.add_paragraph("SUPERFICIE TOTAL: La superficie total de la unidad inmobiliaria es de: --------------------------------",style='maiNN')
        paragrap = document.add_paragraph("ÁREA CERRADA: ",style='maiNN')
        paragrap.add_run(check_content(MakeSQRmtrr(dictAreasCerr[mainDICT[numb][i]])+"mts2")+" ("+MakeSQRmtrr(dictAreasCerr[mainDICT[numb][i]])+'m\u00b2'+"). ----------------------------------------------------------------------------------------------------------")
        paragrap = document.add_paragraph("ÁREA ABIERTA: ",style='maiNN')
        paragrap.add_run(check_content(MakeSQRmtrr(dictAreasAbiert[mainDICT[numb][i]])+"mts2")+" ("+MakeSQRmtrr(dictAreasAbiert[mainDICT[numb][i]])+'m\u00b2'+").")
        paragrap = document.add_paragraph("PAVIMENTO: ",style='maiNN')
        paragrap.add_run(check_content(MakeSQRmtrr(dictAreasPav[mainDICT[numb][i]])+"mts2")+" ("+MakeSQRmtrr(dictAreasPav[mainDICT[numb][i]])+'m\u00b2'+"). -------------")
        paragrap = document.add_paragraph("ÁREA TOTAL DE CONSTRUCCIÓN: ",style='maiNN')
        paragrap.add_run(check_content(MakeSQRmtrr(dictAreasTotal[mainDICT[numb][i]])+"mts2")+" ("+MakeSQRmtrr(dictAreasTotal[mainDICT[numb][i]])+'m\u00b2'+"). ---------------------------------------------------------------------------------")
        paragrap = document.add_paragraph("VALOR DE TERRENO: ",style='maiNN')
        paragrap.add_run(check_content("B/"+MakeSQRmtrr(tabla[numb][i][1])).upper()+" (US$"+MakeSQRmtrr(tabla[numb][i][1])+"). -------------------------------------------------------")
        paragrap = document.add_paragraph("VALOR DE MEJORAS: ",style='maiNN')
        paragrap.add_run(check_content("B/"+MakeSQRmtrr(tabla[numb][i][0])).upper()+" (US$"+MakeSQRmtrr(tabla[numb][i][0])+"). ---------------------------------------------------------")
        paragrap = document.add_paragraph("VALOR TOTAL: ",style='maiNN')
        paragrap.add_run(check_content("B/"+MakeSQRmtrr(tabla[numb][i][2])).upper()+" (US$"+MakeSQRmtrr(tabla[numb][i][2])+"). ------------------------------------------------------------------------")
        paragrap = document.add_paragraph("PORCENTAJE DE PARTICIPACIÓN: ",style='maiNN')
        paragrap.add_run(numero_a_letras(MakeSQRmtrr(tabla[numb][i][3])).upper()+" POR CIENTO ("+MakeSQRmtrr(tabla[numb][i][3])+"%). ------------------------------------------------------------------------------------------------------------------")
    document.save("Output_Docx/Parts/prot-"+str(numb)+".docx")





