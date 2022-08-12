#!/usr/bin/python
# -*- coding: utf-8 -*-
import re
MONEDA_SINGULAR = 'balboa'
MONEDA_PLURAL = 'balboas'

CENTIMOS_SINGULAR = 'centavo'
CENTIMOS_PLURAL = 'centavos'

AREA_SINGULAR = 'metro cuadrados'
AREA_PLURAL = 'metros cuadrados'

CAREA_SINGULAR = 'decímetro cuadrados'
CAREA_PLURAL = 'decímetros cuadrados'


DISTG_SINGULAR = 'grado'
DISTG_PLURAL = 'grados'

CDISTG_SINGULAR = 'centésimo de grado'
CDISTG_PLURAL = 'centésimas de grados'

DIST_SINGULAR = 'metro'
DIST_PLURAL = 'metros'

CDIST_SINGULAR = 'centímetro'
CDIST_PLURAL = 'centímetros'


MAX_NUMERO = 999999999999

UNIDADES = (
    'cero',
    'uno',
    'dos',
    'tres',
    'cuatro',
    'cinco',
    'seis',
    'siete',
    'ocho',
    'nueve'
)

DECENAS = (
    'diez',
    'once',
    'doce',
    'trece',
    'catorce',
    'quince',
    'dieciséis',
    'diecisiete',
    'dieciocho',
    'diecinueve'
)

DIEZ_DIEZ = (
    'cero',
    'diez',
    'veinte',
    'treinta',
    'cuarenta',
    'cincuenta',
    'sesenta',
    'setenta',
    'ochenta',
    'noventa'
)

CIENTOS = (
    '_',
    'ciento',
    'doscientos',
    'trescientos',
    'cuatrocientos',
    'quinientos',
    'seiscientos',
    'setecientos',
    'ochocientos',
    'novecientos'
)

def numero_a_letras(numero):
    numero =float(numero)

    numero_entero = int(numero)
    if numero_entero > MAX_NUMERO:
        raise OverflowError('Número demasiado alto')
    if numero_entero < 0:
        return 'menos %s' % numero_a_letras(abs(numero))
    letras_decimal = ''
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    if parte_decimal > 9:
        letras_decimal = 'punto %s' % numero_a_letras(parte_decimal)
    elif parte_decimal > 0:
        letras_decimal = 'punto cero %s' % numero_a_letras(parte_decimal)
    if (numero_entero <= 99):
        resultado = leer_decenas(numero_entero)
    elif (numero_entero <= 999):
        resultado = leer_centenas(numero_entero)
    elif (numero_entero <= 999999):
        resultado = leer_miles(numero_entero)
    elif (numero_entero <= 999999999):
        resultado = leer_millones(numero_entero)
    else:
        resultado = leer_millardos(numero_entero)
    resultado = resultado.replace('uno mil', 'un mil')
    resultado = resultado.strip()
    resultado = resultado.replace(' _ ', ' ')
    resultado = resultado.replace('  ', ' ')
    if parte_decimal > 0:
        resultado = '%s %s' % (resultado, letras_decimal)
    return resultado



def numero_a_moneda(numero):
    numero_entero = int(numero)
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    centimos = ''
    if parte_decimal == 1:
        centimos = CENTIMOS_SINGULAR
    else:
        centimos = CENTIMOS_PLURAL
    moneda = ''
    if numero_entero == 1:
        moneda = MONEDA_SINGULAR
    else:
        moneda = MONEDA_PLURAL
    letras = numero_a_letras(numero_entero)
    letras = letras.replace('uno', 'un')
    letras_decimal = 'con %s %s' % (numero_a_letras(parte_decimal).replace('uno', 'un'), centimos)
    letras = '%s %s %s' % (letras, moneda, letras_decimal)
    return letras

def numero_a_distancia(numero):
    numero_entero = int(numero)
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    centimos = ''
    if parte_decimal == 1:
        centimos = CDIST_SINGULAR
    else:
        centimos = CDIST_PLURAL
    moneda = ''
    if numero_entero == 1:
        moneda = DIST_SINGULAR
    else:
        moneda = DIST_PLURAL
    letras = numero_a_letras(numero_entero)
    letras = letras.replace('uno', 'un')
    letras_decimal = 'con %s %s' % (numero_a_letras(parte_decimal).replace('uno', 'un'), centimos)
    letras = '%s %s %s' % (letras, moneda, letras_decimal)
    return letras

def numero_a_Grads(numero):
    numero_entero = int(numero)
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    centimos = ''
    if parte_decimal == 1:
        centimos = CDISTG_SINGULAR
    else:
        centimos = CDISTG_PLURAL
    moneda = ''
    if numero_entero == 1:
        moneda = DISTG_SINGULAR
    else:
        moneda = DISTG_PLURAL
    letras = numero_a_letras(numero_entero)
    letras = letras.replace('uno', 'un')
    letras_decimal = 'con %s %s' % (numero_a_letras(parte_decimal).replace('uno', 'un'), centimos)
    letras = '%s %s %s' % (letras, moneda, letras_decimal)
    return letras
    
def numero_a_area(numero):
    numero_entero = int(numero)
    parte_decimal = int(round((abs(numero) - abs(numero_entero)) * 100))
    centimos = ''
    if parte_decimal == 1:
        centimos = CAREA_SINGULAR
    else:
        centimos = CAREA_PLURAL
    moneda = ''
    if numero_entero == 1:
        moneda = AREA_SINGULAR
    else:
        moneda = AREA_PLURAL
    letras = numero_a_letras(numero_entero)
    letras = letras.replace('uno', 'un')
    letras_decimal = 'con %s %s' % (numero_a_letras(parte_decimal).replace('uno', 'un'), centimos)
    letras = '%s %s %s' % (letras, moneda, letras_decimal)
    return letras

def leer_decenas(numero):
    if numero < 10:
        return UNIDADES[numero]
    decena, unidad = divmod(numero, 10)
    if numero <= 19:
        resultado = DECENAS[unidad]
    elif 21 <= numero <= 29:
        resultado = 'veinti%s' % UNIDADES[unidad]
    else:
        resultado = DIEZ_DIEZ[decena]
        if unidad > 0:
            resultado = '%s y %s' % (resultado, UNIDADES[unidad])
    return resultado

def leer_centenas(numero):
    centena, decena = divmod(numero, 100)
    if decena == 0:
        resultado = 'cien'
    else:
        resultado = CIENTOS[centena]
        if decena > 0:
            resultado = '%s %s' % (resultado, leer_decenas(decena))
    return resultado

def leer_miles(numero):
    millar, centena = divmod(numero, 1000)
    resultado = ''
    if (millar == 1):
        resultado = ''
    if (millar >= 2) and (millar <= 9):
        resultado = UNIDADES[millar]
    elif (millar >= 10) and (millar <= 99):
        resultado = leer_decenas(millar)
    elif (millar >= 100) and (millar <= 999):
        resultado = leer_centenas(millar)
    resultado = '%s mil' % resultado
    if centena > 0:
        resultado = '%s %s' % (resultado, leer_centenas(centena))
    return resultado

def leer_millones(numero):
    millon, millar = divmod(numero, 1000000)
    resultado = ''
    if (millon == 1):
        resultado = ' un millon '
    if (millon >= 2) and (millon <= 9):
        resultado = UNIDADES[millon]
    elif (millon >= 10) and (millon <= 99):
        resultado = leer_decenas(millon)
    elif (millon >= 100) and (millon <= 999):
        resultado = leer_centenas(millon)
    if millon > 1:
        resultado = '%s millones' % resultado
    if (millar > 0) and (millar <= 999):
        resultado = '%s %s' % (resultado, leer_centenas(millar))
    elif (millar >= 1000) and (millar <= 999999):
        resultado = '%s %s' % (resultado, leer_miles(millar))
    return resultado

def leer_millardos(numero):
    millardo, millon = divmod(numero, 1000000)
    return '%s millones %s' % (leer_miles(millardo), leer_millones(millon))

def round2(numb):
    numb = round(float(numb),2)
    numb = str(numb)
    numorg = numb
    ind = numb.find('.')
    numb = numb[ind+1:]
    if len(numb) <2:
        numorg += "0"
    return numorg
def check_content(numbstring):
    numbstring = str(numbstring)
    temp = ""
    if str(numbstring) == "NaN":
        raise "Falta Valor Para esta etapa"
    if numbstring[-4:] == "mts2":
        temp = numbstring.replace("mts2","").replace(" ","").replace(",","")
        temp = numero_a_area(float(temp))
    elif numbstring[0:2] == "B/":
        temp = numbstring.replace("B/.","").replace("B/","").replace(" ","").replace(",","")
        temp = numero_a_moneda(float(temp))
    elif numbstring[-1] == "m":
        temp = numbstring.replace("m","").replace(" ","").replace(",","")
        temp = numero_a_distancia(float(temp))
    elif numbstring[-3:] == "deg":
        temp = numbstring.replace("deg","").replace(" ","").replace(",","")
        if int(temp) ==1:
            return "un ("+str(temp)+'\u00B0'+") grado"

        
        temp = numero_a_letras(float(temp))+" ("+str(temp)+'\u00B0'+") grados"
    elif numbstring[-1] == "’" or numbstring[-1] == "'":
        temp = numbstring.replace("’","").replace("'","").replace(" ","").replace(",","")
        
        if int(temp) ==1:
            return "un ("+str(temp)+"') minuto"
        temp = numero_a_letras(float(temp))+" ("+str(temp)+'") minutos'
    elif numbstring[-1] == '”' or numbstring[-1] == '"' :
        temp = numbstring.replace('"',"").replace('”',"").replace(" ","").replace(",","")

        if int(temp) ==1:
            return "un ("+str(temp)+"') segundo"
        temp = numero_a_letras(float(temp))+" ("+str(temp)+'") segundos'
    elif re.search('[a-zA-Z]', numbstring):
        return "it is a letter"
    else:
        temp = numero_a_letras(float(numbstring.replace(" ","").replace(",","")))
    return temp

def parcelTOTXT(cordlist):
    strr = ""
    dictCords = {'NORTH':"Norte",'SOUTH':'Sur','EAST': 'Este','WEST':'Oeste'}
    for cord in cordlist:
        strr += dictCords[cord]+": "
        linds = cordlist[cord]
        cordlist2 = []
        for i in range(len(cordlist[cord])):
            
            lind = cordlist[cord][i].replace(".","").replace(".","").replace(":","")

            if lind[0:5].lower() == "calle":
                try:
                    if "-" in lind[5:]:

                        ind = lind.find("-")
                        try:
                            strr += "Calle "+numero_a_letras(int(lind[5:ind]))+lind[ind:]
                 
                        except:
                            strr += "Calle "+lind[5:ind+1]+numero_a_letras(int(lind[ind+1:]))
                        cordlist2.append(lind[5:])
                    else:
                        strr += "Calle "+numero_a_letras(int(lind[5:]))
                        cordlist2.append("Calle-"+lind[5:])
                except:
                    strr += lind


            elif lind[0:6].lower() == "parque":
                strr += "parque"
            elif "-" in lind:
                ind = lind.find("-")
                manz = lind[:ind]
                numb = int(lind[ind+1:])
                cordlist2.append(lind)
                
                strr += "lote "+manz+"-"+numero_a_letras(numb)
            else:
                listnumb = re.findall(r'\d+',lind)
                if len(listnumb) !=0:
                    for numb in listnumb:
                        lettenumb = numero_a_letras(numb).upper()
                        numb = str(numb)
                        leng = len(numb)
                        ind = lind.find(numb)
                        newstr = " "+lettenumb+" ("+numb+")"
                        lind = lind.replace(numb,newstr)
                if "|" in lind:
                    ind = lind.find("|")
                    lind = lind[:ind]
                else:
                    lind = lind


                strr += lind



            if i+1 != len(cordlist[cord]):
                strr += " y "
        
        if len(cordlist2) > 0:
            strr += " ("
            for i in range(len(cordlist2)):
                strr += cordlist2[i]

                if i+1 != len(cordlist2):
                    strr += " y "
            strr += ")"


        if cord != 'WEST':
            strr += "; "
        else:
            strr += ".---------------"

    return strr



def get_numbandword(numbstring):

    return check_content(numbstring)+" ("+str(numbstring)+")"


def MakeSQRmtrr(floa):
    floa = str(round(floa,2))
    ind = floa.find(".")
    if ind == -1:
        # cero decimals
        floa +=".00" 
    elif len(floa)-ind == 2:
        floa += "0"
    
 
    if len(floa) >=7 and len(floa)<=9:
        floa = floa[:-6] +","+floa[-6:]
    elif len(floa) >=10 and len(floa)<=12:
        floa = floa[:-9]+","+floa[-9:-6] +","+floa[-6:]
    elif len(floa) >=13 and len(floa)<=15:
        floa = floa[:-12]+","+floa[-12:-9]+","+floa[-9:-6] +","+floa[-6:]
    return floa