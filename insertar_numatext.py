from docxtpl import DocxTemplate
from extract_Alcance_PH import *
from numero_letras import *
from WordWrite import *

def partTwo(faSe,listofdatt):
    if faSe ==1:
        docu = DocxTemplate("Input_Docx/Template_A3_A8.docx")
    else:
        docu = DocxTemplate("Input_Docx/Template_A8.docx")
    tot = getTot()

    context = {"PH" : str(listofdatt[0][0]),"TotalNumb" : get_numbandword(str(tot))}

    docu.render(context)
    docu.save("Output_Docx/Parts/part2.docx")

def partThree(faSe,listofdatt):
    if faSe ==1:
        docu = DocxTemplate("Input_Docx/Template_A9_A71.docx")
    else:
        docu = DocxTemplate("Input_Docx/Template_FIN.docx")
    dictEnum = {1:"Primera",2:"Segunda",3:"Tercera",4:"Cuarta",5:"Quinta"}
    context = {"PH" : str(listofdatt[0][0]),"Etapa" : dictEnum[faSe]}

    docu.render(context)
    docu.save("Output_Docx/Parts/part3.docx")



























#import textract
##from numero_letras import *
##
##
##
##txt = textract.process("Pruerba_Remplaso.docx")
##txt = str(txt).replace('(',"&(").replace(')',")&")
##splittext = re.split('&', txt)
##
##ddict = {}
##for i in splittext:
##    if i[0] == '(':
##        strNumb = check_content(i[1:-1])
##        if strNumb != "it is a letter":
##            ddict[i] = strNumb
##
##for p in ddict:
##    print(p)
##    print(ddict[p])
##
##print(listofdatt)
##
##
##
##
##
##
##
##print(len(tipomod))
##print(dicdeff)
##

            


