from medidasExtractor import *
from docxcompose.composer import Composer
from docx import Document as Document_compose
from insertar_numatext import *
from TableCreator import *
def CreateProtFinal(etapas,errrers):
    try:
        llist = listOfDat()
    except Exception as e:
        print("Error: leyendo DATA-GENERAL.xlsx")
        errrers +=1
        print(e)
    try:

        createTabl(etapas)
        print("Exito: Se genero la tabla del Protocolo")

    except Exception as e:
        print("Error: Generando la tabla del Protocolo")
        errrers +=1
        print(e)

    try:
        WRITEPRot(etapas)
        print("Exito: Se genero la descripcion de todas las unidades inmobiliaria")
    except Exception as e:
        print("Error: Generando la descripcion de todas las unidades inmobiliaria")
        errrers +=1
        print(e)
    try:

        partTwo(etapas,llist)
        print("Exito: Se genero parte 2 del Protocolo")
    except Exception as e:
        print("Error: Generando la tabla del Protocolo")
        errrers +=1
        print(e)

    try:
        master = Document_compose("Output_Docx/Parts/prot-"+str(etapas)+".docx")
        composer = Composer(master)
        doc3to8 = Document_compose("Output_Docx/Parts/part2.docx")

        Table = Document_compose("Output_Docx/Parts/table.docx")
        composer.append(doc3to8)
        composer.append(Table)
    
        composer.save("Output_Docx/Final_Protocolo-Etapa-"+str(int(etapas))+".docx")
        print("Exito: Se genero el Protocolo final")
    except Exception as e:
        print("Error: Generando el Protocolo final")
        errrers +=1
        print(e)
    return errrers
    




