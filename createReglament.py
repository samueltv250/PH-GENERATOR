
from WordWrite import *
from insertar_numatext import *
from extract_Alcance_PH import *
from docxcompose.composer import Composer
from docx import Document as Document_compose
from TableCreator import *

def CreateRegFinal(etapas,errrers):
    try:
        llist = listOfDat()
    except Exception as e:
        print("Error: leyendo DATA-GENERAL.xlsx")
        errrers +=1
        print(e)


    try:
        CreateReg(etapas)
        print("Exito: Se genero parte 1 del Reglamento")
    except Exception as e:
        print("Error: Generando parte 1 del Reglamento")
        errrers +=1
        print(e)

    try:

        partTwo(etapas,llist)
        print("Exito: Se genero parte 2 del Reglamento")
    except Exception as e:
        print("Error: Generando la tabla del Reglamento")
        errrers +=1
        print(e)


    try:

        partThree(etapas,llist)
        print("Exito: Se genero parte 3 del Reglamento")
        
    except Exception as e:
        print("Error: Generando parte 3 del Reglamento")
        errrers +=1
        print(e)
    
    try:

        createTabl(etapas)
        print("Exito: Se genero la tabla del Reglamento")

    except Exception as e:
        print("Error: Generando la tabla del Reglamento")
        errrers +=1
        print(e)

    try:
        master = Document_compose("Output_Docx/Parts/part1.docx")
        composer = Composer(master)
        doc2 = Document_compose("Output_Docx/Parts/part2.docx")
        doc3 = Document_compose("Output_Docx/Parts/part3.docx")
        Table = Document_compose("Output_Docx/Parts/table.docx")
        composer.append(doc2)
        composer.append(Table)
        composer.append(doc3)
        composer.save("Output_Docx/Final_Reglamento-Etapa-"+str(int(etapas))+".docx")
        print("Exito: Se genero el Reglamento final")
    except Exception as e:
        print("Error: Generando el Reglamento final")
        errrers +=1
        print(e)
    return errrers





