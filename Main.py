from createReglament import *
from createProtocol import *



def RunALL():
    try:
        llist = listOfDat()
        etapas = int(llist[0][5])
    except Exception as e:
        print(e)
    errrers = 0

    CreateRegFinal(etapas,errrers)
    CreateProtFinal(etapas,errrers)
# RunALL()