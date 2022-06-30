import textract
from numero_letras import *



txt = textract.process("Pruerba_Remplaso.docx")
txt = str(txt).replace('(',"&(").replace(')',")&")
splittext = re.split('&', txt)

ddict = {}
for i in splittext:
    if i[0] == '(':
        strNumb = check_content(i[1:-1])
        if strNumb != "it is a letter":
            ddict[i] = strNumb
            
for p in ddict:
    print(p)
    print(ddict[p])
    
print(listofdatt)
        
        

        
        
        

print(len(tipomod))
print(dicdeff)
    

            


