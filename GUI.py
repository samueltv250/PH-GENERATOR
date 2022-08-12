from Main import *

import sys
import time
from PyQt5.QtWidgets import QLabel, QMainWindow, QPushButton, QApplication, QTextEdit,QLineEdit,QMessageBox
from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import QThread,QTimer
from createReglament import *
from createProtocol import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from threading import *
from PyQt5.QtWidgets import QCheckBox
from TableCreator import *



sttr = "No se encontraron ningunos errores."
PH = ""

def OnBtnClicked(textboxValue):
        """Runs the main function."""
        try:
            
            etapas = int(textboxValue)
        except Exception as e:
            errrers+=1
            print(e)
        errrers = 0
        print('------------------------------------------------------------------------------------------------')
        print('Iniciando generacion del Reglamento para el '+PH+" Etapa "+str(etapas)+'...')
        

        errrers = CreateRegFinal(etapas,errrers)
   
        # time.sleep(5) # time.sleep() is a blocking task that does not allow the Qt event loop to run, which prevents the signal from working properly and GUI updates. The solution is to replace the GUI sleep with QTimer and QEventLoop.
        global sttr
        sttr = "No se encontraron ningunos errores y se guardo el documento como 'Final_Reglamento-Etapa-"+str(int(etapas))+".docx'"
        if errrers != 0:
            sttr = "Error: Aparecieron "+str(errrers)+" Errores en la corrida."
        print('Fin: '+sttr)
        # msgBox = QMessageBox()

        # msgBox.setText(sttr)
        # msgBox.exec()
def OnBtnClickedprot(textboxValue):
        try:
            
            etapas = int(textboxValue)
        except Exception as e:
            errrers+=1
            print(e)
        """Runs the main function."""
        errrers = 0
        print('------------------------------------------------------------------------------------------------')
        print('Iniciando generacion del Protocolo para el '+PH+" Etapa "+str(etapas)+'...')
 
        
        errrers = CreateProtFinal(etapas,errrers)
        # time.sleep(5) # time.sleep() is a blocking task that does not allow the Qt event loop to run, which prevents the signal from working properly and GUI updates. The solution is to replace the GUI sleep with QTimer and QEventLoop.
        global sttr
        sttr = "No se encontraron ningunos errores y se guardo el documento como 'Final_Protocolo-Etapa-"+str(int(etapas))+".docx'"
        if errrers != 0:
            sttr = "Error: Aparecieron "+str(errrers)+" Errores en la corrida."
        print('Fin: '+sttr)
        # msgBox = QMessageBox()

        # msgBox.setText(sttr)
        # msgBox.exec()


def OnBtnClickedrequest():

        """Runs the main function."""
        errrers = 0
        print('------------------------------------------------------------------------------------------------')
        print('Iniciando generacion de solicitud de construccion...')
 
        
        errrers = fullReq()
        # time.sleep(5) # time.sleep() is a blocking task that does not allow the Qt event loop to run, which prevents the signal from working properly and GUI updates. The solution is to replace the GUI sleep with QTimer and QEventLoop.
        global sttr
        sttr = "No se encontraron ningunos errores y se guardo el documento como 'Request.xlsx'"
        if errrers != 0:
            sttr = "Error: Aparecieron "+str(errrers)+" Errores en la corrida."
        print('Fin: '+sttr)
        # msgBox = QMessageBox()

        # msgBox.setText(sttr)
        # msgBox.exec()

class Stream(QtCore.QObject):
    """Redirects console output to text widget."""
    newText = QtCore.pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))


class QMyWindow(QMainWindow):
    EXIT_CODE_REBOOT = -123
    """Main application window."""
    def __init__(self):
        super().__init__()

        self.initUI()

        # Note this sentence can be printed to the console, easy to debug
        sys.stdout = Stream(newText=self.onUpdateText)

        # Initialize a timer
        self.timer = QTimer(self)
        # Connect the timer timeout signal to the slot function showTime()
        self.timer.timeout.connect(self.fun)

        self.num = 0
    def thread(self):
        
        t1=Thread(target=self.Operation)
        t1.daemon = True
        t1.start()

    def threadRequest(self):
        
        t1=Thread(target=self.Operationreque)
        t1.daemon = True
        t1.start()
        
    def threadprot(self):
        
        t2=Thread(target=self.Operationprot)
        t2.daemon = True
        t2.start()
    def reset(self):
        # sys.exit(app.exec_())
        os.execl(sys.executable, sys.executable, *sys.argv)



    def Operation(self):
        self.btn.setEnabled(False)
        self.btn2.setEnabled(False)
        try:

            if self.textbox.text().isdigit():

                etapa =int(self.textbox.text())
            else:
                print('Error: Debe ingresar un numero.')


            OnBtnClicked(etapa)

        except:
            print('Error: Etapa no Valida, Ingrese un numero de etapa valido.')

        self.btn.setEnabled(True)
        self.btn2.setEnabled(True)

    def Operationreque(self):
     
        self.btn3.setEnabled(False)


        OnBtnClickedrequest()

    
        
        self.btn3.setEnabled(True)
        
    def Operationprot(self):
        try:
            self.btn2.setEnabled(False)
            self.btn.setEnabled(False)
            if self.textbox.text().isdigit():
    
                etapa =int(self.textbox.text())
            else:
                print('Error: Debe ingresar un numero.')
            OnBtnClickedprot(etapa)
            self.btn2.setEnabled(True)
            self.btn.setEnabled(True)
        except:
            print('Error Etapa no Valida, Ingrese un numero de etapa valido.')
            
        

        

    def fun(self):
        self.num += 1
        print(self.num)

    def onUpdateText(self, text):
        """Write console output to text widget."""
        cursor = self.process.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.process.setTextCursor(cursor)
        self.process.ensureCursorVisible()

    def closeEvent(self, event):
        """Shuts down application on close."""
        # Return stdout to defaults.
        sys.stdout = sys.__stdout__
        super().closeEvent(event)
    def clickBox(self, state):
        global TablaPrevia
        if state == QtCore.Qt.Checked:
            CheckTab(True)
        else:
            CheckTab(False)

    def initUI(self):
        """Creates UI window on launch."""
        # Button for generating the master list.
        global PH
        myFont2=QtGui.QFont()
        myFont2.setPixelSize(11)

        myFont2.setBold(False)
        self.b = QCheckBox("Usar Tabla Previa?",self)
        self.b.stateChanged.connect(self.clickBox)
        self.b.move(650,300)
        self.b.resize(320,40)
        PH = listOfDat()[0][0]
        self.label_2 = QLabel('Etapa:', self)
        self.label_2.setFont(myFont2)
    
        self.label_2.move(650, 100)

        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.resize(50, 30)


        self.label_1 = QLabel("GENERANDO: "+PH, self)
        self.label_1.setAlignment(QtCore.Qt.AlignCenter)
        # self.label_1.setStyleSheet("border: 1px solid black;")
        self.label_1.move(200, 35)


        self.label_1.resize(500, 50)
        myFont=QtGui.QFont()
        myFont.setPixelSize(22)

        myFont.setBold(True)
        self.label_1.setFont(myFont)


        self.btnr = QPushButton('Reset', self)
        self.btnr.move(830, 10)
        self.btnr.resize(60, 30)
        self.btnr.setStyleSheet("border: 1px solid black;")
        self.btnr.clicked.connect(self.reset)
    
        self.btn3 = QPushButton('Request', self)
        self.btn3.move(650, 250)
        self.btn3.resize(100, 30)
        self.btn3.setStyleSheet("border: 1px solid black;")
        self.btn3.clicked.connect(self.Operationreque)



        self.btn = QPushButton('Reglamento', self)
        self.btn.move(650, 150)
        self.btn.resize(100, 30)
        self.btn.setStyleSheet("border: 1px solid black;")
        self.btn.clicked.connect(self.thread)
    
        self.btn2 = QPushButton('Protocolo', self)
        self.btn2.move(650, 200)
        self.btn2.resize(100, 30)
        self.btn2.setStyleSheet("border: 1px solid black;")
        self.btn2.clicked.connect(self.threadprot)
        self.btn.setFont(myFont2)
        self.btn2.setFont(myFont2)
        self.btnr.setFont(myFont2)
        self.btn3.setFont(myFont2)
        
        self.b.setFont(myFont2)
        self.textbox = QLineEdit(self)
        # self.textbox.setStyleSheet("border: 1px solid black;")
        self.textbox.move(700, 100)
        self.textbox.resize(50, 30)
        # Create the text output widget.
        self.process = QTextEdit(self, readOnly=True)
        self.process.ensureCursorVisible()
        self.process.setLineWrapColumnOrWidth(590)
        self.process.setLineWrapMode(QTextEdit.FixedPixelWidth)
        self.process.setFixedWidth(600)
        self.process.setFixedHeight(450)
        self.process.move(30, 100)
        self.process.setFont(myFont2)
        # Set window size and title, then show the window.
        self.setGeometry(300, 300, 900, 600)
        self.setWindowTitle('PH-GENERATOR')
        self.show()
    
    


if __name__ =='__main__':
    # Run the application.
    app = QApplication(sys.argv)
    gui = QMyWindow()
    print("Ingrese la Etapa que desea generar y presione el boton para el documento que desea generar.")
    sys.exit(app.exec_())

