from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
import WindowUI
import sys
##################################################
#       class to handle the view
##################################################
class Window(QMainWindow):

    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,CASbook = None,PSbook = None):
        super(Window,self).__init__()
        self.init_Window()
        self.init_model(CASbook,PSbook)
    def init_Window(self):
        self.ui = WindowUI.Ui_MainWindow()
        self.ui.setupUi(self)
    def init_model(self,CASbook,PSbook):
        self.ui.name_cas.setText(CASbook)
        self.ui.name_ps.setText(PSbook)
    
    ##################################################
    #       Bind event
    ##################################################
    def bind_open_cas(self,func):
        self.ui.open_cas.clicked.connect(func)
    def bind_open_ps(self,func):
        self.ui.open_ps.clicked.connect(func)
       #self.ui.open_cas.clicked.connect(open_file_dialog)
    ##################################################
    #       Custom slot
    ##################################################
    def update_cas_file(self,filename):
        self.ui.name_cas.setText(filename)
    def update_ps_file(self,filename):
        self.ui.name_ps.setText(filename)
       

def open_file_dialog():
    filedialog = QFileDialog()
    fileName = filedialog.getOpenFileName()
    return fileName
