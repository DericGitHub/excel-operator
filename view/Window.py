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
    def __init__(self):
        super(Window,self).__init__()
        self.init_Window()
    def init_Window(self):
        self.ui = WindowUI.Ui_MainWindow()
        self.ui.setupUi(self)
    
    ##################################################
    #       Bind event
    ##################################################
    def bind_open_cas(self,func):
        self.ui.open_cas.clicked.connect(func)
        self.ui.open_cas.clicked.connect(self.open_file_dialog)

    def open_file_dialog(self):
        filedialog = QFileDialog()
        #filedialog.show()
        fileName = filedialog.getOpenFileName()
        print fileName
