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
        
