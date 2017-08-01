import sys
from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
from view import Window
##################################################
#       class for handling application logic
##################################################
class MainController(object):
    def __init__(self):
        self._application = None
        self._window = None
        self._PSbook = None
        self._CASbook = None
        self._PSstack = None
        self._CASstack = None
        self.init_GUI()
    
    def init_GUI(self):
        self._application = QApplication(sys.argv)
        self._window = Window.Window()
        self._window.show()
        sys.exit(self._application.exec_())
        



