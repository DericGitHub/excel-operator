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
        self.bind_GUI_event()
        self.show_GUI()
    
    def init_GUI(self):
        self._application = QApplication(sys.argv)
        self._window = Window.Window()
    def bind_GUI_event(self):
        self._window.bind_open_cas(self.testfunc)
    def show_GUI(self):
        self._window.show()
        sys.exit(self._application.exec_())       
    def testfunc(self):
        print 'test'
        



