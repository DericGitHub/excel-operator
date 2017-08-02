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
        self._PSbook_name = ''
        self._CASbook = None
        self._CASbook_name = ''
        self._PSstack = None
        self._CASstack = None
        self.init_GUI()
        self.bind_GUI_event()
        self.show_GUI()
    
    def init_GUI(self):
        self._application = QApplication(sys.argv)
        self._window = Window.Window(self._CASbook_name,self._PSbook_name)
    def bind_GUI_event(self):
        self._window.bind_open_cas(self.open_cas)
        self._window.bind_open_ps(self.open_ps)
    def show_GUI(self):
        self._window.show()
        sys.exit(self._application.exec_())       
    def open_cas(self):
        try:
            filename = Window.open_file_dialog()
        except:
            filename = None
        self._CASbook_name = filename
        self._window.update_cas_file(self._CASbook_name)
    def open_ps(self):
        try:
            filename = Window.open_file_dialog()
        except:
            filename = None
        self._PSbook_name = filename
        self._window.update_ps_file(self._PSbook_name)


