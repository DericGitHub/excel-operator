from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
import WindowUI
import sys
from model import *
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

        #       sheets_cas_model
        self._sheets_cas_model = QStandardItemModel()
        self.ui.sheets_cas.setModel(self._sheets_cas_model)

        #       sheets_ps_model
        self._sheets_ps_model = QStandardItemModel()
        self.ui.sheets_ps.setModel(self._sheets_ps_model)

        #       preview_model
        self._preview_model = QStandardItemModel()
        self.ui.preview.setModel(self._preview_model)
    
    ##################################################
    #       Bind event
    ##################################################
    def bind_open_cas(self,func):
        self.ui.open_cas.clicked.connect(func)
    def bind_open_ps(self,func):
        self.ui.open_ps.clicked.connect(func)
    def bind_select_cas_sheet(self,func):
        self.ui.sheets_cas.currentIndexChanged.connect(func)
    def bind_select_ps_sheet(self,func):
        #self.ui.sheets_ps.currentIndexChanged.connect(func)
        self.ui.sheets_ps.currentIndexChanged.connect(func)
    ##################################################
    #       Custom slot
    ##################################################
    def update_cas_file(self,filename):
        self.ui.name_cas.setText(filename)
    def update_ps_file(self,filename):
        self.ui.name_ps.setText(filename)
    def update_cas_sheets(self,sheets_name):
        self._sheets_cas_model.clear()
        for sheet_name in sheets_name:
            item = QStandardItem(sheet_name)
            self._sheets_cas_model.appendRow(item)
    def update_ps_sheets(self,sheets_name):
        self._sheets_ps_model.clear()
        for sheet_name in sheets_name:
            item = QStandardItem(sheet_name)
            self._sheets_ps_model.appendRow(item)
    def update_preview(self):
        for i in range(3):
            items =[]
            for j in range(50):
                item = QStandardItem(str(j))
                if i == 0:
                    item.setCheckState(False)
                    item.setCheckable(True)
                items.append(item)
            
            self._preview_model.appendColumn(items)
        print self._preview_model
      

def open_file_dialog():
    filedialog = QFileDialog()
    fileName = filedialog.getOpenFileName()
    return fileName
