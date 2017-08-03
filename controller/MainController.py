import sys
from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
from view import Window
from model import *
##################################################
#       class for handling application logic
##################################################
class MainController(object):
    def __init__(self):
        self._application = None
        self._window = None
        self._PSbook = None
        self._PSbook_name = ''
        self._PSbook_sheets = QStandardItemModel()
        self._PSbook_current_sheet = None
        self._PSbook_current_sheet_name = None
        self._CASbook = None
        self._CASbook_name = ''
        self._CASbook_sheets = QStandardItemModel()
        self._CASbook_current_sheet = None
        self._CASbook_current_sheet_name = None
        self._PSstack = None
        self._CASstack = None
        self.init_GUI()
        self.bind_GUI_event()
        self.show_GUI()
        self.start_xlwings_app()
    
    def init_GUI(self):
        self._application = QApplication(sys.argv)
        self._window = Window.Window(self._CASbook_name,self._PSbook_name)
    def bind_GUI_event(self):
        self._window.bind_open_cas(self.open_cas)
        self._window.bind_open_ps(self.open_ps)
        self._window.bind_select_cas_sheet(self.select_cas_sheet)
        self._window.bind_select_ps_sheet(self.select_ps_sheet)
    def show_GUI(self):
        self._window.show()
        sys.exit(self._application.exec_())       
    def start_xlwings_app(self):
        pass
    def open_cas(self):
        try:
            filename = Window.open_file_dialog()
        except:
            filename = None
        self._CASbook = CASbook.CASbook(str(filename))
        self._CASbook_name = filename
        self._CASbook.update_model()
        self.refresh_cas_book_name(self._CASbook.workbook_name)
        self.refresh_cas_sheet_name(self._CASbook.sheet_name_model)
        #self._window.update_cas_file(self._CASbook_name)
        #self._window.update_cas_sheets(self._CASbook.sheets_name)
    def open_ps(self):
        try:
            filename = Window.open_file_dialog()
            print filename
        except:
            filename = None
        self._PSbook = PSbook.PSbook(str(filename))
        #try:
        #    self._PSbook = PSbook.PSbook(filename)
        #    print "open succeed"
        #except:
        #    print "not a excel"
        self._PSbook_name = filename
        self._PSbook.update_model()
        #self._window.update_ps_file(self._PSbook_name)
        #self._window.update_ps_sheets(self._PSbook.sheets_name)
        self.refresh_ps_book_name(self._PSbook.workbook_name)
        self.refresh_ps_sheet_name(self._PSbook.sheet_name_model)
    def select_cas_sheet(self,sheet_name):
        self._CASbook_current_sheet = self._CASbook.sheets_name[sheet_name]
        print 'cas_sheet :%s'%self._CASbook_current_sheet
    def select_ps_sheet(self,sheet_name):
        pass
        self._PSbook_current_sheet_name = self._PSbook.sheets_name[sheet_name]
        print self._PSbook_current_sheet_name
        self._PSbook_current_sheet = self._PSbook.sheets[self._PSbook_current_sheet_name]
        self._PSbook_current_sheet.update_model()
        #self.refresh_ps_sheet_name(self._PSbook_current_sheet.sheet_name_model)
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_ps_header(self._PSbook_current_sheet.ps_header_model)
        print 'ps_sheet :%s'%self._PSbook_current_sheet

    def refresh_cas_book_name(self,model):
        self._window.update_cas_file(model)
    def refresh_ps_book_name(self,model):
        self._window.update_ps_file(model)
    def refresh_cas_sheet_name(self,model):
        self._window.update_cas_sheets(model)
    def refresh_ps_sheet_name(self,model):
        self._window.update_ps_sheets(model)
    def refresh_preview(self,model):
        self._window.update_preview(model)
    def refresh_ps_header(self,model):
        self._window.update_ps_header(model)
