import openpyxl as xl

from Worksheet import Worksheet
from PyQt4.QtGui import *
from PyQt4.QtCore import *

##################################################
#       abstract class to handle workbook
##################################################
class Workbook(object):

    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,workbook = None):
        if workbook != None:
            self.init_book(workbook)
        self.init_model()

    def init_book(self,workbook):
        self._workbook_name = workbook
        self._sheets = {}
        self._sheets_cnt = None
        self._sheets_name = []
        self._workbook = xl.load_workbook(self._workbook_name)
        self._current_sheet = None
    def init_model(self):
        self._book_name = None
        self._sheet_name_model = QStandardItemModel()
    def update_model(self):
        self.update_workbook_name()
        self.update_sheet_name_model()
    def update_workbook_name(self):
        self._workbook_name = self._workbook_name
    def update_sheet_name_model(self):
        self._sheet_name_model.clear()
        for sheet_name in self._sheets_name:
            item_sheet_name = QStandardItem(sheet_name)
            self._sheet_name_model.appendRow(item_sheet_name)
            

    def load_sheets(self,sheet_cls,sheets):
        sheet_cnt = 0
        for sheet in sheets:
            self._sheets[sheet.title] = sheet_cls(sheet)
            sheet_cnt += 1
        self._sheets_cnt = sheet_cnt

    def load_sheets_name(self,sheets):
        for sheet in sheets:
            self._sheets_name.append(sheet)

    ##################################################
    #       user interface
    ##################################################
    def open(self,path_name):
        pass
    def save_as(self,path_name):
        pass
    def save(self):
        pass
    def recover(self,direction):
        pass

    ##################################################
    #       data 
    ##################################################
    @property
    def sheets(self):
        return self._sheets
    @property
    def sheets_name(self):
        return self._sheets_name
    @property
    def sheet_name_model(self):
        return self._sheet_name_model
    @property
    def workbook_name(self):
        return self._workbook_name
