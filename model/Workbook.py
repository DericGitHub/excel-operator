import openpyxl as xl

from Worksheet import Worksheet

##################################################
#       abstract class to handle workbook
##################################################
class Workbook(object):

    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,workbook = None):
        self._workbook = workbook
        self._sheets = {}
        self._sheets_cnt = None
        self._sheets_name = []
        if self._workbook != None:
            self.init_book()
        self.init_model()

    def init_book(self):
        self._workbook = xl.load_workbook(file_name)
        self._current_sheet = None
    def init_model(self):
        self._book_name = None
        self._sheet_name_model = QStandardItemModel()

    def update_book_name(self,name):
        self._book_name = name
    def update_sheet_name_model(self):
        self._sheet_name_model.clear()
        for sheet_name in self._sheets_name:
            item_sheet_name = QStandardItem(sheets_name)
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

    def init_book(self):
        self._workbook = xl.load_workbook(file_name)
        self._current_sheet = None
    def init_model(self):
        self._book_name = None
        self._sheet_name_model = QStandardItemModel()

    def update_book_name(self,name):
        self._book_name = name
    def update_sheet_name_model(self,model):
        self._sheet_name_model.clear()
        for sheet in 
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
