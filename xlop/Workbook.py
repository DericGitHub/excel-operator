import openpyxl as xl
import xlwings as xw
from Worksheet import Worksheet

def Workbook(object):
    def __init__(self,file_path):
        self._workbook = xl.load_workbook(file_path)
        self._sheets = []
        self._sheets_cnt = self.load_sheets()
        self._sheet_names = []
    
        pass
    

    def load_sheets(self):
        sheet_cnt = 0
        for sheet in self._workbook.sheets:
            self._sheets.append(Worksheet(sheet))
            cnt += 1
        return sheet_cnt

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
