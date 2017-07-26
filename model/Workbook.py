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
        self._sheet_names = []

    def load_sheets(self,sheet_cls,sheets):
        sheet_cnt = 0
        for sheet in sheets:
            self._sheets[sheet.title] = sheet_cls(sheet)
            sheet_cnt += 1
        self._sheets_cnt = sheet_cnt

    def load_sheet_names(self,sheets):
        for sheet in sheets:
            self._sheet_names.append(sheet)

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
        return self._sheet_names
