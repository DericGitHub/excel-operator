import xlwings as xw
from Workbook import Workbook
from CASsheet import CASsheet

##################################################
#       class for CAS book handling
##################################################
class CASbook(Workbook):
    
    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,file_name = None,app = None):
        super(CASbook,self).__init__(file_name,app)
        if file_name != None:
            self.init_cas_book()
    def init_cas_book(self):
        self.load_sheets(CASsheet,self._workbook.worksheets,self._workbook_wr.sheets)
