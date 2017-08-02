import openpyxl as xl
from Workbook import Workbook
from CASsheet import CASsheet

##################################################
#       class for CAS book handling
##################################################
class CASbook(Workbook):
    
    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,file_name = None):
        self._workbook = None
        if file_name != None:
            self._workbook = xl.load_workbook(file_name)
        super(CASbook,self).__init__(self._workbook)
        self.load_sheets(CASsheet,self._workbook.worksheets)
        self.load_sheets_name(self._workbook.sheetnames)
