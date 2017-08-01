import openpyxl as xl
from Workbook import Workbook
from PSsheet import PSsheet

##################################################
#       class for PS book handling
##################################################
class PSbook(Workbook):
    
    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,file_name = None):
        self._workbook = None
        if file_name != None:
            self._workbook = xl.load_workbook(file_name)
        super(PSbook,self).__init__(self._workbook)
        self.load_sheets(PSsheet,self._workbook.worksheets)
        self.load_sheet_names(self._workbook.sheetnames)
