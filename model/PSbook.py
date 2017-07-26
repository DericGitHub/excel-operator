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
    def __init__(self,workbook = None):

        super(PSbook,self).__init__(workbook)
        self.load_sheets(PSsheet,self._workbook.worksheets)
        self.load_sheet_names(self._workbook.sheetnames)
