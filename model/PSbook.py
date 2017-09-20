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
    def __init__(self,file_name = None,app = None):
        super(PSbook,self).__init__(file_name,app)
        if file_name != None:
            self.init_ps_book()
    def init_ps_book(self):
        self.load_sheets(PSsheet,self._workbook.worksheets,self._workbook_wr.sheets)
        


