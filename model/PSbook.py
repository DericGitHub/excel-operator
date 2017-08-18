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
        super(PSbook,self).__init__(file_name)
        if file_name != None:
            self.init_ps_book()
    def init_ps_book(self):
        self.load_sheets(PSsheet,self._workbook.sheets)
    def save_as(self,path_name):
        self._workbook.save(path_name)
    @property
    def virtual_workbook(self):
        return xl.writer.excel.save_virtual_workbook(self._workbook)
        


