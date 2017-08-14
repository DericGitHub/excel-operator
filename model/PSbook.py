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
        #self._workbook = xl.load_workbook(file_name)
        #self._current_sheet = None
        self.load_sheets(PSsheet,self._workbook.worksheets)
        self.load_sheets_name(self._workbook.sheetnames)
    def save_as(self,path_name):
        f = open(path_name,'wb')
        f.write(xl.writer.excel.save_virtual_workbook(self._workbook))
        f.close()
        #self._workbook.save(path_name)
    def virtual_workbook(self):
        return xl.writer.excel.save_virtual_workbook(self._workbook)
        


