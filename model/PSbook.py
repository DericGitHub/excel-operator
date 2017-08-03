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
        if file_name != None:
            self.init_book()
        super(PSbook,self).__init__(self._workbook)
        self.load_sheets(PSsheet,self._workbook.worksheets)
        self.load_sheets_name(self._workbook.sheetnames)
    def init_book(self):
        self._workbook = xl.load_workbook(file_name)
        self._current_sheet = None
        self.load_sheets(PSsheet,self._workbook.worksheets)
        self.load_sheets_name(self._workbook.sheetnames)
    def init_model(self):
        self._sheet_name_model = QStandardItemModel()


