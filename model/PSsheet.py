import openpyxl as xl
from Worksheet import Worksheet

##################################################
#       class for PS sheet handling
##################################################
class PSsheet(Worksheet):
    def __init__(self,sheet):
        super(PSsheet,self).__init__(sheet)
        self.load_rows(sheet.rows)
        self.load_cols(sheet.columns)
    
