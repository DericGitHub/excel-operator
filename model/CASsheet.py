import openpyxl as xl
from Worksheet import Worksheet

##################################################
#       class for CAS sheet handling
##################################################
class CASsheet(Worksheet):
    def __init__(self,sheet = None):
        super(CASsheet,self).__init__(sheet)
        self.load_rows(sheet.rows)
        self.load_cols(sheet.columns)
    
    def locate_xmlname(self):
        for row in self.rows:
            for cell in row:
                if cell.value == 'xmlname':
                    return cell.coordinate
        return None
    
                    
