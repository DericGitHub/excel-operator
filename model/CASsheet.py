import openpyxl as xl
from Worksheet import Worksheet

##################################################
#       class for CAS sheet handling
##################################################
class CASsheet(Worksheet):
    def __init__(self,sheet = None,sheet_wr = None):
        super(CASsheet,self).__init__(sheet,sheet_wr)
        #self._worksheet_wr  = sheet_wr
        self.init_cas_sheet()
        self.init_cas_model()

    def init_cas_sheet(self):
        if self._xmlname != None:
            self._subject_matter = self.search_header_by_value(u'Subject Matter/\nFunctional Area')
            self._container_name = self.search_header_by_value(u'Container Name\nTechnical Specification')
    def init_cas_model(self):
        pass
    
    def cell(self,row,col):
        return self._worksheet_wr.range((row,col))
    
 
    
    
                    
