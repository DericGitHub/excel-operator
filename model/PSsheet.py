import openpyxl as xl
from Worksheet import Worksheet
from PyQt4.QtGui import *
from PyQt4.QtCore import *
##################################################
#       class for PS sheet handling
##################################################
class PSsheet(Worksheet):
    def __init__(self,sheet = None):
        super(PSsheet,self).__init__(sheet)
        if sheet != None:
            self.init_sheet()
        self.init_model()
        self.load_rows(sheet.rows)
        self.load_cols(sheet.columns)
        
    def init_sheet(self):
        self._row_max = self._worksheet.max_row
        self._col_max = self._worksheet.max_column
        self._xmlname_coordinate = self.search_by_value('xmlname')

    
    def init_model(self):
        self._preview_model = QStandardItemModel()
        self._preview_model.setColumnCount(3)
        self._preview_model.setHeaderData(0,Qt.Horizontal,'xmlname')
    def update_model(self):
        self.init_sheet()
        if self._xmlname_coordinate != None:
            for xml_name in self.xml_names():
                print xml_name
                item = QStandardItem(xml_name)
                self._preview_model.appendRow(item)
        return self._preview_model
    def search_by_value(self,value):
        for row in self.rows:
            for cell in row:
                if cell.value == value:
                    return {'row':cell.row,'col':cell.col_idx}
        return None
    def xml_names(self):
        def get_value(var):
            if var.value != None:
                return var.value
            else:
                return 'None'
        # reverse first
        cells = list(self._worksheet.iter_cols(min_col=self._xmlname_coordinate['col'],min_row=self._xmlname_coordinate['row']+1,max_col=self._xmlname_coordinate['col'],max_row=self.row_max).next())
        while cells[-1].value == None:
            cells.pop()
        return map(get_value,cells)
            
#        for cell in cells:
#            if cell.value != None:
#                yield cell.value
#            else:
#                yield 'None'
    
    @property
    def preview_model(self):
        return self._preview_model
#    @property
#    def xml_names(self):
#        return self.xml_names()
                    
