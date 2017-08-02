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
        self.init_model()
        self.load_rows(sheet.rows)
        self.load_cols(sheet.columns)
        

    def init_model(self):
        self._preview_model = QStandardItemModel()
        self._preview_model.setColumnCount(3)
        self._preview_model.setHeaderData(0,Qt.Horizontal,'xmlname')
    def update_model(self):
        if self.locate_xmlname() != None:
            item = QStandardItem(self.locate_xmlname())
            self._preview_model.appendRow(item)
        return self._preview_model
    def locate_xmlname(self):
        for row in self.rows:
            for cell in row:
                if cell.value == 'xmlname':
                    return cell.coordinate
        return None
    def xml_names(self):
        #self._worksheet.get_squared_range()
        pass
    
    @property
    def preview_model(self):
        return self._preview_model
                    
