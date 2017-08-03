import openpyxl as xl
from Worksheet import Worksheet
from Workcell import *
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
        
    def init_sheet(self):
        self._row_max = self._worksheet.max_row
        self._col_max = self._worksheet.max_column
        self._xmlname = self.search_xmlname_by_value()
        if self._xmlname != None:
            self._subject_matter = self.search_by_value('xmlname')
            self._container_name = self.search_by_value('xmlname')
        self.load_rows(self._worksheet.rows)
        self.load_cols(self._worksheet.columns)

    
    def init_model(self):
        self._preview_model = QStandardItemModel()
        self._preview_model.setColumnCount(3)
        self._preview_model.setHeaderData(0,Qt.Horizontal,'subject matter')
        self._preview_model.setHeaderData(1,Qt.Horizontal,'container name')
        self._preview_model.setHeaderData(2,Qt.Horizontal,'xmlname')
        self._ps_header_model = QStandardItemModel()
        
    def update_model(self):
        self.init_model()
        self.init_sheet()
        if self._xmlname != None:
            for xml_name in self.xml_names():
                item_subject_matter = QStandardItem(xml_name.value)
                item_container_name = QStandardItem(xml_name.value)
                item_xml_name = QStandardItem(xml_name.value)
                self._preview_model.appendRow((item_subject_matter,item_container_name,item_xml_name))
            for header in self.headers():
                item_header = QStandardItem(header.value)
                item_header.setCheckState(False)
                item_header.setCheckable(True)
                self._ps_header_model.appendRow(item_header)
        return self._preview_model
    def search_by_value(self,value):
        for row in self.rows:
            for cell in row:
                if cell.value == value:
                    return Workcell(cell)
        return None
    def search_xmlname_by_value(self):
        for row in self.rows:
            for cell in row:
                if cell.value == 'xmlname':
                    return XmlName(cell)
        return None
    def xml_names(self):
        cells = list(self._worksheet.iter_cols(min_col=self._xmlname.col,min_row=self._xmlname.row+1,max_col=self._xmlname.col,max_row=self.max_row).next())
        while cells[-1].value == None:
            cells.pop()
        return map(lambda x:XmlName(x),cells)
    def headers(self):
        cells = list(self._worksheet.iter_rows(min_col=self.min_col,min_row=self._xmlname.row,max_col=self.max_col,max_row=self._xmlname.row).next())
        while cells[-1].value == None:
            cells.pop()
        while cells[0].value == None:
            cells.pop(0)
        return map(lambda x:Header(x),cells)
    def subject_matters_by_xml_name(self,xml_name):
     #   self._worksheet.cell(row = xml_name.row,col = self._subject_matter_coordinate
        pass
    def container_names_by_xml_name(self,xml_name):
        pass
            
    
    @property
    def preview_model(self):
        return self._preview_model
    @property
    def ps_header_model(self):
        return self._ps_header_model



