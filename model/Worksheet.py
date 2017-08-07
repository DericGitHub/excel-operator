import openpyxl as xl
from Workcell import *
from PyQt4.QtGui import *
from PyQt4.QtCore import *


##################################################
#       Abstract class to handle worksheet
##################################################
class Worksheet(object):

    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,sheet = None,xmlname_coordinate = None):
        self._worksheet = sheet
        self._rows = []
        self._min_row = 0
        self._max_row = None
        self._cols = []
        self._min_col = 0
        self._max_col = None
        self._xmlname_coordinate = xmlname_coordinate
        self._title = {}
        if sheet != None:
            self.init_sheet()
        self.init_model()

    def init_sheet(self):
        self._row_max = self._worksheet.max_row
        self._col_max = self._worksheet.max_column
        self._xmlname = self.search_xmlname_by_value()
        if self._xmlname != None:
            self._status = self.search_header_by_value(u'Status(POR,INIT,PREV)')
            self._subject_matter = self.search_header_by_value(u'Subject Matter/\nFunctional Area')
            self._container_name = self.search_header_by_value(u'Container Name\nTechnical Specification')
        self.load_rows(self._worksheet.rows)
        self.load_cols(self._worksheet.columns)

    
    def init_model(self):
        self._preview_model = QStandardItemModel()
        self._preview_model.setColumnCount(4)
        self._preview_model.setHeaderData(0,Qt.Horizontal,'status')
        self._preview_model.setHeaderData(1,Qt.Horizontal,'subject matter')
        self._preview_model.setHeaderData(2,Qt.Horizontal,'container name')
        self._preview_model.setHeaderData(3,Qt.Horizontal,'xmlname')
        self._header_model = QStandardItemModel()
        self._extended_preview_model = QStandardItemModel()

    ##################################################
    #       Model api
    ##################################################
    def update_model(self):
        self.init_model()
        self.init_sheet()
        if self._xmlname != None:
            for xml_name in self.xml_names():
                item_status = QPreviewItem(self._status.get_item_by_xmlname(xml_name))
                item_subject_matter = QPreviewItem(self._subject_matter.get_item_by_xmlname(xml_name))
                item_container_name = QPreviewItem(self._container_name.get_item_by_xmlname(xml_name))
                item_xml_name = QPreviewItem(xml_name)
                self._preview_model.appendRow((item_status,item_subject_matter,item_container_name,item_xml_name))

            for header in self.headers():
                item_header = QHeaderItem(header)
                item_header.setCheckState(Qt.Unchecked)
                item_header.setCheckable(True)
                self._header_model.appendRow(item_header)

            for row in self.rows:
                cell_list = []
                for cell in row:
                    cell_list.append(QPreviewItem(cell))
                self._extended_preview_model.appendRow(cell_list)
 
    ##################################################
    #       model data operation
    ##################################################
    def search_by_value(self,value):
        for row in self.rows:
            for cell in row:
                if cell.value == value:
                    return Workcell(cell)
        return None
    def search_header_by_value(self,value):
        for row in self.rows:
            for cell in row:
                if cell.value == value:
                    return Header(cell)
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
    def select_all_headers(self):
        for i in range(self._header_model.rowCount()):
            item = self._header_model.item(i)
            item.setCheckState(Qt.Checked)
    def unselect_all_headers(self):
        for i in range(self._header_model.rowCount()):
            item = self._header_model.item(i)
            item.setCheckState(Qt.Unchecked)
    ##################################################
    #       Worksheet property
    ##################################################
    @property
    def preview_model(self):
        return self._preview_model
    @property
    def header_model(self):
        return self._header_model
    @property
    def extended_preview_model(self):
        return self._extended_preview_model

    ##################################################
    #       Sync
    ##################################################
    def checked_headers(self):
        for i in range(self._header_model.rowCount()):
            item = self._header_model.item(i)
            if item.checkState() != 0:
                print "%s:%s"%(item.cell.value,item.checkState())

 
    def locate_xmlname(self):
        for row in self._rows:
            for cell in row:
                if cell.value == 'xmlname':
                    return cell.co
    def load_rows(self,rows):
        for row in rows:
            self._rows.append(row)

    def load_cols(self,cols):
        for col in cols:
            self._cols.append(col)
    def load_title(self):
        pass
        

        



    ##################################################
    #       User interface
    ##################################################
    def del_row(self,row_index):
        pass
    def append_row(self,row_index):
        pass
    def title(self):
        pass

    ##################################################
    #       Data
    ##################################################
    @property
    def worksheet(self):
        return self._worksheet
    @property
    def rows(self):
        return self._rows
    @property
    def cols(self):
        return self._cols
    @property
    def max_row(self):
        return self._max_row
    @property
    def min_row(self):
        return self._min_row
    @property
    def max_col(self):
        return self._max_col
    @property
    def min_col(self):
        return self._min_col
class QPreviewItem(QStandardItem):
    def __init__(self,cell):
        if cell.value != None:
            super(QPreviewItem,self).__init__(cell.value)
        else:
            super(QPreviewItem,self).__init__('')
        self._cell = cell
    
    @property
    def cell(self):
        return self._cell
class QHeaderItem(QStandardItem):
    def __init__(self,cell):
        if cell.value != None:
            super(QHeaderItem,self).__init__(cell.value)
        else:
            super(QHeaderItem,self).__init__('')
        self._cell = cell
    
    @property
    def cell(self):
        return self._cell
        
 
