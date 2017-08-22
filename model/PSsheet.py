import openpyxl as xl
from Worksheet import Worksheet,QPreviewItem
from Workcell import *
from PyQt4.QtGui import *
from PyQt4.QtCore import *
from copy import copy
import time
##################################################
#       class for PS sheet handling
##################################################
UP = 0
DOWN = 1
nn = 0
def pt(step):
    global nn
    print 'step %s:takes %s'%(step,time.time()-nn)
    nn = time.time()
class PSsheet(Worksheet):
    def __init__(self,sheet = None,sheet_wr = None):
        super(PSsheet,self).__init__(sheet,sheet_wr)
        self.init_ps_sheet()
        self.init_ps_model()
    
    def init_ps_sheet(self):
        if self._xmlname != None:
            self._status = self.search_header_by_value(u'Status(POR,INIT,PREV)')
            self._subject_matter = self.search_header_by_value(u'Subject Matter/\nFunctional Area')
            self._container_name = self.search_header_by_value(u'Container Name\nTechnical Specification')
    def init_ps_model(self):
        self._preview_model = QStandardItemModel()
        self._preview_model.setColumnCount(4)
        self._preview_model.setHeaderData(0,Qt.Horizontal,'status')
        self._preview_model.setHeaderData(1,Qt.Horizontal,'subject matter')
        self._preview_model.setHeaderData(2,Qt.Horizontal,'container name')
        self._preview_model.setHeaderData(3,Qt.Horizontal,'xmlname')
        self._extended_preview_model = QStandardItemModel()
    def update_model(self):
        pt(1)
        super(PSsheet,self).update_model()
        pt(2)
        self.init_ps_model()
        pt(3)
        self.init_ps_sheet()
        pt(4)
#        column1 = self._status.column
#        column2 = self._subject_matter.column
#        column3 = self._container_name.column
        if self._xmlname != None:
            for xml_name in self.xml_names():
#                row = xml_name.row
#                item_status = QPreviewItem(self._worksheet.range(row,column1))
#                item_subject_matter = QPreviewItem(self._worksheet.range(row,column2))
#                item_container_name = QPreviewItem(self._worksheet.range(row,column3))
                item_status = QPreviewItem(self._status.get_item_by_xmlname(xml_name))
                item_subject_matter = QPreviewItem(self._subject_matter.get_item_by_xmlname(xml_name))
                item_container_name = QPreviewItem(self._container_name.get_item_by_xmlname(xml_name))
                #item_coordinate = QStandardItem('row:%s,col:%s'%(xml_name.row,xml_name.col))
                item_xml_name = QPreviewItem(xml_name)
                self._preview_model.appendRow((item_status,item_subject_matter,item_container_name,item_xml_name))
                #self._preview_model.appendRow((item_status,item_subject_matter,item_coordinate,item_xml_name))
            #for row in self.rows:
            #    cell_list = []
            #    for cell in row:
            #        cell_list.append(QPreviewItem(cell))
            #    self._extended_preview_model.appendRow(cell_list)
        pt(5)
    def status(self):
        cells = list(self._worksheet.iter_cols(min_col=self._status.col,min_row=self._status.row+1,max_col=self._status.col,max_row=self.max_row).next())
        while cells[-1].value == None:
            cells.pop()
        return map(lambda x:Status(x,self._worksheet_wr),cells)
    def cell(self,row,col):
        #return self._worksheet_wr.range(row,col)
        return self._worksheet.cell(row=row,column=col)
    
    def add_row(self,start_pos,offset,orientation):
        loop = offset
        while loop > 0:
            self._worksheet_wr.api.Rows[start_pos].Insert(-4121)
            loop -= 1
#    def add_row(self,start_pos,offset,orientation):
#        rows = []
#        if orientation == 0:
#            real_start_pos = start_pos
#        elif orientation == 1:
#            real_start_pos = start_pos + 1
#        for row in self._worksheet.iter_rows(min_row=real_start_pos,max_row=self._worksheet.max_row):
#            rows.append(row)
#        #print(rows)
#        for row in reversed(rows):
#            row = reversed(row)
#            adjust_row_height = True
#            for cell in row:
#                #print("%s cell.row = %d,offset = %d"%(cell,cell.row,offset))
#                offset_cell = cell.parent.cell(row=cell.row+offset,column=cell.col_idx)
#                #print("%s cell.row = %d,offset = %d"%(offset_cell,offset_cell.row,offset))
#                if adjust_row_height == True:
#                    self._worksheet.row_dimensions[offset_cell.row] = self._worksheet.row_dimensions[cell.row]
#                    adjust_row_height = False
#                offset_cell.value = cell.value
#                #offset_cell.font = cell.font.copy()
#                #offset_cell.border = cell.border.copy()
#                #offset_cell.fill = cell.fill.copy()
#                #offset_cell.number_format = cell.number_format
#                offset_cell.protection = cell.protection.copy()
#                #offset_cell.alignment = cell.alignment.copy()
#        for row in self._worksheet.iter_rows(min_row=real_start_pos,max_row=start_pos+offset):
#            adjust_row_height = True
#            for cell in row:
#                if adjust_row_height == True:
#                    self._worksheet.row_dimensions[cell.row].ht = None
#                    adjust_row_height = False
#                cell.value = None



    def delete_row(self,start_pos,offset):
        self._worksheet_wr.range('%d:%d'%(start_pos,start_pos+offset-1)).api.Delete()
#    def delete_row(self,start_pos,offset):
#        rows = []
#        for row in self._worksheet.iter_rows(min_row=start_pos+offset,max_row=self._worksheet.max_row):
#            for cell in row:
#                offset_cell = cell.parent.cell(row=cell.row-offset,column=cell.col_idx)
#                offset_cell.value = cell.value
#                offset_cell.font = cell.font.copy()
#                offset_cell.border = cell.border.copy()
#                offset_cell.fill = cell.fill.copy()
#                offset_cell.number_format = cell.number_format
#                offset_cell.protection = cell.protection.copy()
#                offset_cell.alignment = cell.alignment.copy()
#        for row in self._worksheet.iter_rows(min_row=self._worksheet.max_row-offset+1,max_row=self._worksheet.max_row):
#            adjust_row_height = True
#            for cell in row:
#                if adjust_row_height == True:
#                    self._worksheet.row_dimensions[cell.row].ht = None
#                    adjust_row_height = False
#                cell.value = None
    def lock_row(self,row,status):
        self._worksheet_wr.api.Rows[row].Style.Locked = True
#    def lock_row(self,row,status):
#        for row in self._worksheet.iter_rows(min_row=row,max_row=row,min_col=self._worksheet.min_column,max_col=self._worksheet.max_column):
#            for cell in row:
#                new_protection = copy(cell.protection)
#                new_protection.locked = status
#                cell.protection = new_protection
    def lock_sheet(self,status):
        if status == True or status == False:
            self._worksheet_wr.api.Protect(status)
#    def lock_sheet(self,status):
#        new_protection = copy(self._worksheet.protection)
#        new_protection.sheet = status
#        new_protection.objects = status
#        new_protection.scenarios = status
#        self._worksheet.protection = new_protection
    def unlock_all_cells(self):
        self._worksheet_wr.api.Cells.Style.Locked = False

#    def unlock_all_cells(self):
#        for cell in self._worksheet.get_cell_collection():
#            new_protection = copy(cell.protection)
#            new_protection.locked = False
#            cell.protection = new_protection
            
 
#        if sheet != None:
#            self.init_sheet()
#        self.init_model()
#        
#    def init_sheet(self):
#        self._row_max = self._worksheet.max_row
#        self._col_max = self._worksheet.max_column
#        self._xmlname = self.search_xmlname_by_value()
#        if self._xmlname != None:
#            self._status = self.search_header_by_value(u'Status(POR,INIT,PREV)')
#            self._subject_matter = self.search_header_by_value(u'Subject Matter/\nFunctional Area')
#            self._container_name = self.search_header_by_value(u'Container Name\nTechnical Specification')
#        self.load_rows(self._worksheet.rows)
#        self.load_cols(self._worksheet.columns)
#
#    
#    def init_model(self):
#        self._preview_model = QStandardItemModel()
#        self._preview_model.setColumnCount(4)
#        self._preview_model.setHeaderData(0,Qt.Horizontal,'status')
#        self._preview_model.setHeaderData(1,Qt.Horizontal,'subject matter')
#        self._preview_model.setHeaderData(2,Qt.Horizontal,'container name')
#        self._preview_model.setHeaderData(3,Qt.Horizontal,'xmlname')
#        self._ps_header_model = QStandardItemModel()
#        self._extended_preview_model = QStandardItemModel()
#        
#    def update_model(self):
#        self.init_model()
#        self.init_sheet()
#        if self._xmlname != None:
#            for xml_name in self.xml_names():
#                item_status = QPreviewItem(self._status.get_item_by_xmlname(xml_name))
#                item_subject_matter = QPreviewItem(self._subject_matter.get_item_by_xmlname(xml_name))
#                item_container_name = QPreviewItem(self._container_name.get_item_by_xmlname(xml_name))
#                item_xml_name = QPreviewItem(xml_name)
#                self._preview_model.appendRow((item_status,item_subject_matter,item_container_name,item_xml_name))
#            for header in self.headers():
#                item_header = QStandardItem(header.value)
#                item_header.setCheckState(False)
#                item_header.setCheckable(True)
#                self._ps_header_model.appendRow(item_header)
#            for row in self.rows:
#                cell_list = []
#                for cell in row:
#                    cell_list.append(QPreviewItem(cell))
#                self._extended_preview_model.appendRow(cell_list)
#                    
#        #return self._preview_model
#    def search_by_value(self,value):
#        for row in self.rows:
#            for cell in row:
#                if cell.value == value:
#                    return Workcell(cell)
#        return None
#    def search_header_by_value(self,value):
#        for row in self.rows:
#            for cell in row:
#                if cell.value == value:
#                    return Header(cell)
#        return None
#    def search_xmlname_by_value(self):
#        for row in self.rows:
#            for cell in row:
#                if cell.value == 'xmlname':
#                    return XmlName(cell)
#        return None
#    def xml_names(self):
#        cells = list(self._worksheet.iter_cols(min_col=self._xmlname.col,min_row=self._xmlname.row+1,max_col=self._xmlname.col,max_row=self.max_row).next())
#        while cells[-1].value == None:
#            cells.pop()
#        return map(lambda x:XmlName(x),cells)
#    def headers(self):
#        cells = list(self._worksheet.iter_rows(min_col=self.min_col,min_row=self._xmlname.row,max_col=self.max_col,max_row=self._xmlname.row).next())
#        while cells[-1].value == None:
#            cells.pop()
#        while cells[0].value == None:
#            cells.pop(0)
#        return map(lambda x:Header(x),cells)
#    def subject_matters_by_xml_name(self,xml_name):
#     #   self._worksheet.cell(row = xml_name.row,col = self._subject_matter_coordinate
#        pass
#    def container_names_by_xml_name(self,xml_name):
#        pass
#            
#    
#    @property
#    def preview_model(self):
#        return self._preview_model
#    @property
#    def ps_header_model(self):
#        return self._ps_header_model
#    @property
#    def extended_preview_model(self):
#        return self._extended_preview_model



#class QPreviewItem(QStandardItem):
#    def __init__(self,cell):
#        if cell.value != None:
#            super(QPreviewItem,self).__init__(cell.value)
#        else:
#            super(QPreviewItem,self).__init__('')
#        self._cell = cell
#    
#    @property
#    def cell(self):
#        return self._cell
        
        
