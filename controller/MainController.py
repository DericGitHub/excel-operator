# -*- coding: utf-8 -*-
import sys
from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
from view import Window
from model import *
from copy import copy
import xlwings as xw
##################################################
#       class for handling application logic
##################################################
class MainController(object):
    ##################################################
    #       Initial method
    ##################################################
    def __init__(self):
        self._application = None
        self._xw_app = None
        self._window = None
        self._PSbook = None
        self._PSbook_name = ''
        self._PSbook_sheets = QStandardItemModel()
        self._PSbook_current_sheet = None
        self._PSbook_current_sheet_name = None
        self._CASbook = None
        self._CASbook_wr = None
        self._CASbook_name = ''
        self._CASbook_sheets = QStandardItemModel()
        self._CASbook_current_sheet = None
        self._CASbook_current_sheet_name = None
        self._PSstack = None
        self._CASstack = None
        self.start_xlwings_app()
        self.init_GUI()
        self.bind_GUI_event()
        self.show_GUI()
    
    def init_GUI(self):
        self._application = QApplication(sys.argv)
        self._window = Window.Window(self._CASbook_name,self._PSbook_name)
    def bind_GUI_event(self):
        self._window.bind_open_cas(self.open_cas)
        self._window.bind_open_ps(self.open_ps)
        self._window.bind_save_cas(self.save_cas)
        self._window.bind_save_ps(self.save_ps)
        self._window.bind_saveas_cas(self.saveas_cas)
        self._window.bind_saveas_ps(self.saveas_ps)
        self._window.bind_select_cas_sheet(self.select_cas_sheet)
        self._window.bind_select_ps_sheet(self.select_ps_sheet)
        self._window.bind_select_preview(self.select_preview)
        self._window.bind_sync_ps_to_cas(self.select_sync_ps_to_cas)
        self._window.bind_sync_cas_to_ps(self.select_sync_cas_to_ps)
        self._window.bind_sync_select_all_ps_headers(self.select_sync_select_all_ps_headers)
        self._window.bind_sync_select_all_cas_headers(self.select_sync_select_all_cas_headers)
    def show_GUI(self):
        self._window.show()
        sys.exit(self._application.exec_())       
    def start_xlwings_app(self):
        self._xw_app = xw.App()
    ##################################################
    #       Actions
    ##################################################
    def open_cas(self):
        try:
            filename = Window.open_file_dialog()
        except:
            filename = None
        self._CASbook = CASbook.CASbook(str(filename),self._xw_app)
        self._CASbook_wr = self._CASbook.workbook_wr
        self._CASbook_name = filename
        self._CASbook.update_model()
        self.refresh_cas_book_name(self._CASbook.workbook_name)
        self.refresh_cas_sheet_name(self._CASbook.sheet_name_model)
        #self._window.update_cas_file(self._CASbook_name)
        #self._window.update_cas_sheets(self._CASbook.sheets_name)
    def open_ps(self):
        try:
            filename = Window.open_file_dialog()
            print filename
        except:
            filename = None
        self._PSbook = PSbook.PSbook(str(filename))
        #try:
        #    self._PSbook = PSbook.PSbook(filename)
        #    print "open succeed"
        #except:
        #    print "not a excel"
        self._PSbook_name = filename
        self._PSbook.update_model()
        #self._window.update_ps_file(self._PSbook_name)
        #self._window.update_ps_sheets(self._PSbook.sheets_name)
        self.refresh_ps_book_name(self._PSbook.workbook_name)
        self.refresh_ps_sheet_name(self._PSbook.sheet_name_model)
    def save_cas(self):
        pass
    def save_ps(self):
        pass
    def saveas_cas(self):
        pass
    def saveas_ps(self):
        fileName = Window.save_file_dialog()
        self._PSbook.save_as(str(fileName))

    def select_cas_sheet(self,sheet_name):
        self._CASbook_current_sheet_name = self._CASbook.sheets_name[sheet_name]
        print self._CASbook_current_sheet_name
        self._CASbook_current_sheet = self._CASbook.sheets[self._CASbook_current_sheet_name]
        self._CASbook_current_sheet.update_model()
        self.refresh_cas_header(self._CASbook_current_sheet.header_model)
        print 'cas_sheet :%s'%self._CASbook_current_sheet
    def select_ps_sheet(self,sheet_name):
        self._PSbook_current_sheet_name = self._PSbook.sheets_name[sheet_name]
        print self._PSbook_current_sheet_name
        self._PSbook_current_sheet = self._PSbook.sheets[self._PSbook_current_sheet_name]
        self._PSbook_current_sheet.update_model()
        #self.refresh_ps_sheet_name(self._PSbook_current_sheet.sheet_name_model)
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_ps_header(self._PSbook_current_sheet.header_model)
        print 'ps_sheet :%s'%self._PSbook_current_sheet
    def select_preview(self,index):
        print self._PSbook_current_sheet._preview_model.itemFromIndex(index).cell.value
        self.refresh_message(self._PSbook_current_sheet._preview_model.itemFromIndex(index).cell.value)
    def select_extended_preview(self,index):
        pass
    def select_sync_select_all_ps_headers(self,state):
        if state == Qt.Checked:
            self._PSbook_current_sheet.select_all_headers()
        elif state == Qt.Unchecked:
            self._PSbook_current_sheet.unselect_all_headers()
    def select_sync_select_all_cas_headers(self,state):
        if state == Qt.Checked:
            self._CASbook_current_sheet.select_all_headers()
        elif state == Qt.Unchecked:
            self._CASbook_current_sheet.unselect_all_headers()
    ##################################################
    #       Sync
    ##################################################
    def select_sync_ps_to_cas(self):
        xml_names_ps = []
        headers_cas = []
        sync_list = []
        # pick out all xmlnames in cas file
        xml_names_cas = self._CASbook_current_sheet.xml_names()
        headers_ps = self._PSbook_current_sheet.checked_headers()
        
        for xml_name_cas in xml_names_cas:
            # there are some white space row in xmlname of cas file
            if xml_name_cas.value == None:
                print 'xml_name_cas is None'
                continue
            #pick out specified xmlname in ps file
            xml_name_ps = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_cas.value)
            if xml_name_ps == None:
                print 'could not find %s in ps file'%xml_name_cas.value
                continue
            #xml_names_ps.append(xml_name_ps)
            #sync_list.append((xml_name_ps,xml_name_cas))
            
            for header_ps in headers_ps:
                print '%s'%header_ps.value
                header_cas = self._CASbook_current_sheet.search_header_by_value(header_ps.value)
                if header_cas == None:
                    print 'could not find %s in cas file'%header_cas.value
                    continue
                #headers_cas.append(header_cas)
                #sync_list.append((xml_name_ps,header_ps,xml_name_cas,header_cas))
                
                source_item = self._PSbook_current_sheet.cell(xml_name_ps.row,header_ps.col)
                target_item = self._CASbook_current_sheet.cell(xml_name_cas.row,header_cas.col)
                target_item.value = source_item.value
        print 'sync ps to cas done'
#        for pair in sync_list:
#            #source_item = pair[0].get_item_by_header(pair[1])
#            #target_item = pair[2].get_item_by_header(pair[3])
#            source_item = self._PSbook_current_sheet.cell(pair[0].row,pair[1].col)
#            target_item = self._CASbook_current_sheet.cell(pair[2].row,pair[3].col)
#            #print '<<<<<source>>>>>ps:header=%s,xmlname=%s,value=%s,position:row %s,col %s'%(source_item.header.value,source_item.xmlname.value,source_item.value,source_item.row,source_item.col)
#            #print '>>>>>target<<<<<cas:header=%s,xmlname=%s,value=%s,position:row %s,col %s'%(target_item.header.value,target_item.xmlname.value,target_item.value,target_item.row,target_item.col)
#            target_item.value = source_item.value
#            #self._CASbook_current_sheet.cell(pair[2].row,pair[3].col).value = self._PSbook_current_sheet.cell(pair[0].row,pair[1].col).value
#        print 'sync ps to cas done'

    def select_sync_cas_to_ps(self):
        xml_names_ps = []
        headers_cas = []
        headers_ps = []
        sync_list = []
        # pick out all xmlnames in cas file
        xml_names_cas = self._CASbook_current_sheet.xml_names()
        headers_cas = self._CASbook_current_sheet.checked_headers()
        
        for xml_name_cas in xml_names_cas:
            # there are some white space row in xmlname of cas file
            if xml_name_cas.value == None:
                print 'xml_name_cas is None'
                continue
            #pick out spcified xmlname in ps file
            xml_name_ps = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_cas.value)
            if xml_name_ps == None:
                print 'could not find %s in ps file'%xml_name_cas.value
                continue
            
            for header_cas in headers_cas:
                header_ps = self._PSbook_current_sheet.search_header_by_value(header_cas.value)
                if header_ps == None:
                    print 'could not find %s in cas file'%header_ps.value
                    continue
                headers_ps.append(header_ps)
                sync_list.append((xml_name_cas,header_cas,xml_name_ps,header_ps))
        for pair in sync_list:
            #source_item = pair[0].get_item_by_header(pair[1])
            #print source_item._cell.col_idx
            #target_item = pair[2].get_item_by_header(pair[3])
            source_item = self._CASbook_current_sheet.cell(pair[0].row,pair[1].col)
            target_item = self._PSbook_current_sheet.cell(pair[2].row,pair[3].col)
            #print '<<<<<source>>>>>cas:header=%s,xmlname=%s,value=%s,position:row %d,col %d'%(source_item.header.value,\
            #        source_item.xmlname.value,\
            #        source_item.value,\
            #        source_item.row,\
            #        source_item.col)
            #print '>>>>>target<<<<<ps:header=%s,xmlname=%s,value=%s,position:row %d,col %d'%(target_item.header.value,target_item.xmlname.value,target_item.value,target_item.row,target_item.col)
            target_item.value = source_item.value
            #self._PSbook_current_sheet.cell(pair[2].row,pair[3].col).value = self._CASbook_current_sheet.cell(pair[0].row,pair[1].col).value
            # do a copy style
            alignment = copy(target_item.cell.alignment)
            alignment.wrapText = True
            target_item.cell.alignment = alignment
        # adjust column width
        for header_ps in headers_ps:
            self._PSbook_current_sheet._worksheet.column_dimensions[header_ps.col_letter].width = 25
        print 'sync cas to ps done'



    def refresh_cas_book_name(self,model):
        self._window.update_cas_file(model)
    def refresh_ps_book_name(self,model):
        self._window.update_ps_file(model)
    def refresh_cas_sheet_name(self,model):
        self._window.update_cas_sheets(model)
    def refresh_ps_sheet_name(self,model):
        self._window.update_ps_sheets(model)
    def refresh_preview(self,model):
        self._window.update_preview(model)
    def refresh_ps_header(self,model):
        self._window.update_ps_header(model)
    def refresh_cas_header(self,model):
        self._window.update_cas_header(model)
    def refresh_message(self,model):
        self._window.update_message(model)
