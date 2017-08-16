# -*- coding: utf-8 -*- 
import sys 
import os
from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
from view import Window
from model import *
from copy import copy
from controller.FileStack import *
import xlwings as xw
import random
import string
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
        self._PSbook_current_idx = None
        self._PSbook_current_sheet_name = None
        self._PSbook_autosave_flag = False
        self._CASbook = None
        self._CASbook_wr = None
        self._CASbook_name = ''
        self._CASbook_sheets = QStandardItemModel()
        self._CASbook_current_sheet = None
        self._CASbook_current_idx = None
        self._CASbook_current_sheet_name = None
        self._CASbook_autosave_flag = False
        self._PSstack = None
        self._CASstack = None
        self.start_xlwings_app()
        self.init_model()
        self.init_file_stack()
        self.init_GUI()
        self.bind_GUI_event()
        self.show_GUI()
    def __del__(self):
        self._xw_app.quit()
    def init_model(self):
        self._comparison_append_model = QStandardItemModel()
        self._comparison_delete_model = QStandardItemModel()
        self._preview_selected_cell = None
    def init_file_stack(self):
        self._PSstack = FileStack()
        self._CASstack = FileStack()
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
        self._window.bind_comparison_start(self.comparison_start)
        self._window.bind_comparison_delete(self.comparison_delete)
        self._window.bind_comparison_append(self.comparison_append)
        self._window.bind_comparison_select_all_delete(self.comparison_select_all_delete)
        self._window.bind_comparison_select_all_append(self.comparison_select_all_append)
        self._window.bind_preview_add(self.preview_add)
        self._window.bind_preview_delete(self.preview_delete)
        self._window.bind_preview_lock(self.preview_lock)
        self._window.bind_undo_cas(self.undo_cas)
        self._window.bind_undo_ps(self.undo_ps)
        self._window.bind_select_extended_preview(self.select_extended_preview)
    def show_GUI(self):
        self._window.show()
        sys.exit(self._application.exec_())       
    def start_xlwings_app(self):
        self._xw_app = xw.App()#visible=False)
    ##################################################
    #       Actions
    ##################################################

    
    def open_cas(self):
        #########################
        #   Open cas
        #########################
        try:
            filename = Window.open_file_dialog()
            self.open_cas_by_name(filename)
            self._CASstack = FileStack()
        except:
            filename = None
    def open_cas_by_name(self,filename):
        self._CASbook = CASbook.CASbook(str(filename),self._xw_app)
        self._CASbook_wr = self._CASbook.workbook_wr
        self._CASbook_name = filename
        #########################
        #   Update model
        #########################
        self.init_model()
        self._CASbook.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_cas_book_name(self._CASbook.workbook_name)
        self.refresh_cas_sheet_name(self._CASbook.sheet_name_model)
        self.store_cas_file('original')
        #self._window.update_cas_file(self._CASbook_name)
        #self._window.update_cas_sheets(self._CASbook.sheets_name)

    def open_cas_by_bytesio(self,bytesio):
        self._CASbook = CASbook.CASbook(bytesio,self._xw_app)
        self._CASbook_wr = self._CASbook.workbook_wr
        #########################
        #   Update model
        #########################
        self.init_model()
        self._CASbook.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_cas_sheet_name(self._CASbook.sheet_name_model)
        #self._window.update_cas_file(self._CASbook_name)
        #self._window.update_cas_sheets(self._CASbook.sheets_name)
         
    def open_ps(self):
        #########################
        #   Open ps
        #########################
        try:
            filename = Window.open_file_dialog()
            self.open_ps_by_name(str(filename))
            self._PSstack = FileStack()
        except:
            filename = None
    def open_ps_by_name(self,filename):
        self._PSbook = PSbook.PSbook(filename)
        #try:
        #    self._PSbook = PSbook.PSbook(filename)
        #    print "open succeed"
        #except:
        #    print "not a excel"
        self._PSbook_name = filename
        #########################
        #   Update model
        #########################
        self.init_model()
        self._PSbook.update_model()
        #self._window.update_ps_file(self._PSbook_name)
        #self._window.update_ps_sheets(self._PSbook.sheets_name)
        #########################
        #   Refresh UI
        #########################
        self.refresh_ps_book_name(self._PSbook.workbook_name)
        self.refresh_ps_sheet_name(self._PSbook.sheet_name_model)
        self.store_ps_file('original',self._PSbook.virtual_workbook)

    def open_ps_by_bytesio(self,bytesio):
        self._PSbook = PSbook.PSbook(bytesio)
        #try:
        #    self._PSbook = PSbook.PSbook(bytesio)
        #    print "open succeed"
        #except:
        #    print "not a excel"
        #self._PSbook_name = bytesio
        #########################
        #   Update model
        #########################
        self.init_model()
        self._PSbook.update_model()
        #self._window.update_ps_file(self._PSbook_name)
        #self._window.update_ps_sheets(self._PSbook.sheets_name)
        #########################
        #   Refresh UI
        #########################
        #self.refresh_ps_book_name(self._PSbook.workbook_name)
        self.refresh_ps_sheet_name(self._PSbook.sheet_name_model)


    def save_cas(self):
        pass
    def save_ps(self):
        pass
    def saveas_cas(self):
        fileName = Window.save_file_dialog()
        self._CASbook.workbook_wr.save(str(fileName))
        self._CASbook = CASbook.CASbook(str(fileName),self._xw_app)
        self._CASbook_wr = self._CASbook.workbook_wr
        self._CASbook_name = fileName
        self._CASstack = FileStack()
        #########################
        #   Update model
        #########################
        self.init_model()
        self._CASbook.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_cas_book_name(self._CASbook.workbook_name)
        self.refresh_cas_sheet_name(self._CASbook.sheet_name_model)

        self._CASbook.save_as(str(fileName))
        self.refresh_message('save cas to %s'%fileName)
    def saveas_ps(self):
        #'''
        #Solution 1
        #'''
        fileName = Window.save_file_dialog()
        self._PSbook.save_as(str(fileName))
        self._PSbook = PSbook.PSbook(str(fileName))
        self._PSbook_name = fileName
        self._PSstack = FileStack()
        #########################
        #   Update model
        #########################
        self.init_model()
        self._PSbook.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_ps_book_name(self._PSbook.workbook_name)
        self.refresh_ps_sheet_name(self._PSbook.sheet_name_model)
        self.refresh_message('save ps to %s'%fileName)
        #Solution 2
        #fileName = Window.save_file_dialog()
        #self._PSbook.save_as(str(fileName))
        #self.refresh_message('save ps to %s'%fileName)
        

    def select_cas_sheet(self,sheet_idx):
        print 'auto_save = %s'%self._CASbook_autosave_flag
        if self._CASbook_autosave_flag != True:
            self._CASbook_current_sheet_idx = sheet_idx
        else:
            self._CASbook_autosave_flag = False
        print 'stored idx %d'%self._CASbook_current_sheet_idx
        print 'select sheet %d'%sheet_idx

        self._CASbook_current_sheet_name = self._CASbook.sheets_name[sheet_idx]
        print self._CASbook_current_sheet_name
        self._CASbook_current_sheet = self._CASbook.sheets[self._CASbook_current_sheet_name]
        #########################
        #   Update model
        #########################
        self.init_model()
        self._CASbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_cas_header(self._CASbook_current_sheet.header_model)
        self.comparison_start()
        self.refresh_message('select cas sheet:%s'%self._CASbook_current_sheet_idx)
        print 'cas_sheet :%s'%self._CASbook_current_sheet
    def select_ps_sheet(self,sheet_idx):
        print 'auto_save = %s'%self._PSbook_autosave_flag
        if self._PSbook_autosave_flag != True:
            self._PSbook_current_sheet_idx = sheet_idx
        else:
            self._PSbook_autosave_flag = False
        print 'stored idx %d'%self._PSbook_current_sheet_idx
        print 'select sheet %d'%sheet_idx
        self._PSbook_current_sheet_name = self._PSbook.sheets_name[sheet_idx]
        print self._PSbook_current_sheet_name
        self._PSbook_current_sheet = self._PSbook.sheets[self._PSbook_current_sheet_name]
        #########################
        #   Update model
        #########################
        self.init_model()
        self._PSbook_current_sheet.update_model()
        #self.refresh_ps_sheet_name(self._PSbook_current_sheet.sheet_name_model)
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_ps_header(self._PSbook_current_sheet.header_model)
        self.comparison_start()
        self.refresh_message('select ps sheet:%s'%self._PSbook_current_sheet_name)
        print 'ps_sheet :%s'%self._PSbook_current_sheet
    def select_preview(self,index):
        print self._PSbook_current_sheet._preview_model.itemFromIndex(index).cell.value
        self._preview_selected_cell = self._PSbook_current_sheet._preview_model.itemFromIndex(index).cell
        #########################
        #   Refresh UI
        #########################
        self.refresh_message(self._PSbook_current_sheet._preview_model.itemFromIndex(index).cell.value)
        self.refresh_selected_cell((self._preview_selected_cell.row,self._preview_selected_cell.col))
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
        #########################
        #   Data sync
        #########################
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
        #########################
        #   Update model
        #########################
        self._CASbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.store_cas_file('sync ps to cas')
        self.refresh_message('sync ps to cas done')
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
        #########################
        #   Data sync
        #########################
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
            alignment = copy(target_item.alignment)
            alignment.wrapText = True
            target_item.alignment = alignment
        # adjust column width
        for header_ps in headers_ps:
            self._PSbook_current_sheet._worksheet.column_dimensions[header_ps.col_letter].width = 25
        print 'sync cas to ps done'
        #########################
        #   Update model
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.store_ps_file('sync cas to ps',self._PSbook.virtual_workbook)
        self.refresh_message('sync cas to ps done')
    ##################################################
    #       Comparison
    ##################################################
    def comparison_start(self):
        self.init_model()
        try:
            if self._PSbook_current_sheet != None and self._CASbook_current_sheet != None:
                append_list = list(set(self._CASbook_current_sheet.xml_names_value()).difference(set(self._PSbook_current_sheet.xml_names_value())))
                delete_list = list(set(self._PSbook_current_sheet.xml_names_value()).difference(set(self._CASbook_current_sheet.xml_names_value())))
                for xml_name_value in append_list:
                    xml_name = self._CASbook_current_sheet.search_xmlname_by_value(xml_name_value)
                    item_append = QComparisonItem(xml_name)
                    item_append.setCheckState(Qt.Unchecked)
                    item_append.setCheckable(True)
                    self._comparison_append_model.appendRow(item_append)
                for xml_name_value in delete_list:
                    xml_name = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_value)
                    item_delete = QComparisonItem(xml_name)
                    item_delete.setCheckState(Qt.Unchecked)
                    item_delete.setCheckable(True)
                    self._comparison_delete_model.appendRow(item_delete)
        finally:
            self.refresh_comparison_append_list(self._comparison_append_model)
            self.refresh_comparison_delete_list(self._comparison_delete_model)
            self.refresh_message('comparison done')
        
                
            
    def comparison_delete(self):
        #########################
        #   Data operation
        #########################
        for delete_item in self.checked_delete():
            self._PSbook_current_sheet.delete_row(delete_item.cell.row,1)
            self._comparison_delete_model.removeRow(delete_item.row())
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_comparison_delete_list(self._comparison_delete_model)
        self.store_ps_file('comparison delete',self._PSbook.virtual_workbook)
        self.comparison_start()
        self.refresh_message('comparison delete done')
        print 'comparison delete done'
    def comparison_append(self):
        #########################
        #   Data operation
        #########################
        if self._preview_selected_cell != None:
            self._PSbook_current_sheet.add_row(self._preview_selected_cell.row,len(self.checked_append()),PSsheet.DOWN)
            overwrite_row = self._preview_selected_cell.row
            for append_item in self.checked_append():
                overwrite_row += 1
                self._PSbook_current_sheet.worksheet.cell(row=overwrite_row,column=self._PSbook_current_sheet.xmlname.col).value = append_item.value
                self._comparison_append_model.removeRow(append_item.row())
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_comparison_append_list(self._comparison_append_model)
        self.store_ps_file('comparison append',self._PSbook.virtual_workbook)
        self.comparison_start()
        self.refresh_message('comparison append done')
        print 'comparison append done'
        #for append_item in self.checked_append():
        #    self._PSbook_current_sheet.add_row(
    def comparison_select_all_delete(self,state):
        for i in range(self._comparison_delete_model.rowCount()):
            item = self._comparison_delete_model.item(i)
            item.setCheckState(state)
    def comparison_select_all_append(self,state):
        for i in range(self._comparison_append_model.rowCount()):
            item = self._comparison_append_model.item(i)
            item.setCheckState(state)
    def checked_delete(self):
        items = []
        for i in range(self._comparison_delete_model.rowCount()):
            item = self._comparison_delete_model.item(i)
            if item.checkState() == Qt.Checked:
                items.append(item)
        return items
    def checked_append(self):
        items = []
        for i in range(self._comparison_append_model.rowCount()):
            item = self._comparison_append_model.item(i)
            if item.checkState() == Qt.Checked:
                items.append(item)
        return items
    ##################################################
    #       Preview
    ##################################################
    def preview_add(self):
        #########################
        #   Data operation
        #########################
        if self._preview_selected_cell != None:
            self._PSbook_current_sheet.add_row(self._preview_selected_cell.row,1,PSsheet.DOWN)
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_message('added one line below row %d'%self._preview_selected_cell.row)
        self.store_ps_file('add',self._PSbook.virtual_workbook)

    def preview_delete(self):
        #########################
        #   Data operation
        #########################
        if self._preview_selected_cell != None:
            self._PSbook_current_sheet.delete_row(self._preview_selected_cell.row,1)
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_message('deleted row %d'%self._preview_selected_cell.row)
        self.store_ps_file('delete',self._PSbook.virtual_workbook)

    def preview_lock(self):
        #for item in self._PSbook_current_sheet.extended_preview_model():
        self._PSbook_current_sheet.unlock_all_cells()
        for status in self._PSbook_current_sheet.status():
            if status.value == 'POR':
                self._PSbook_current_sheet.lock_row(status.row,True)
        self._PSbook_current_sheet.lock_sheet(True)
        self.store_ps_file('lock',self._PSbook.virtual_workbook)
        self.refresh_message('lock sheet done')
    ##################################################
    #       Recover
    ##################################################
    def recover_ps_sheet_selected(self):
        print 'recover to sheet %d'%self._PSbook_current_sheet_idx
        self._window.update_ps_header_selected(self._PSbook_current_sheet_idx)
        self.select_ps_sheet(self._PSbook_current_sheet_idx)
    def recover_cas_sheet_selected(self):
        print 'recover to sheet %d'%self._CASbook_current_sheet_idx
        self._window.update_cas_header_selected(self._CASbook_current_sheet_idx)
        self.select_cas_sheet(self._CASbook_current_sheet_idx)
    def store_ps_file(self,action,file_content):
        self._PSbook_autosave_flag = True
        self._PSstack.push(FilePack(action,file_content))
        self.open_ps_by_bytesio(self._PSstack.currentFile.fh)
        self.recover_ps_sheet_selected()
#    def store_cas_file(self,action):
#        cas_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,8))
#        print 'name: %s'%cas_file_name
#        bytesio = BytesIO()
#        self._CASbook.workbook_wr.save(cas_file_name)
#        with open(cas_file_name+r'.xlsx','rb') as cas_file:
#            bytesio.write(cas_file.read())
#        #self._CASbook_autosave_flag = True
#        self._CASstack.push(FilePack(action,bytesio.getvalue()))
#        bytesio.close()
    def store_cas_file(self,action):
        cas_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,8))
        print 'name: %s'%cas_file_name
        self._CASbook.workbook_wr.save(cas_file_name)
        self._CASstack.push(CasPack(action,cas_file_name))
    def undo_ps(self):
        f = self._PSstack.pop()
        if f != None:
            self._PSbook_autosave_flag = True
            self.open_ps_by_bytesio(f[1].fh)
            self.recover_ps_sheet_selected()
            self.comparison_start()
            self.refresh_message('revert action:%s'%f[0])
        else:
            self.refresh_message('Already at oldest change')
            
    def undo_cas(self):
        f = self._CASstack.pop()
        if f != None:
            self._CASbook_autosave_flag = True
            self.open_cas_by_bytesio(f[1].file_name+r'.xlsx')
            self.recover_cas_sheet_selected()
            self.comparison_start()
            self.refresh_message('revert action:%s'%f[0])
        else:
            self.refresh_message('Already at oldest change')

    def select_extended_preview(self):
        pass



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
    def refresh_comparison_delete_list(self,model):
        self._window.update_comparison_delete_list(model)
    def refresh_comparison_append_list(self,model):
        self._window.update_comparison_append_list(model)
    def refresh_message(self,model):
        self._window.update_message(model)
    def refresh_selected_cell(self,model):
        self._window.update_selected_cell(model)




class QComparisonItem(QStandardItem):
    def __init__(self,cell):
        if cell.value != None:
            super(QComparisonItem,self).__init__(cell.value)
        else:
            super(QComparisonItem,self).__init__('')
        self._cell = cell
    
    @property
    def cell(self):
        return self._cell
    @property
    def value(self):
        return self._cell.value
#    @property
#    def row(self):
#        return self._cell.row
#    @property
#    def col(self):
#        return self._cell.col
    @property
    def col_letter(self):
        return self._cell.col_letter

