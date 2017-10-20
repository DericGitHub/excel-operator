# -*- coding: UTF-8 -*- 
import sys 
import os
from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
from view import Window,ExtendedPreview
from model import Workbook,Worksheet,Workcell,CASbook,CASsheet,PSbook,PSsheet
from copy import copy
from controller.FileStack import *
import xlwings as xw
import random
import string
import shutil
import time
import logging
import re
def model2list(model):
    result = []
    for i in range(model.rowCount()):
        result.append(model.item(i).value)
    return result
def qsort(arr,first,last):
    if first < last:
        div = Partition(arr,first,last)
        qsort(arr,first,div)
        qsort(arr,div+1,last)
    else:
        return
def Partition(arr,first,last):
    i = first -1
    for j in range(first,last):
        if arr[j].cell.row <= arr[last].cell.row:
            i=i+1
            arr[i],arr[j]=arr[j],arr[i]
    arr[i+1],arr[last]=arr[last],arr[i+1]
    return i
class MainControllerUI(QObject):
    filename_pattern = re.compile(r'([^<>/\\\|:""\*\?]+)\.\w+$')
    cas_pattern = re.compile(r'.*cas.*',re.I)
    ps_pattern = re.compile(r'.*ps.*',re.I)
    def __init__(self,queue_wr=None,queue_rd=None):
        super(MainControllerUI,self).__init__()
        self._application = None
        self._progressBar_status = 0
        self._CASbook_modified = False
        self._PSbook_modified = False
        self._queue_wr = queue_wr
        self._queue_rd = queue_rd
        self._status = True
    def run(self):
        self.init_GUI()
        self.init_worker()
        self.show_GUI()
    def __del__(self):
        if self._loop_thread is not None:
            self._loop_thread.stop()
            self._loop_thread.quit()
            self._loop_thread.wait()
    def init_GUI(self):
        self._application = QApplication(sys.argv)
        self._window = Window.Window()
        self._extended_preview = ExtendedPreview.ExtendedPreview()
        self.bind_GUI_event()
    def init_worker(self):
        self._loop_thread = MainControllerUILoop(self._queue_rd)
        self.bind_worker_event(self._loop_thread)
        self._loop_thread.start()

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
        self._window.bind_comparison_append_with_color(self.comparison_append_with_color)
        self._window.bind_preview_add(self.preview_add)
        self._window.bind_preview_delete(self.preview_delete)
        self._window.bind_preview_lock(self.preview_lock)
        self._window.bind_undo_cas(self.undo_cas)
        self._window.bind_undo_ps(self.undo_ps)
        self._window.bind_select_extended_preview(self.select_extended_preview)
        self._window.bind_ps_header_changed(self.ps_header_changed)
        self._window.bind_cas_header_changed(self.cas_header_changed)
        self._window.bind_comparison_append_list_changed(self.comparison_append_list_changed)
        self._window.bind_comparison_delete_list_changed(self.comparison_delete_list_changed)
        self._window.bind_search(self._window.search_preview)
    @pyqtSlot()
    def open_cas(self):
        if self._CASbook_modified == True:
            ret = self._window.open_file_confirm()
            if ret == 0:
                self.save_cas()
            elif ret == 1:
                self.saveas_cas()
        filepath = Window.open_file_dialog()
        if filepath != None and filepath != '':
            try:
                str(filepath)
                filename = self.filename_pattern.search(filepath).group(1)
                if self.cas_pattern.search(str(filename)):
                    self._queue_wr.put((r'open_cas',str(filepath)))
                else:
                    self._window.pop_up_message('CAS file not found')
            except:
                self._window.pop_up_message('File name unsupported')
               
    @pyqtSlot()
    def open_ps(self):
        if self._PSbook_modified == True:
            ret = self._window.open_file_confirm()
            if ret == 0:
                self.save_ps()
            elif ret == 1:
                self.saveas_ps()
        filepath = Window.open_file_dialog()
        if filepath != None and filepath != '':
            try:
                str(filepath)
                filename = self.filename_pattern.search(filepath).group(1)
                if self.ps_pattern.search(str(filename)):
                    self._queue_wr.put((r'open_ps',str(filepath)))
                else:
                    self._window.pop_up_message('PS file not found')
            except:
                self._window.pop_up_message('File name unsupported')
    @pyqtSlot()
    def save_cas(self):
        self._queue_wr.put((r'save_cas',))
    @pyqtSlot()
    def save_ps(self):
        self._queue_wr.put((r'save_ps',))
    @pyqtSlot()
    def saveas_cas(self):
        filename = Window.save_file_dialog()
        try:
            self._queue_wr.put((r'saveas_cas',str(filename)))
        except:
            self._window.pop_up_message('File name unsupported')
    @pyqtSlot()
    def saveas_ps(self):
        filename = Window.save_file_dialog()
        try:
            self._queue_wr.put((r'saveas_ps',str(filename)))
        except:
            self._window.pop_up_message('File name unsupported')
    @pyqtSlot(int)
    def select_cas_sheet(self,index):
        self._queue_wr.put((r'select_cas_sheet',index))
    @pyqtSlot(int)
    def select_ps_sheet(self,index):
        self._queue_wr.put((r'select_ps_sheet',index))
    @pyqtSlot(int)
    def select_preview(self,index):
        self._queue_wr.put((r'select_preview',index.row(),index.column()))
    @pyqtSlot()
    def select_sync_ps_to_cas(self):
        self._queue_wr.put((r'select_sync_ps_to_cas',))
    @pyqtSlot()
    def select_sync_cas_to_ps(self):
        self._queue_wr.put((r'select_sync_cas_to_ps',))
    @pyqtSlot(bool)
    def select_sync_select_all_ps_headers(self,state):
        for i in range(self._window.ui.ps_header.model().rowCount()):
            self._window.ui.ps_header.model().item(i).setCheckState(state)
        self._queue_wr.put((r'select_sync_select_all_ps_headers',state))
    @pyqtSlot(bool)
    def select_sync_select_all_cas_headers(self,state):
        for i in range(self._window.ui.cas_header.model().rowCount()):
            self._window.ui.cas_header.model().item(i).setCheckState(state)
        self._queue_wr.put((r'select_sync_select_all_cas_headers',state))
    @pyqtSlot()
    def comparison_start(self):
        self._queue_wr.put((r'comparison_start',))
    @pyqtSlot()
    def comparison_delete(self):
        self._queue_wr.put((r'comparison_delete',))
    @pyqtSlot()
    def comparison_append(self):
        self._queue_wr.put((r'comparison_append',))
    @pyqtSlot(bool)
    def comparison_select_all_delete(self,state):
        for i in range(self._window.ui.comparison_delete_list.model().rowCount()):
            self._window.ui.comparison_delete_list.model().item(i).setCheckState(state)
        self._queue_wr.put((r'comparison_select_all_delete',state))
    @pyqtSlot(bool)
    def comparison_select_all_append(self,state):
        for i in range(self._window.ui.comparison_append_list.model().rowCount()):
            self._window.ui.comparison_append_list.model().item(i).setCheckState(state)
        self._queue_wr.put((r'comparison_select_all_append',state))
    @pyqtSlot(bool)
    def comparison_append_with_color(self,state):
        self._queue_wr.put((r'comparison_append_with_color',state))
    @pyqtSlot()
    def preview_add(self):
        self._queue_wr.put((r'preview_add',))
    @pyqtSlot()
    def preview_delete(self):
        self._queue_wr.put((r'preview_delete',))
    @pyqtSlot()
    def preview_lock(self):
        self._queue_wr.put((r'preview_lock',))
    @pyqtSlot()
    def undo_cas(self):
        self._queue_wr.put((r'undo_cas',))
    @pyqtSlot()
    def undo_ps(self):
        self._queue_wr.put((r'undo_ps',))
    @pyqtSlot()
    def select_extended_preview(self):
        self._extended_preview.show()
        self._queue_wr.put((r'select_extended_preview',))

    @pyqtSlot(QModelIndex)
    def ps_header_changed(self,index):
        self._queue_wr.put((r'ps_header_changed',int(index.row()),int(self._window.ui.ps_header.model().itemFromIndex(index).checkState())))
    @pyqtSlot(QModelIndex)
    def cas_header_changed(self,index):
        self._queue_wr.put((r'cas_header_changed',int(index.row()),int(self._window.ui.cas_header.model().itemFromIndex(index).checkState())))
    @pyqtSlot(QModelIndex)
    def comparison_append_list_changed(self,index):
        self._queue_wr.put((r'comparison_append_list_changed',int(index.row()),int(self._window.ui.comparison_append_list.model().itemFromIndex(index).checkState())))
    @pyqtSlot(QModelIndex)
    def comparison_delete_list_changed(self,index):
        self._queue_wr.put((r'comparison_delete_list_changed',int(index.row()),int(self._window.ui.comparison_delete_list.model().itemFromIndex(index).checkState())))
    @pyqtSlot(bool)
    def set_CASbook_modified(self,state):
        self._CASbook_modified = state
    @pyqtSlot(bool)
    def set_PSbook_modified(self,state):
        self._PSbook_modified = state
        

    @pyqtSlot(float)
    def animation_progressBar(self,model):
        if self._progressBar_status == 0 and model ==100:
            while self._progressBar_status <100:
                self._progressBar_status += 0.05
                self._window.update_progressBar(self._progressBar_status)
        else:
            self._progressBar_status = model
            self._window.update_progressBar(self._progressBar_status)
    @pyqtSlot(str)
    def open_action_confirm(self,action):
        ret = self._window.open_action_confirm(action)
        if ret == 0:
            self._queue_wr.put((r'comparison_append_to_tail',))
        
    def bind_worker_event(self,worker):
        worker.signal_refresh_cas_book_name.connect(self._window.update_cas_file)
        worker.signal_refresh_ps_book_name.connect(self._window.update_ps_file)
        worker.signal_refresh_cas_sheet_name.connect(self._window.update_cas_sheets)
        worker.signal_refresh_ps_sheet_name.connect(self._window.update_ps_sheets)
        worker.signal_refresh_preview.connect(self._window.update_preview)
        worker.signal_refresh_ps_header.connect(self._window.update_ps_header)
        worker.signal_refresh_cas_header.connect(self._window.update_cas_header)
        worker.signal_refresh_comparison_delete_list.connect(self._window.update_comparison_delete_list)
        worker.signal_refresh_comparison_append_list.connect(self._window.update_comparison_append_list)
        worker.signal_refresh_message.connect(self._window.update_message)
        worker.signal_refresh_msg.connect(self._window.update_msg)
        worker.signal_refresh_warning.connect(self._window.pop_up_message)
        worker.signal_refresh_dialog.connect(self.open_action_confirm)
        worker.signal_refresh_selected_cell.connect(self._window.update_selected_cell)
        worker.signal_refresh_progressBar.connect(self._window.update_progressBar)
        worker.signal_animation_progressBar.connect(self.animation_progressBar)
        worker.signal_refresh_ps_header_selected.connect(self._window.update_ps_header_selected)
        worker.signal_refresh_cas_header_selected.connect(self._window.update_cas_header_selected)
        worker.signal_refresh_extended_preview.connect(self._extended_preview.update_extended_preview)
        worker.signal_set_CASbook_modified.connect(self.set_CASbook_modified)
        worker.signal_set_PSbook_modified.connect(self.set_PSbook_modified)
        
    def show_GUI(self):
        self._window.show()
        sys.exit(self._application.exec_())       
class MainControllerUILoop(QThread):
    signal_refresh_cas_book_name = pyqtSignal(str)
    signal_refresh_ps_book_name = pyqtSignal(str)
    signal_refresh_cas_sheet_name = pyqtSignal(list)
    signal_refresh_ps_sheet_name = pyqtSignal(list)
    signal_refresh_preview = pyqtSignal(list)
    signal_refresh_ps_header = pyqtSignal(list)
    signal_refresh_cas_header = pyqtSignal(list)
    signal_refresh_comparison_delete_list = pyqtSignal(list)
    signal_refresh_comparison_append_list = pyqtSignal(list)
    signal_refresh_message = pyqtSignal(str)
    signal_refresh_msg = pyqtSignal(str)
    signal_refresh_warning = pyqtSignal(str)
    signal_refresh_dialog = pyqtSignal(str)
    signal_refresh_selected_cell = pyqtSignal(list)
    signal_refresh_progressBar = pyqtSignal(int)
    signal_animation_progressBar = pyqtSignal(int)
    signal_refresh_ps_header_selected = pyqtSignal(int)
    signal_refresh_cas_header_selected = pyqtSignal(int)
    signal_refresh_extended_preview = pyqtSignal(list)
    signal_set_CASbook_modified = pyqtSignal(bool)
    signal_set_PSbook_modified = pyqtSignal(bool)
    def __init__(self,queue_rd=None,parent=None):
        super(MainControllerUILoop,self).__init__(parent)
        self._status = True
        self._queue_rd = queue_rd
    def run(self):
        while self._status == True:
            if not self._queue_rd.empty():
                task = self._queue_rd.get()
                #print 'ui got task:%s'%str(task)
                if task[0] == r'refresh_cas_book_name':
                    self.signal_refresh_cas_book_name.emit(task[1])
                elif task[0] == r'refresh_ps_book_name':
                    self.signal_refresh_ps_book_name.emit(task[1])
                elif task[0] == r'refresh_cas_sheet_name':
                    self.signal_refresh_cas_sheet_name.emit(task[1])
                elif task[0] == r'refresh_ps_sheet_name':
                    self.signal_refresh_ps_sheet_name.emit(task[1])
                elif task[0] == r'refresh_preview':
                    self.signal_refresh_preview.emit(task[1])
                elif task[0] == r'refresh_ps_header':
                    self.signal_refresh_ps_header.emit(task[1])
                elif task[0] == r'refresh_cas_header':
                    self.signal_refresh_cas_header.emit(task[1])
                elif task[0] == r'refresh_comparison_delete_list':
                    self.signal_refresh_comparison_delete_list.emit(task[1])
                elif task[0] == r'refresh_comparison_append_list':
                    self.signal_refresh_comparison_append_list.emit(task[1])
                elif task[0] == r'refresh_message':
                    self.signal_refresh_message.emit(task[1])
                elif task[0] == r'refresh_msg':
                    self.signal_refresh_msg.emit(task[1])
                elif task[0] == r'refresh_warning':
                    self.signal_refresh_warning.emit(task[1])
                elif task[0] == r'refresh_dialog':
                    self.signal_refresh_dialog.emit(task[1])
                elif task[0] == r'refresh_selected_cell':
                    self.signal_refresh_selected_cell.emit(task[1])
                elif task[0] == r'refresh_progressBar':
                    self.signal_refresh_progressBar.emit(task[1])
                elif task[0] == r'animation_progressBar':
                    self.signal_animation_progressBar.emit(task[1])
                elif task[0] == r'refresh_ps_header_selected':
                    self.signal_refresh_ps_header_selected.emit(task[1])
                elif task[0] == r'refresh_cas_header_selected':
                    self.signal_refresh_cas_header_selected.emit(task[1])
                elif task[0] == r'refresh_extended_preview':
                    self.signal_refresh_extended_preview.emit(task[1])
                elif task[0] == r'refresh_CASbook_modified':
                    self.signal_set_CASbook_modified.emit(task[1])
                elif task[0] == r'refresh_PSbook_modified':
                    self.signal_set_PSbook_modified.emit(task[1])
                elif task[0] == r'stop':
                    self.stop()
            time.sleep(0.05)
    def stop(self):
        self._status = False
##################################################
#       class for handling application logic
##################################################
class MainController(object):
    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,queue_wr=None,queue_rd=None):
        super(MainController,self).__init__()
        self._status = True
        self._queue_wr = queue_wr
        self._queue_rd = queue_rd
    def run(self):
        #import pythoncom
        #pythoncom.CoInitialize() 
        self._xw_app = None
        self._xw_app_2 = None
        self._PSbook = None
        self._PSbook_name = ''
        self._PSbook_sheets = QStandardItemModel()
        self._PSbook_current_sheet = None
        self._PSbook_current_sheet_idx = 0
        self._PSbook_current_sheet_name = None
        self._PSbook_autosave_flag = False
        self._PSbook_modified = False
        self._CASbook = None
        self._CASbook_name = ''
        self._CASbook_sheets = QStandardItemModel()
        self._CASbook_current_sheet = None
        self._CASbook_current_sheet_idx = 0
        self._CASbook_current_sheet_name = None
        self._CASbook_autosave_flag = False
        self._CASbook_modified = False
        self._PSstack = None
        self._CASstack = None
        self._progressBar_status = 0
        self._append_color = False
        self.init_logging()
        self.init_tmp_directory()
        self.start_xlwings_app()
        self.init_model()
        self.init_file_stack()
        while self._status == True:
            if not self._queue_rd.empty():
                task = self._queue_rd.get()
                #print 'controller got task:%s'%str(task)
                try:
                    if task[0] == 'open_cas':
                        self.open_cas_by_name(task[1])
                    elif task[0] == 'open_ps':
                        self.open_ps_by_name(task[1])
                    elif task[0] == 'save_cas':
                        self.save_cas()
                    elif task[0] == 'save_ps':
                        self.save_ps()
                    elif task[0] == 'saveas_cas':
                        self.saveas_cas(task[1])
                    elif task[0] == 'saveas_ps':
                        self.saveas_ps(task[1])
                    elif task[0] == 'select_cas_sheet':
                        self.select_cas_sheet(task[1])
                    elif task[0] == 'select_ps_sheet':
                        self.select_ps_sheet(task[1])
                    elif task[0] == 'select_preview':
                        self.select_preview(task[1],task[2])
                    elif task[0] == 'select_sync_ps_to_cas':
                        self.select_sync_ps_to_cas()
                    elif task[0] == 'select_sync_cas_to_ps':
                        self.select_sync_cas_to_ps()
                    elif task[0] == 'select_sync_select_all_ps_headers':
                        self.select_sync_select_all_ps_headers(task[1])
                    elif task[0] == 'select_sync_select_all_cas_headers':
                        self.select_sync_select_all_cas_headers(task[1])
                    elif task[0] == 'comparison_start':
                        self.comparison_start()
                    elif task[0] == 'comparison_delete':
                        self.comparison_delete()
                    elif task[0] == 'comparison_append':
                        self.comparison_append()
                    elif task[0] == 'comparison_select_all_delete':
                        self.comparison_select_all_delete(task[1])
                    elif task[0] == 'comparison_select_all_append':
                        self.comparison_select_all_append(task[1])
                    elif task[0] == 'comparison_append_with_color':
                        self.comparison_append_with_color(task[1])
                    elif task[0] == 'preview_add':
                        self.preview_add()
                    elif task[0] == 'preview_delete':
                        self.preview_delete()
                    elif task[0] == 'preview_lock':
                        self.preview_lock()
                    elif task[0] == 'undo_cas':
                        self.undo_cas()
                    elif task[0] == 'undo_ps':
                        self.undo_ps()
                    elif task[0] == 'select_extended_preview':
                        self.select_extended_preview()
                    elif task[0] == 'ps_header_changed':
                        self.ps_header_changed(task[1],task[2])
                    elif task[0] == 'cas_header_changed':
                        self.cas_header_changed(task[1],task[2])
                    elif task[0] == 'comparison_append_list_changed':
                        self.comparison_append_list_changed(task[1],task[2])
                    elif task[0] == 'comparison_delete_list_changed':
                        self.comparison_delete_list_changed(task[1],task[2])
                    elif task[0] == 'comparison_append_to_tail':
                        self.comparison_append_to_tail()
                    elif task[0] == 'stop':
                        self.stop()
                except BaseException as e:
                    self.refresh_warning(str(e))
                    print e
            time.sleep(0.05)
    def stop(self):
        self._status = False

    def __del__(self):
        del self._CASbook
        del self._PSbook
        self._xw_app.quit()
        if len(self._xw_app_2.books) == 0:
            self._xw_app_2.quit()
        shutil.rmtree('tmp')
        print 'remove tmp'

    def init_logging(self):
        logging.basicConfig(level=logging.DEBUG,format='%(asctime)s        %(message)s',datefmt='%d %b %Y %H:%M:%S',filename=time.strftime('%m%d-%H-%M-%S')+'.log',filemode='w')
    def init_tmp_directory(self):
        if not os.path.isdir('tmp'):
            print 'create tmp'
            os.mkdir('tmp')
        else:
            print 'tmp exists'
    def init_model(self):
        self._comparison_append_model = QStandardItemModel()
        self._comparison_delete_model = QStandardItemModel()
        self._preview_selected_cell = None
    def init_file_stack(self):
        self._PSstack = FileStack()
        self._CASstack = FileStack()
    def start_xlwings_app(self):
        self._xw_app_2 = xw.App(visible=False,add_book=False)
        self._xw_app = xw.App(visible=False,add_book=False)
    
    def open_cas_by_name(self,filename):
        del self._CASbook
        self._CASbook = CASbook.CASbook(self.copy_cas(filename),self._xw_app)
        self._CASbook_name = filename
        #########################
        #   Update model
        #########################
        self.init_model()
        self._CASbook.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_cas_book_name(self._CASbook_name)
        self.refresh_cas_sheet_name(self._CASbook.sheetnames)
        self.refresh_msg('open cas file:%s'%self._CASbook_name)
        self.CASbook_modified = False

    def open_cas_by_bytesio(self,bytesio):
        del self._CASbook
        self._CASbook = CASbook.CASbook(bytesio,self._xw_app)
        #########################
        #   Update model
        #########################
        self.init_model()
        self._CASbook.update_model()
        #########################
        #   Refresh UI
        #########################
        self.CASbook_modified = False
         
    def open_ps_by_name(self,filename):
        del self._PSbook
        self._PSbook = PSbook.PSbook(self.copy_ps(filename),self._xw_app)
        self._PSbook_name = filename
        #########################
        #   Update model
        #########################
        self.init_model()
        self._PSbook.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_ps_book_name(self._PSbook_name)
        self.refresh_ps_sheet_name(self._PSbook.sheetnames)
        self.refresh_msg('open ps file:%s'%self._PSbook_name)
        self.PSbook_modified = False

    def open_ps_by_bytesio(self,bytesio):
        del self._PSbook
        self._PSbook = PSbook.PSbook(bytesio,self._xw_app)
        #########################
        #   Update model
        #########################
        self.init_model()
        self._PSbook.update_model()
        #########################
        #   Refresh UI
        #########################

        #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        ' The line below impact speed much, so remove it '
        #
        #self.refresh_ps_sheet_name(self._PSbook.sheetnames)
        #
        #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        #>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

        self.PSbook_modified = False


    def save_cas(self):
        self._CASbook.save_as(self._CASbook_name)
        self.open_cas_by_name(self._CASbook_name)
        self.recover_cas_sheet_selected()
        self.refresh_msg('saved cas file:%s'%self._CASbook_name)
        self.CASbook_modified = False
    def save_ps(self):
        self._PSbook.save_as(self._PSbook_name)
        self.open_ps_by_name(self._PSbook_name)
        self.recover_ps_sheet_selected()
        self.refresh_msg('saved ps file:%s'%self._PSbook_name)
        self.PSbook_modified = False
         
    def saveas_cas(self,fileName):
        self._CASbook.save_as(str(fileName))
        self._CASbook_name = str(fileName)
        self.open_cas_by_name(self._CASbook_name)
        self.recover_cas_sheet_selected()
        self.refresh_msg('saved cas file:%s'%str(fileName))
        self.CASbook_modified = False
    def saveas_ps(self,fileName):
        self._PSbook.save_as(str(fileName))
        self._PSbook_name = str(fileName)
        self.open_ps_by_name(self._PSbook_name)
        self.recover_ps_sheet_selected()
        self.refresh_msg('saved ps file:%s'%str(fileName))
        self.PSbook_modified = False
        

    def select_cas_sheet(self,sheet_idx):
        if self._CASbook_autosave_flag != True:
            if self._CASbook_current_sheet_idx != sheet_idx:
                self.refresh_msg('select cas sheet:%s'%self._CASbook.workbook.sheetnames[sheet_idx])
            self._CASbook_current_sheet_idx = sheet_idx
        else:
            self._CASbook_autosave_flag = False
        self._CASbook_current_sheet = self._CASbook.sheets[sheet_idx]
        #########################
        #   Update model
        #########################
        self.init_model()
        self._CASbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_cas_header(self._CASbook_current_sheet.header_list)
        self.comparison_start()
    def select_ps_sheet(self,sheet_idx):
        if self._PSbook_autosave_flag != True:
            if self._PSbook_current_sheet_idx != sheet_idx:
                self.refresh_msg('select ps sheet:%s'%self._PSbook.workbook.sheetnames[sheet_idx])
            self._PSbook_current_sheet_idx = sheet_idx
            print 'current_idx set to:%s'%self._PSbook_current_sheet_idx
        else:
            self._PSbook_autosave_flag = False
        self._PSbook_current_sheet = self._PSbook.sheets[sheet_idx]
        #########################
        #   Update model
        #########################
        self.init_model()
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_ps_header(self._PSbook_current_sheet.header_list)
        self.comparison_start()
    def select_preview(self,row,column):
        self._preview_selected_cell = self._PSbook_current_sheet._preview_model.item(row,column).cell
        #########################
        #   Refresh UI
        #########################
        self.refresh_selected_cell([self._preview_selected_cell.row,self._preview_selected_cell.col])
        self.refresh_msg('select cell:row %s,column %s'%(self._preview_selected_cell.row,self._preview_selected_cell.col))

    def select_sync_select_all_ps_headers(self,state):
        if state == Qt.Checked:
            self._PSbook_current_sheet.select_all_headers()
            self.refresh_msg('select all ps headers')
        elif state == Qt.Unchecked:
            self._PSbook_current_sheet.unselect_all_headers()
            self.refresh_msg('unselect all ps headers')
    def select_sync_select_all_cas_headers(self,state):
        if state == Qt.Checked:
            self._CASbook_current_sheet.select_all_headers()
            self.refresh_msg('select all cas headers')
        elif state == Qt.Unchecked:
            self._CASbook_current_sheet.unselect_all_headers()
            self.refresh_msg('unselect all cas headers')
    ##################################################
    #       Sync
    ##################################################
    @pyqtSlot()
    def select_sync_ps_to_cas(self):
        xml_names_row = []
        headers_column = []
        sync_list = []
        # pick out all xmlnames in cas file
        xml_names_cas = self._CASbook_current_sheet.xml_names()
        headers_ps = self._PSbook_current_sheet.checked_headers()
        
        #self.animation_progressBar(0)
        for xml_name_cas in xml_names_cas:
            # there are some white space row in xmlname of cas file
            if xml_name_cas.value == None:
                continue
            #pick out specified xmlname in ps file
            xml_name_ps = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_cas.value)
            if xml_name_ps == None:
                continue
            xml_names_row.append((xml_name_ps.row,xml_name_cas.row))
        #self.animation_progressBar(33)
            
        for header_ps in headers_ps:
            if header_ps.value == None or header_ps.value == 'xmlname':
                continue
            header_cas = self._CASbook_current_sheet.search_header_by_value(header_ps.value)
            if header_cas == None:
                continue
            headers_column.append((header_ps.column,header_cas.column))
        #self.animation_progressBar(66)        
        if len(headers_column) != 0:
            step = float("%0.2f"%(90.0 / len(headers_column) / len(xml_names_row)))
        else:
            self.refresh_warning('Please select items to sync.')
            return
        count = 0
        for columns in headers_column:
            for rows in xml_names_row:
                self._CASbook_current_sheet.cell_wr(rows[1],columns[1]).value = self._PSbook_current_sheet.cell(rows[0],columns[0]).value
                count += step
                self.animation_progressBar(count)        
        #########################
        #   Update model
        #########################
        self._CASbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.store_cas_file('sync ps to cas')
        self.refresh_msg('sync ps to cas done')
        self.animation_progressBar(100)        
        self.CASbook_modified = True

    def select_sync_cas_to_ps(self):
        #########################
        #   Data sync
        #########################
        xml_names_row = []
        headers_column = []
        headers_ps = []
        sync_list = []
        to_lock_flag = False
        if self._PSbook_current_sheet.lock_sheet_status() == True:
            to_lock_flag = True
            self._PSbook_current_sheet.unlock_sheet()

        # pick out all xmlnames in cas file
        xml_names_cas = self._CASbook_current_sheet.xml_names()
        headers_cas = self._CASbook_current_sheet.checked_headers()
        
        #self.animation_progressBar(0)
        for xml_name_cas in xml_names_cas:
            # there are some white space row in xmlname of cas file
            if xml_name_cas.value == None:
                continue
            #pick out spcified xmlname in ps file
            xml_name_ps = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_cas.value)
            if xml_name_ps == None:
                continue
            xml_names_row.append((xml_name_ps.row,xml_name_cas.row))
        #self.animation_progressBar(33)
            
        for header_cas in headers_cas:
            if header_cas.value == 'xmlname':
                continue
            header_ps = self._PSbook_current_sheet.search_header_by_value(header_cas.value)
            if header_ps == None:
                continue
            headers_column.append((header_ps.column,header_cas.column))
        #self.animation_progressBar(66)
        if len(headers_column) != 0:
            step = float("%0.2f"%(90.0 / len(headers_column) / len(xml_names_row)))
        else:
            self.refresh_warning('Please select items to sync.')
            return
        count = 0
        for columns in headers_column:
            for rows in xml_names_row:
                self._PSbook_current_sheet.cell_wr(rows[0],columns[0]).value = self._CASbook_current_sheet.cell(rows[1],columns[1]).value
                count += step
                self.animation_progressBar(count)
        #########################
        #   Update model
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_ps_header(self._PSbook_current_sheet.header_list)
        if to_lock_flag == True:
            to_lock_flag = False
            self._PSbook_current_sheet.lock_sheet()
        self.store_ps_file('sync cas to ps')
        self.refresh_msg('sync cas to ps done')
        self.animation_progressBar(100)
        self.PSbook_modified = True

    ##################################################
    #       Comparison
    ##################################################
    def comparison_start(self):
        #self.animation_progressBar(0)
        self.init_model()
        append_list = []
        delete_list = []
        t1 = []
        t2 = []
        try:
            if self._PSbook_current_sheet != None and self._CASbook_current_sheet != None and self._PSbook_current_sheet.xmlname != None and self._CASbook_current_sheet.xmlname != None:
                append_list = list(set(self._CASbook_current_sheet.xml_names_value()).difference(set(self._PSbook_current_sheet.xml_names_value())))
                delete_list = list(set(self._PSbook_current_sheet.xml_names_value()).difference(set(self._CASbook_current_sheet.xml_names_value())))
                for xml_name_value in append_list:
                    cell1 = self._CASbook_current_sheet.search_xmlname_by_value(xml_name_value)
                    t1.append(cell1)
                qsort(t1,0,len(t1)-1)
                print 'sort finish 5'
                for xml_name in t1:
                    item_append = QComparisonItem(xml_name)
                    item_append.setCheckState(Qt.Unchecked)
                    item_append.setCheckable(True)
                    self._comparison_append_model.appendRow(item_append)
                #self.animation_progressBar(40)
                for xml_name_value in delete_list:
                    cell2 = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_value)
                    t2.append(cell2)
                qsort(t2,0,len(t2)-1)
                for xml_name in t2:
                    item_delete = QComparisonItem(xml_name)
                    item_delete.setCheckState(Qt.Unchecked)
                    item_delete.setCheckable(True)
                    self._comparison_delete_model.appendRow(item_delete)
                #self.animation_progressBar(80)
        finally:
            self.refresh_comparison_append_list(model2list(self._comparison_append_model))
            self.refresh_comparison_delete_list(model2list(self._comparison_delete_model))
            #self.animation_progressBar(100)
        
                
            
    def comparison_delete(self):
        self.animation_progressBar(0)
        #########################
        #   Data operation
        #########################
        to_lock_flag = False
        if self._PSbook_current_sheet.lock_sheet_status() == True:
            to_lock_flag = True
            self._PSbook_current_sheet.unlock_sheet()
        if self.checked_delete_count() != 0:
            step = float("%0.2f"%(90.0 / self.checked_delete_count()))
        else:
            self.refresh_warning('Please select items to delete.')
            return
        count = 0
        for delete_item in self.checked_delete():
            #print 'delete %s line %s'%(delete_item.value,delete_item.cell.row)
            if delete_item.value == None:
                continue
            self._PSbook_current_sheet.delete_row(delete_item.cell.row,1)
            #self._comparison_delete_model.removeRow(delete_item.row())
            count += step
            self.animation_progressBar(count)
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_comparison_delete_list(model2list(self._comparison_delete_model))
        if to_lock_flag == True:
            to_lock_flag = False
            self._PSbook_current_sheet.lock_sheet()
        self.store_ps_file('comparison delete')
        self.comparison_start()
        self.refresh_msg('comparison delete done')
        self.animation_progressBar(100)
        self.PSbook_modified = True

    def comparison_append(self):
        self.animation_progressBar(0)
        #########################
        #   Data operation
        #########################
        if self._preview_selected_cell != None:
            to_lock_flag = False
            if self._PSbook_current_sheet.lock_sheet_status() == True:
                to_lock_flag = True
                self._PSbook_current_sheet.unlock_sheet()
            self._PSbook_current_sheet.add_row(self._preview_selected_cell.row,self.checked_append_count(),PSsheet.DOWN)
            overwrite_row = self._preview_selected_cell.row
            if self.checked_append_count() != 0:
                step = float("%0.2f"%(90.0 / self.checked_append_count()))
            else:
                self.refresh_warning('Please select items to append.')
                return
            count = 0
            for append_item in self.checked_append():
                overwrite_row += 1
                self._PSbook_current_sheet.cell_wr(overwrite_row,self._PSbook_current_sheet.xmlname.col).value = append_item.value
                if self._append_color == Qt.Checked:
                    self._PSbook_current_sheet.cell_wr(overwrite_row,self._PSbook_current_sheet.xmlname.col).color = (255,31,31)
                self._comparison_append_model.removeRow(append_item.row())
                count += step
                self.animation_progressBar(count)
            #########################
            #   Update model   
            #########################
            self._PSbook_current_sheet.update_model()
            #########################
            #   Refresh UI
            #########################
            self.refresh_preview(self._PSbook_current_sheet.preview_model)
            self.refresh_comparison_append_list(model2list(self._comparison_append_model))
            if to_lock_flag == True:
                to_lock_flag = False
                self._PSbook_current_sheet.lock_sheet()
            self.store_ps_file('comparison append')
            self.refresh_msg('comparison append done')
            self.animation_progressBar(100)
        else:
            self.animation_progressBar(100)
            self.refresh_dialog('Append to the tail of form')
    def comparison_append_to_tail(self):
        self.animation_progressBar(0)
        to_lock_flag = False
        if self._PSbook_current_sheet.lock_sheet_status() == True:
            to_lock_flag = True
            self._PSbook_current_sheet.unlock_sheet()
        self._PSbook_current_sheet.add_row(self._PSbook_current_sheet.last_xmlname_row,self.checked_append_count(),PSsheet.DOWN)
        overwrite_row = self._PSbook_current_sheet.last_xmlname_row
        if self.checked_append_count() != 0:
            step = float("%0.2f"%(90.0 / self.checked_append_count()))
        else:
            self.refresh_warning('Please select items to append.')
            return
        count = 0
        for append_item in self.checked_append():
            overwrite_row += 1
            self._PSbook_current_sheet.cell_wr(overwrite_row,self._PSbook_current_sheet.xmlname.col).value = append_item.value
            if self._append_color == Qt.Checked:
                self._PSbook_current_sheet.cell_wr(overwrite_row,self._PSbook_current_sheet.xmlname.col).color = (255,31,31)
            self._comparison_append_model.removeRow(append_item.row())
            count += step
            self.animation_progressBar(count)
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_comparison_append_list(model2list(self._comparison_append_model))
        if to_lock_flag == True:
            to_lock_flag = False
            self._PSbook_current_sheet.lock_sheet()
        self.store_ps_file('comparison append to tail')
        self.refresh_msg('comparison append to tail done')
        self.animation_progressBar(100)
        

    def comparison_select_all_delete(self,state):
        for i in range(self._comparison_delete_model.rowCount()):
            item = self._comparison_delete_model.item(i)
            item.setCheckState(state)
        if state == Qt.Checked:
            self.refresh_msg('select all delete items')
        else:
            self.refresh_msg('unselect all delete items')

    def comparison_select_all_append(self,state):
        for i in range(self._comparison_append_model.rowCount()):
            item = self._comparison_append_model.item(i)
            item.setCheckState(state)
        if state == Qt.Checked:
            self.refresh_msg('select all append items')
        else:
            self.refresh_msg('unselect all append items')
    def comparison_append_with_color(self,state):
        self._append_color = state

    def checked_delete(self):
        items = []
        for i in range(self._comparison_delete_model.rowCount()):
            item = self._comparison_delete_model.item(i)
            if item.checkState() == Qt.Checked:
                items.append(item)
        items.reverse()
        self.refresh_msg('checked delete items:%s'%str(map(lambda x:x.value,items)))
        return items
    def checked_delete_count(self):
        count = 0
        for i in range(self._comparison_delete_model.rowCount()):
            if self._comparison_delete_model.item(i).checkState() == Qt.Checked:
                count += 1
        return count 
    def checked_append(self):
        items = []
        for i in range(self._comparison_append_model.rowCount()):
            item = self._comparison_append_model.item(i)
            if item.checkState() == Qt.Checked:
                items.append(item)
        self.refresh_msg('checked append items:%s'%str(map(lambda x:x.value,items)))
        return items
    def checked_append_count(self):
        count = 0
        for i in range(self._comparison_append_model.rowCount()):
            if self._comparison_append_model.item(i).checkState() == Qt.Checked:
                count += 1
        return count
    ##################################################
    #       Preview
    ##################################################
    def preview_add(self):
        self.animation_progressBar(0)
        #########################
        #   Data operation
        #########################
        if self._preview_selected_cell != None:
            to_lock_flag = False
            if self._PSbook_current_sheet.lock_sheet_status() == True:
                to_lock_flag = True
                self._PSbook_current_sheet.unlock_sheet()
            self._PSbook_current_sheet.add_row(self._preview_selected_cell.row,1,PSsheet.DOWN)
            #self.animation_progressBar(80)
            #########################
            #   Update model   
            #########################
            self._PSbook_current_sheet.update_model()
            #########################
            #   Refresh UI
            #########################
            self.refresh_preview(self._PSbook_current_sheet.preview_model)
            self.refresh_msg('added one row below row %d'%self._preview_selected_cell.row)
            if to_lock_flag == True:
                to_lock_flag = False
                self._PSbook_current_sheet.lock_sheet()
            self.store_ps_file('add')
            self.animation_progressBar(100)
            self.PSbook_modified = True
        else:
            self.animation_progressBar(100)
            self.refresh_warning('please select a cell in preview')

    def preview_delete(self):
        self.animation_progressBar(0)
        #########################
        #   Data operation
        #########################
        if self._preview_selected_cell != None:
            to_lock_flag = False
            if self._PSbook_current_sheet.lock_sheet_status() == True:
                to_lock_flag = True
                self._PSbook_current_sheet.unlock_sheet()
            self._PSbook_current_sheet.delete_row(self._preview_selected_cell.row,1)
            #self.animation_progressBar(80)
            #########################
            #   Update model   
            #########################
            self._PSbook_current_sheet.update_model()
            #########################
            #   Refresh UI
            #########################
            self.refresh_preview(self._PSbook_current_sheet.preview_model)
            self.refresh_msg('deleted row %d'%self._preview_selected_cell.row)
            if to_lock_flag == True:
                to_lock_flag = False
                self._PSbook_current_sheet.lock_sheet()
            self.store_ps_file('delete')
            self.animation_progressBar(100)
            self.PSbook_modified = True
        else:
            self.animation_progressBar(100)
            self.refresh_warning('please select a cell in preview')

    def preview_lock(self):
        self.animation_progressBar(0)
        self._PSbook_current_sheet.unlock_sheet()
        self._PSbook_current_sheet.unlock_all_cells()
        self.animation_progressBar(20)
        for status in self._PSbook_current_sheet.status():
            if status.value == 'POR':
                self._PSbook_current_sheet.lock_row(status.row,True)
        self.animation_progressBar(80)
        self._PSbook_current_sheet.lock_sheet()
        self.store_ps_file('lock')
        self.refresh_msg('lock sheet done')
        self.animation_progressBar(100)
        self.PSbook_modified = True
    def ps_header_changed(self,row,state):
        self._PSbook_current_sheet.header_model.item(row).setCheckState(state)
    def cas_header_changed(self,row,state):
        self._CASbook_current_sheet.header_model.item(row).setCheckState(state)
    def comparison_append_list_changed(self,row,state):
        self._comparison_append_model.item(row).setCheckState(state)
    def comparison_delete_list_changed(self,row,state):
        self._comparison_delete_model.item(row).setCheckState(state)

    ##################################################
    #       Recover
    ##################################################
    def recover_ps_sheet_selected(self):
        self.refresh_ps_header_selected(self._PSbook_current_sheet_idx)
        self.select_ps_sheet(self._PSbook_current_sheet_idx)
    def recover_cas_sheet_selected(self):
        self.refresh_cas_header_selected(self._CASbook_current_sheet_idx)
        self.select_cas_sheet(self._CASbook_current_sheet_idx)
    def store_ps_file(self,action):
        ps_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))
        self._PSbook.save_as(ps_file_name)
        self._PSstack.push(PsPack(action,ps_file_name+r'.xlsx'))
        self._PSbook_autosave_flag = True
        self.open_ps_by_bytesio(ps_file_name+r'.xlsx')
        self.recover_ps_sheet_selected()
    def store_ps_file_without_open(self,action):
        ps_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))
        self._PSbook.save_as(ps_file_name)
        self._PSstack.push(PsPack(action,ps_file_name))

    def store_cas_file(self,action):
        cas_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))
        self._CASbook.save_as(cas_file_name)
        self._CASstack.push(CasPack(action,cas_file_name+r'.xlsx'))
        self._CASbook_autosave_flag = True
        self.open_cas_by_bytesio(cas_file_name+r'.xlsx')
        self.recover_cas_sheet_selected()
    def store_cas_file_without_open(self,action):
        cas_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))
        self._CASbook.save_as(cas_file_name)
        self._CASstack.push(CasPack(action,cas_file_name))
    def copy_cas(self,filename):
        cas_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))+r'.xlsx'
        shutil.copy(filename,cas_file_name)
        self._CASstack = FileStack()
        self._CASstack.push(CasPack('original',cas_file_name))
        return cas_file_name
    def copy_ps(self,filename):
        ps_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))+r'.xlsx'
        shutil.copy(filename,ps_file_name)
        self._PSstack = FileStack()
        self._PSstack.push(PsPack('original',ps_file_name))
        return ps_file_name
    def undo_ps(self):
        self.animation_progressBar(0)
        f = self._PSstack.pop()
        if f != None:
            self._PSbook_autosave_flag = True
            self.open_ps_by_bytesio(f[1].file_name)
            #self.animation_progressBar(50)
            self.recover_ps_sheet_selected()
            self.animation_progressBar(100)
            self.refresh_msg('ps file revert action:%s'%f[0])
        else:
            self.refresh_msg('Already at oldest change')
            self.animation_progressBar(100)
            self.PSbook_modified = False
    def undo_cas(self):
        self.animation_progressBar(0)
        f = self._CASstack.pop()
        if f != None:
            self._CASbook_autosave_flag = True
            self.open_cas_by_bytesio(f[1].file_name)
            #self.animation_progressBar(50)
            self.recover_cas_sheet_selected()
            self.animation_progressBar(100)
            self.refresh_msg('cas file revert action:%s'%f[0])
        else:
            self.refresh_msg('Already at oldest change')
            self.animation_progressBar(100)
            self.CASbook_modified = False

    def select_extended_preview(self):
        self.refresh_extended_preview(self._PSbook_current_sheet.extended_preview_model_list)


    @property
    def CASbook_modified(self):
        return self._CASbook_modified
    @CASbook_modified.setter
    def CASbook_modified(self,value):
        self._CASbook_modified = value
        self._queue_wr.put(('refresh_CASbook_modified',value))
    @property
    def PSbook_modified(self):
        return self._PSbook_modified
    @PSbook_modified.setter
    def PSbook_modified(self,value):
        self._PSbook_modified = value
        self._queue_wr.put(('refresh_PSbook_modified',value))
         

    def refresh_cas_book_name(self,model):
        self._queue_wr.put(('refresh_cas_book_name',model))
    def refresh_ps_book_name(self,model):
        self._queue_wr.put(('refresh_ps_book_name',model))
    def refresh_cas_sheet_name(self,model):
        self._queue_wr.put(('refresh_cas_sheet_name',model))
    def refresh_ps_sheet_name(self,model):
        self._queue_wr.put(('refresh_ps_sheet_name',model))
    def refresh_preview(self,model):
        self._queue_wr.put(('refresh_preview',model))
    def refresh_ps_header(self,model):
        self._queue_wr.put(('refresh_ps_header',model))
    def refresh_cas_header(self,model):
        self._queue_wr.put(('refresh_cas_header',model))
    def refresh_comparison_delete_list(self,model):
        self._queue_wr.put(('refresh_comparison_delete_list',model))
    def refresh_comparison_append_list(self,model):
        self._queue_wr.put(('refresh_comparison_append_list',model))
    def refresh_msg(self,model):
        self._queue_wr.put(('refresh_msg',model))
        logging.info(str(model))
    def refresh_warning(self,model):
        self._queue_wr.put(('refresh_warning',model))
        logging.info(str(model))
    def refresh_dialog(self,model):
        self._queue_wr.put(('refresh_dialog',model))
        logging.info(str(model))
    def refresh_selected_cell(self,model):
        self._queue_wr.put(('refresh_selected_cell',model))
    def refresh_progressBar(self,model):
        self._queue_wr.put(('refresh_progressBar',model))
    def animation_progressBar(self,model):
        self._queue_wr.put(('animation_progressBar',model))
    def refresh_ps_header_selected(self,model):
        self._queue_wr.put(('refresh_ps_header_selected',model))
    def refresh_cas_header_selected(self,model):
        self._queue_wr.put(('refresh_cas_header_selected',model))
    def refresh_extended_preview(self,model):
        self._queue_wr.put(('refresh_extended_preview',model))

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
    @property
    def col_letter(self):
        return self._cell.col_letter

