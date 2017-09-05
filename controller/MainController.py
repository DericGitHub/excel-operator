# -*- coding: utf-8 -*- 
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
import Queue
#def pt(step):
#    pass
mc = 0
def pt(step):
    global mc 
    if mc == 0:
        mc = time.time()
    print 'step %s:takes %s'%(step,time.time()-mc)
    mc = time.time()
class MainControllerUI(QObject):
    open_cas_signal = pyqtSignal(str)
    open_ps_signal = pyqtSignal(str)
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
        while self._status == True:
            if not self._queue_rd.empty():
                task = self._queue_rd.get()
                print 'ui got task:%s'%task[0]
                if task[0] == r'refresh_cas_book_name':
                    self._window.update_cas_file(task[1])
                elif task[0] == r'refresh_ps_book_name':
                    self._window.update_ps_file(task[1])
                elif task[0] == r'refresh_cas_sheet_name':
                    self._window.update_cas_sheets(task[1])
                elif task[0] == r'refresh_ps_sheet_name':
                    self._window.update_ps_sheets(task[1])
                elif task[0] == r'refresh_preview':
                    self._window.update_preview(task[1])
                elif task[0] == r'refresh_ps_header':
                    self._window.update_ps_header(task[1])
                elif task[0] == r'refresh_cas_header':
                    self._window.update_cas_header(task[1])
                elif task[0] == r'refresh_comparison_delete_list':
                    self._window.update_comparison_delete_list(task[1])
                elif task[0] == r'refresh_comparison_append_list':
                    self._window.update_comparison_append_list(task[1])
                elif task[0] == r'refresh_message':
                    self._window.update_message(task[1])
                elif task[0] == r'refresh_msg':
                    self._window.update_msg(task[1])
                elif task[0] == r'refresh_selected_cell':
                    self._window.update_selected_cell(task[1])
                elif task[0] == r'refresh_progressBar':
                    self._window.update_progressBar(task[1])
                elif task[0] == r'animation_progressBar':
                    self.animation_progressBar(task[1])
                elif task[0] == r'refresh_ps_header_selected':
                    self._window.update_ps_header_selected(task[1])
                elif task[0] == r'refresh_cas_header_selected':
                    self._window.update_cas_header_selected(task[1])
    def __del__(self):
        self._worker.wait()
        self._worker.quit()
    def init_GUI(self):
        self._application = QApplication(sys.argv)
        self._window = Window.Window()
    def init_worker(self):
        #self._worker.finished.connect(self._worker.wait())
        self.bind_GUI_event()
        #self.bind_worker_event(self._worker)

    def bind_GUI_event(self):
        self._window.bind_open_cas(self.open_cas)
        #self._window.bind_open_cas(worker.open_cas)
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
    @pyqtSlot()
    def open_cas(self):
        #########################
        #   Open cas
        #########################
#        try:
#            filename = Window.open_file_dialog()
#            self.open_cas_by_name(filename)
#        except:
#            filename = None
        if self._CASbook_modified == True:
            if self._window.open_file_confirm() == True:
                filename = Window.open_file_dialog()
                self._queue_wr.put((r'open_cas',str(filename)))
        else:
            filename = Window.open_file_dialog()
            self._queue_wr.put((r'open_cas',str(filename)))
    @pyqtSlot()
    def open_ps(self):
        #########################
        #   Open ps
        #########################
#        try:
#            filename = Window.open_file_dialog()
#            self.open_ps_by_name(str(filename))
#        except:
#            filename = None
        if self._PSbook_modified == True:
            if self._window.open_file_confirm() == True:
                filename = Window.open_file_dialog()
                self._queue_wr.put((r'open_ps',str(filename)))
        else:
            filename = Window.open_file_dialog()
            self._queue_wr.put((r'open_ps',str(filename)))
    @pyqtSlot()
    def save_cas(self):
        self._queue_wr.put((r'save_cas',))
    @pyqtSlot()
    def save_ps(self):
        self._queue_wr.put((r'save_ps',))
    @pyqtSlot()
    def saveas_cas(self):
        filename = Window.save_file_dialog()
        self._queue_wr.put((r'saveas_cas',str(filename)))
    @pyqtSlot()
    def saveas_ps(self):
        filename = Window.save_file_dialog()
        self._queue_wr.put((r'saveas_ps',str(filename)))
    @pyqtSlot(int)
    def select_cas_sheet(self,index):
        self._queue_wr.put((r'select_cas_sheet',index))
    @pyqtSlot(int)
    def select_ps_sheet(self,index):
        self._queue_wr.put((r'select_ps_sheet',index))
    @pyqtSlot(int)
    def select_preview(self,index):
        self._queue_wr.put((r'select_preview',index))
    @pyqtSlot()
    def select_sync_ps_to_cas(self):
        self._queue_wr.put((r'select_sync_ps_to_cas',))
    @pyqtSlot()
    def select_sync_cas_to_ps(self):
        self._queue_wr.put((r'select_sync_cas_to_ps',))
    @pyqtSlot(bool)
    def select_sync_select_all_ps_headers(self,state):
        self._queue_wr.put((r'select_sync_select_all_ps_headers',state))
    @pyqtSlot(bool)
    def select_sync_select_all_cas_headers(self,state):
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
        self._queue_wr.put((r'comparison_select_all_delete',state))
    @pyqtSlot(bool)
    def comparison_select_all_append(self,state):
        self._queue_wr.put((r'comparison_select_all_append',state))
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

#    @pyqtSlot(str)
#    def refresh_cas_book_name(self,model):
#        self._window.update_cas_file(model)
#    @pyqtSlot(str)
#    def refresh_ps_book_name(self,model):
#        self._window.update_ps_file(model)
#    @pyqtSlot(QStandardItemModel)
#    def refresh_cas_sheet_name(self,model):
#        self._window.update_cas_sheets(model)
#    @pyqtSlot(QStandardItemModel)
#    def refresh_ps_sheet_name(self,model):
#        self._window.update_ps_sheets(model)
#    @pyqtSlot(QStandardItemModel)
#    def refresh_preview(self,model):
#        self._window.update_preview(model)
#    @pyqtSlot(QStandardItemModel)
#    def refresh_ps_header(self,model):
#        self._window.update_ps_header(model)
#    @pyqtSlot(QStandardItemModel)
#    def refresh_cas_header(self,model):
#        self._window.update_cas_header(model)
#    @pyqtSlot(QStandardItemModel)
#    def refresh_comparison_delete_list(self,model):
#        self._window.update_comparison_delete_list(model)
#    @pyqtSlot(QStandardItemModel)
#    def refresh_comparison_append_list(self,model):
#        self._window.update_comparison_append_list(model)
#    @pyqtSlot(str)
#    def refresh_message(self,model):
#        self._window.update_message(model)
#    @pyqtSlot(str)
#    def refresh_msg(self,model):
#        self._window.update_msg(model)
#        logging.info(str(model))
#    @pyqtSlot(list)
#    def refresh_selected_cell(self,model):
#        self._window.update_selected_cell(model)
#    @pyqtSlot(int)
#    def refresh_progressBar(self,model):
#        self._window.update_progressBar(model)
    @pyqtSlot(int)
    def animation_progressBar(self,model):
        if self._progressBar_status < model:
            while self._progressBar_status < model:
                self._progressBar_status += 0.002
                #self.refresh_progressBar(self._progressBar_status)
                self._window.update_progressBar(self._progressBar_status)
        else:
            self._progressBar_status = 0
            #self.refresh_progressBar(self._progressBar_status)
            self._window.update_progressBar(self._progressBar_status)
            while self._progressBar_status < model:
                self._progressBar_status += 0.002
                #self.refresh_progressBar(self._progressBar_status)
                self._window.update_progressBar(self._progressBar_status)
        #self._window.bind_select_extended_preview(worker.select_extended_preview)
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
        worker.signal_refresh_selected_cell.connect(self._window.update_selected_cell)
        worker.signal_refresh_progressBar.connect(self._window.update_progressBar)
        worker.signal_animation_progressBar.connect(self.animation_progressBar)
    def show_GUI(self):
        self._window.show()
        sys.exit(self._application.exec_())       
class MainControllerWoker(QThread):
    def __init__(self,parent=None):
        super(MainControllerWoker,self).__init__(parent)
        self.instance = MainController()
    def run(self):
        pass
##################################################
#       class for handling application logic
##################################################
class MainController(object):
    signal_refresh_cas_book_name = pyqtSignal(str)
    signal_refresh_ps_book_name = pyqtSignal(str)
    signal_refresh_cas_sheet_name = pyqtSignal(QStandardItemModel)
    signal_refresh_ps_sheet_name = pyqtSignal(QStandardItemModel)
    signal_refresh_preview = pyqtSignal(QStandardItemModel)
    signal_refresh_ps_header = pyqtSignal(QStandardItemModel)
    signal_refresh_cas_header = pyqtSignal(QStandardItemModel)
    signal_refresh_comparison_delete_list = pyqtSignal(QStandardItemModel)
    signal_refresh_comparison_append_list = pyqtSignal(QStandardItemModel)
    signal_refresh_message = pyqtSignal(str)
    signal_refresh_msg = pyqtSignal(str)
    signal_refresh_selected_cell = pyqtSignal(list)
    signal_refresh_progressBar = pyqtSignal(int)
    signal_animation_progressBar = pyqtSignal(int)
    signal_refresh_ps_header_selected = pyqtSignal(int)
    signal_refresh_cas_header_selected = pyqtSignal(int)
    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,queue_wr=None,queue_rd=None):
#        self._application = None
        super(MainController,self).__init__()
        self._queue = Queue.Queue(maxsize=1)
        self._status = True
        self._queue_wr = queue_wr
        self._queue_rd = queue_rd
    def send(self,task):
        if not self._queue.full():
            self._queue.put(task)
    def run(self):
        import pythoncom
        pythoncom.CoInitialize() 
        self._xw_app = None
        self._window = None
        self._PSbook = None
        self._PSbook_name = ''
        self._PSbook_sheets = QStandardItemModel()
        self._PSbook_current_sheet = None
        self._PSbook_current_sheet_idx = 0
        self._PSbook_current_sheet_name = None
        self._PSbook_autosave_flag = False
        self._PSbook_modified = False
        self._CASbook = None
        self._CASbook_wr = None
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
        self.init_logging()
        self.init_tmp_directory()
        self.start_xlwings_app()
        self.init_model()
        self.init_file_stack()
        self.go_to_open_cas_filename = None
        self.go_to_open_ps_filename = None
        self.go_to_save_cas_filename = None
        self.go_to_save_ps_filename = None
        while self._status == True:
            if not self._queue_rd.empty():
                task = self._queue_rd.get()
                print 'controller got task:%s'%task[0]
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
                    self.select_preview(task[1])
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


#        self.init_GUI()
#        self.bind_GUI_event()
#        self.show_GUI()
    def __del__(self):
        #del self._CASbook
        #del self._PSbook
        self._xw_app.quit()
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
        self._application.exec_()
        #sys.exit(self._application.exec_())       
    def start_xlwings_app(self):
        self._xw_app = xw.App(visible=False)
    ##################################################
    #       Actions
    ##################################################

    
    @pyqtSlot()
    def open_cas(self):
        #########################
        #   Open cas
        #########################
#        try:
#            filename = Window.open_file_dialog()
#            self.open_cas_by_name(filename)
#        except:
#            filename = None
        if self._CASbook_modified == True:
            if self._window.open_file_confirm() == True:
                filename = Window.open_file_dialog()
                self.open_cas_by_name(str(filename))
        else:
            filename = Window.open_file_dialog()
            self.open_cas_by_name(str(filename))
            
            
    def open_cas_by_name(self,filename):
        self._CASbook = CASbook.CASbook(filename,self._xw_app)
        self._CASbook_wr = self._CASbook.workbook_wr
        self._CASbook_name = filename
        #########################
        #   Update model
        #########################
        self.init_model()
        self._CASstack = FileStack()
        self._CASbook.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_cas_book_name(self._CASbook.workbook_name)
        self.refresh_cas_sheet_name(self._CASbook.sheet_name_model)
        self.refresh_msg('open cas file:%s'%self._CASbook_name)
        self.store_cas_file_without_open('original')
        self._CASbook_modified = False
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
        self._CASbook_modified = False
        #self._window.update_cas_file(self._CASbook_name)
        #self._window.update_cas_sheets(self._CASbook.sheets_name)
         
    @pyqtSlot()
    def open_ps(self):
        #########################
        #   Open ps
        #########################
#        try:
#            filename = Window.open_file_dialog()
#            self.open_ps_by_name(str(filename))
#        except:
#            filename = None
        if self._PSbook_modified == True:
            if self._window.open_file_confirm() == True:
                filename = Window.open_file_dialog()
                self.open_ps_by_name(str(filename))
        else:
            filename = Window.open_file_dialog()
            self.open_ps_by_name(str(filename))
#        if self._PSbook_modified == True:
#            pass
#        else:
#            pt(1)
#            filename = Window.open_file_dialog()
#            pt(2)
#            self.open_ps_by_name(str(filename))
#            pt(3)
    def open_ps_by_name(self,filename):
        print 'case 1'
        self._PSbook = PSbook.PSbook(filename,self._xw_app)
        print 'case 2'
        #try:
        #    self._PSbook = PSbook.PSbook(filename)
        #    #print "open succeed"
        #except:
        #    #print "not a excel"
        self._PSbook_name = filename
        #########################
        #   Update model
        #########################
        self.init_model()
        self._PSstack = FileStack()
        self._PSbook.update_model()
        #self._window.update_ps_file(self._PSbook_name)
        #self._window.update_ps_sheets(self._PSbook.sheets_name)
        #########################
        #   Refresh UI
        #########################
        self.refresh_ps_book_name(self._PSbook.workbook_name)
        self.refresh_ps_sheet_name(self._PSbook.sheet_name_model)
        self.refresh_msg('open ps file:%s'%self._PSbook_name)
        self.store_ps_file_without_open('original')
        self._PSbook_modified = False

    @pyqtSlot()
    def open_ps_by_bytesio(self,bytesio):
        self._PSbook = PSbook.PSbook(bytesio,self._xw_app)
        #try:
        #    self._PSbook = PSbook.PSbook(bytesio)
        #    #print "open succeed"
        #except:
        #    #print "not a excel"
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
        self._PSbook_modified = False


    @pyqtSlot()
    def save_cas(self):
        self._CASbook.save_as(self._CASbook_name)
        ##self._CASbook.workbook.save(self._CASbook_name)
        #self.open_cas_by_bytesio(self._CASbook_name)
        #self.recover_cas_sheet_selected()
        self.refresh_message('saved cas file')
        self.refresh_msg('saved cas file:%s'%self._CASbook_name)
        self._CASbook_modified = False
    @pyqtSlot()
    def save_ps(self):
        self._PSbook.save_as(self._PSbook_name)
        ##self._PSbook.workbook.save(self._PSbook_name)
        #self.open_ps_by_bytesio(self._PSbook_name)
        #self.recover_ps_sheet_selected()
        self.refresh_message('saved ps file')
        self.refresh_msg('saved ps file:%s'%self._PSbook_name)
        self._PSbook_modified = False
#    def save_ps(self):
#        self._PSbook_autosave_flag = True
#        self._PSbook.save_as(self._PSbook_name)
#        self.open_ps_by_bytesio(self._PSbook_name)
#        self.recover_ps_sheet_selected()
#        self.refresh_message('saved ps file')
        
    @pyqtSlot()
    def saveas_cas(self,fileName):
#    def saveas_cas(self):
#        fileName = Window.save_file_dialog()
        self._CASbook.workbook_wr.save(str(fileName))
        #self._CASbook = CASbook.CASbook(str(fileName),self._xw_app)
        #self._CASbook_wr = self._CASbook.workbook_wr
        #self._CASbook_name = fileName
        #self._CASstack = FileStack()
        ##########################
        ##   Update model
        ##########################
        #self.init_model()
        #self._CASbook.update_model()
        ##########################
        ##   Refresh UI
        ##########################
        #self.refresh_cas_book_name(self._CASbook.workbook_name)
        #self.refresh_cas_sheet_name(self._CASbook.sheet_name_model)

        ##self._CASbook.save_as(str(fileName))
        self.refresh_message('save cas to %s'%fileName)
        self.refresh_msg('saved cas file:%s'%self._CASbook_name)
        self._CASbook_modified = False
    @pyqtSlot()
    def saveas_ps(self,fileName):
#    def saveas_ps(self):
#        #'''
#        #Solution 1
#        #'''
#        fileName = Window.save_file_dialog()
        self._PSbook.save_as(str(fileName))
        #self._PSbook = PSbook.PSbook(str(fileName),self._xw_app)
        #self._PSbook_name = fileName
        #self._PSstack = FileStack()
        ##########################
        ##   Update model
        ##########################
        #self.init_model()
        #self._PSbook.update_model()
        ##########################
        ##   Refresh UI
        ##########################
        #self.refresh_ps_book_name(self._PSbook.workbook_name)
        #self.refresh_ps_sheet_name(self._PSbook.sheet_name_model)
        self.refresh_message('save ps to %s'%fileName)
        self.refresh_msg('saved ps file:%s'%self._PSbook_name)
        self._PSbook_modified = False
        #Solution 2
        #fileName = Window.save_file_dialog()
        #self._PSbook.save_as(str(fileName))
        #self.refresh_message('save ps to %s'%fileName)
        

    @pyqtSlot(int)
    def select_cas_sheet(self,sheet_idx):
        #print 'auto_save = %s'%self._CASbook_autosave_flag
        if self._CASbook_autosave_flag != True:
            if self._CASbook_current_sheet_idx != sheet_idx:
                self.refresh_msg('select cas sheet:%s'%self._CASbook.workbook.sheetnames[sheet_idx])
            self._CASbook_current_sheet_idx = sheet_idx
        else:
            self._CASbook_autosave_flag = False
        #print 'stored idx %d'%self._CASbook_current_sheet_idx
        #print 'select sheet %d'%sheet_idx

        #self._CASbook_current_sheet_name = self._CASbook.sheets_name[sheet_idx]
        #print self._CASbook_current_sheet_name
        self._CASbook_current_sheet = self._CASbook.sheets[sheet_idx]
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
        #print 'cas_sheet :%s'%self._CASbook_current_sheet
    @pyqtSlot(int)
    def select_ps_sheet(self,sheet_idx):
        #print 'auto_save = %s'%self._PSbook_autosave_flag
        if self._PSbook_autosave_flag != True:
            if self._PSbook_current_sheet_idx != sheet_idx:
                self.refresh_msg('select ps sheet:%s'%self._PSbook.workbook.sheetnames[sheet_idx])
            self._PSbook_current_sheet_idx = sheet_idx
        else:
            self._PSbook_autosave_flag = False
        #print 'stored idx %d'%self._PSbook_current_sheet_idx
        #print 'select sheet %d'%sheet_idx
        #self._PSbook_current_sheet_name = self._PSbook.sheets_name[sheet_idx]
        #print self._PSbook_current_sheet_name
        self._PSbook_current_sheet = self._PSbook.sheets[sheet_idx]
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
        self.refresh_message('select ps sheet:%s'%self._PSbook_current_sheet_idx)
        #print 'ps_sheet :%s'%self._PSbook_current_sheet
    @pyqtSlot(int)
    def select_preview(self,index):
        #print self._PSbook_current_sheet._preview_model.itemFromIndex(index).cell.value
        self._preview_selected_cell = self._PSbook_current_sheet._preview_model.itemFromIndex(index).cell
        #########################
        #   Refresh UI
        #########################
        self.refresh_message(self._PSbook_current_sheet._preview_model.itemFromIndex(index).cell.value)
        self.refresh_selected_cell([self._preview_selected_cell.row,self._preview_selected_cell.col])
        self.refresh_msg('select cell:row %s,column %s'%(self._preview_selected_cell.row,self._preview_selected_cell.col))

    @pyqtSlot()
    def select_extended_preview(self,index):
        pass
    @pyqtSlot(bool)
    def select_sync_select_all_ps_headers(self,state):
        if state == Qt.Checked:
            self._PSbook_current_sheet.select_all_headers()
            self.refresh_msg('select all ps headers')
        elif state == Qt.Unchecked:
            self._PSbook_current_sheet.unselect_all_headers()
            self.refresh_msg('unselect all ps headers')
    @pyqtSlot(bool)
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
#        self._worker = OperateThread(self.select_sync_ps_to_cas)
#        self._worker.finished.connect(self.select_sync_ps_to_cas_complete)
#        self._worker.start()
#    def select_sync_ps_to_cas_complete(self):
#        self._worker.quit()
#        del self._worker
#        #########################
#        #   Refresh UI
#        #########################
#        self.refresh_message('sync ps to cas done')
#        self.refresh_msg('sync ps to cas done')
#        self.animation_progressBar(100)
#        self._CASbook_modified = True
# 
#        
#                
#                
#    def select_sync_ps_to_cas_worker(self):
        #########################
        #   Data sync
        #########################
        print 'sync start'
        xml_names_row = []
        headers_column = []
        sync_list = []
        # pick out all xmlnames in cas file
        xml_names_cas = self._CASbook_current_sheet.xml_names()
        headers_ps = self._PSbook_current_sheet.checked_headers()
        
        self.animation_progressBar(0)
        for xml_name_cas in xml_names_cas:
            # there are some white space row in xmlname of cas file
            if xml_name_cas.value == None:
                #print 'xml_name_cas is None'
                continue
            #pick out specified xmlname in ps file
            xml_name_ps = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_cas.value)
            if xml_name_ps == None:
                #print 'could not find %s in ps file'%xml_name_cas.value
                continue
            xml_names_row.append((xml_name_ps.row,xml_name_cas.row))
            #sync_list.append((xml_name_ps,xml_name_cas))
        self.animation_progressBar(33)
            
        for header_ps in headers_ps:
            if header_ps.value == None:
                continue
            #print '%s'%header_ps.value
            header_cas = self._CASbook_current_sheet.search_header_by_value(header_ps.value)
            if header_cas == None:
                #print 'could not find %s in cas file'%header_cas.value
                continue
            headers_column.append((header_ps.column,header_cas.column))
            #sync_list.append((xml_name_ps,header_ps,xml_name_cas,header_cas))
        self.animation_progressBar(66)        
        
        for rows in xml_names_row:
            for columns in headers_column:
                self._CASbook_current_sheet.cell_wr(rows[1],columns[1]).value = self._PSbook_current_sheet.cell(rows[0],columns[0]).value

#                source_item = self._PSbook_current_sheet.cell(xml_name_ps.row,header_ps.col)
#                target_item = self._CASbook_current_sheet.cell(xml_name_cas.row,header_cas.col)
#                target_item.value = source_item.value
        #########################
        #   Update model
        #########################
        self._CASbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
#        self.refresh_cas_header(self._CASbook_current_sheet.header_model)
        self.store_cas_file('sync ps to cas')
        print 'sync completed'
#        self.refresh_message('sync ps to cas done')
#        self.refresh_msg('sync ps to cas done')
#        self.animation_progressBar(100)
#        self._CASbook_modified = True
#        for pair in sync_list:
#            #source_item = pair[0].get_item_by_header(pair[1])
#            #target_item = pair[2].get_item_by_header(pair[3])
#            source_item = self._PSbook_current_sheet.cell(pair[0].row,pair[1].col)
#            target_item = self._CASbook_current_sheet.cell(pair[2].row,pair[3].col)
#            #print '<<<<<source>>>>>ps:header=%s,xmlname=%s,value=%s,position:row %s,col %s'%(source_item.header.value,source_item.xmlname.value,source_item.value,source_item.row,source_item.col)
#            #print '>>>>>target<<<<<cas:header=%s,xmlname=%s,value=%s,position:row %s,col %s'%(target_item.header.value,target_item.xmlname.value,target_item.value,target_item.row,target_item.col)
#            target_item.value = source_item.value
#            #self._CASbook_current_sheet.cell(pair[2].row,pair[3].col).value = self._PSbook_current_sheet.cell(pair[0].row,pair[1].col).value
#        #print 'sync ps to cas done'

#    def select_sync_ps_to_cas(self):
#        #########################
#        #   Data sync
#        #########################
#        xml_names_ps = []
#        headers_cas = []
#        sync_list = []
#        # pick out all xmlnames in cas file
#        xml_names_cas = self._CASbook_current_sheet.xml_names()
#        headers_ps = self._PSbook_current_sheet.checked_headers()
#        
#        for xml_name_cas in xml_names_cas:
#            # there are some white space row in xmlname of cas file
#            if xml_name_cas.value == None:
#                #print 'xml_name_cas is None'
#                continue
#            #pick out specified xmlname in ps file
#            xml_name_ps = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_cas.value)
#            if xml_name_ps == None:
#                #print 'could not find %s in ps file'%xml_name_cas.value
#                continue
#            #xml_names_ps.append(xml_name_ps)
#            #sync_list.append((xml_name_ps,xml_name_cas))
#            
#            for header_ps in headers_ps:
#                #print '%s'%header_ps.value
#                header_cas = self._CASbook_current_sheet.search_header_by_value(header_ps.value)
#                if header_cas == None:
#                    #print 'could not find %s in cas file'%header_cas.value
#                    continue
#                #headers_cas.append(header_cas)
#                #sync_list.append((xml_name_ps,header_ps,xml_name_cas,header_cas))
#                
#                source_item = self._PSbook_current_sheet.cell(xml_name_ps.row,header_ps.col)
#                target_item = self._CASbook_current_sheet.cell(xml_name_cas.row,header_cas.col)
#                target_item.value = source_item.value
#        #########################
#        #   Update model
#        #########################
#        self._CASbook_current_sheet.update_model()
#        #########################
#        #   Refresh UI
#        #########################
#        self.refresh_cas_header(self._CASbook_current_sheet.header_model)
#        self.store_cas_file('sync ps to cas')
#        self.refresh_message('sync ps to cas done')
##        for pair in sync_list:
##            #source_item = pair[0].get_item_by_header(pair[1])
##            #target_item = pair[2].get_item_by_header(pair[3])
##            source_item = self._PSbook_current_sheet.cell(pair[0].row,pair[1].col)
##            target_item = self._CASbook_current_sheet.cell(pair[2].row,pair[3].col)
##            #print '<<<<<source>>>>>ps:header=%s,xmlname=%s,value=%s,position:row %s,col %s'%(source_item.header.value,source_item.xmlname.value,source_item.value,source_item.row,source_item.col)
##            #print '>>>>>target<<<<<cas:header=%s,xmlname=%s,value=%s,position:row %s,col %s'%(target_item.header.value,target_item.xmlname.value,target_item.value,target_item.row,target_item.col)
##            target_item.value = source_item.value
##            #self._CASbook_current_sheet.cell(pair[2].row,pair[3].col).value = self._PSbook_current_sheet.cell(pair[0].row,pair[1].col).value
##        #print 'sync ps to cas done'
#
    @pyqtSlot()
    def select_sync_cas_to_ps(self):
        print 'sync cas to ps'
        #########################
        #   Data sync
        #########################
        xml_names_row = []
        headers_column = []
        headers_ps = []
        sync_list = []
        # pick out all xmlnames in cas file
        pt('sync1')
        xml_names_cas = self._CASbook_current_sheet.xml_names()
        headers_cas = self._CASbook_current_sheet.checked_headers()
        
        pt('sync2')
        self.animation_progressBar(0)
        for xml_name_cas in xml_names_cas:
            # there are some white space row in xmlname of cas file
            if xml_name_cas.value == None:
                #print 'xml_name_cas is None'
                continue
            #pick out spcified xmlname in ps file
            xml_name_ps = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_cas.value)
            if xml_name_ps == None:
                #print 'could not find %s in ps file'%xml_name_cas.value
                continue
            xml_names_row.append((xml_name_ps.row,xml_name_cas.row))
        self.animation_progressBar(33)
            
        pt('sync3')
        for header_cas in headers_cas:
            header_ps = self._PSbook_current_sheet.search_header_by_value(header_cas.value)
            if header_ps == None:
                #print 'could not find %s in cas file'%header_ps.value
                continue
            headers_column.append((header_ps.column,header_cas.column))
        self.animation_progressBar(66)
        pt('sync4')
        for columns in headers_column:
            for rows in xml_names_row:
                self._PSbook_current_sheet.cell_wr(rows[0],columns[0]).value = self._CASbook_current_sheet.cell(rows[1],columns[1]).value
        #self._PSbook_current_sheet.auto_fit([columns[0] for columns in headers_column])
        pt('sync5')
        #########################
        #   Update model
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        pt('sync6')
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_ps_header(self._PSbook_current_sheet.header_model)
        #self.store_ps_file('sync cas to ps',self._PSbook.virtual_workbook)
        self.store_ps_file('sync cas to ps')
        pt('sync7')
        self.refresh_message('sync cas to ps done')
        self.refresh_msg('sync cas to ps done')
        self.animation_progressBar(100)
        self._PSbook_modified = True
#    def select_sync_cas_to_ps(self):
#        #########################
#        #   Data sync
#        #########################
#        xml_names_ps = []
#        headers_cas = []
#        headers_ps = []
#        sync_list = []
#        # pick out all xmlnames in cas file
#        xml_names_cas = self._CASbook_current_sheet.xml_names()
#        headers_cas = self._CASbook_current_sheet.checked_headers()
#        
#        for xml_name_cas in xml_names_cas:
#            # there are some white space row in xmlname of cas file
#            if xml_name_cas.value == None:
#                #print 'xml_name_cas is None'
#                continue
#            #pick out spcified xmlname in ps file
#            xml_name_ps = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_cas.value)
#            if xml_name_ps == None:
#                #print 'could not find %s in ps file'%xml_name_cas.value
#                continue
#            
#            for header_cas in headers_cas:
#                header_ps = self._PSbook_current_sheet.search_header_by_value(header_cas.value)
#                if header_ps == None:
#                    #print 'could not find %s in cas file'%header_ps.value
#                    continue
#                headers_ps.append(header_ps)
#                sync_list.append((xml_name_cas,header_cas,xml_name_ps,header_ps))
#        for pair in sync_list:
#            #source_item = pair[0].get_item_by_header(pair[1])
#            #print source_item._cell.col_idx
#            #target_item = pair[2].get_item_by_header(pair[3])
#            source_item = self._CASbook_current_sheet.cell(pair[0].row,pair[1].col)
#            target_item = self._PSbook_current_sheet.cell(pair[2].row,pair[3].col)
#            #print '<<<<<source>>>>>cas:header=%s,xmlname=%s,value=%s,position:row %d,col %d'%(source_item.header.value,\
#            #        source_item.xmlname.value,\
#            #        source_item.value,\
#            #        source_item.row,\
#            #        source_item.col)
#            #print '>>>>>target<<<<<ps:header=%s,xmlname=%s,value=%s,position:row %d,col %d'%(target_item.header.value,target_item.xmlname.value,target_item.value,target_item.row,target_item.col)
#            target_item.value = source_item.value
#            #self._PSbook_current_sheet.cell(pair[2].row,pair[3].col).value = self._CASbook_current_sheet.cell(pair[0].row,pair[1].col).value
#            # do a copy style
#            alignment = copy(target_item.alignment)
#            alignment.wrapText = True
#            target_item.alignment = alignment
#        # adjust column width
#        for header_ps in headers_ps:
#            self._PSbook_current_sheet._worksheet.column_dimensions[header_ps.col_letter].width = 25
#        #print 'sync cas to ps done'
#        #########################
#        #   Update model
#        #########################
#        self._PSbook_current_sheet.update_model()
#        #########################
#        #   Refresh UI
#        #########################
#        self.refresh_preview(self._PSbook_current_sheet.preview_model)
#        #self.store_ps_file('sync cas to ps',self._PSbook.virtual_workbook)
#        self.store_ps_file('sync cas to ps')
#        self.refresh_message('sync cas to ps done')
    ##################################################
    #       Comparison
    ##################################################
    @pyqtSlot()
    def comparison_start(self):
        self.animation_progressBar(0)
        self.init_model()
        try:
            if self._PSbook_current_sheet != None and self._CASbook_current_sheet != None and self._PSbook_current_sheet.xmlname != None and self._CASbook_current_sheet.xmlname != None:
                append_list = list(set(self._CASbook_current_sheet.xml_names_value()).difference(set(self._PSbook_current_sheet.xml_names_value())))
                delete_list = list(set(self._PSbook_current_sheet.xml_names_value()).difference(set(self._CASbook_current_sheet.xml_names_value())))
                for xml_name_value in append_list:
                    xml_name = self._CASbook_current_sheet.search_xmlname_by_value(xml_name_value)
                    item_append = QComparisonItem(xml_name)
                    item_append.setCheckState(Qt.Unchecked)
                    item_append.setCheckable(True)
                    self._comparison_append_model.appendRow(item_append)
                self.animation_progressBar(40)
                for xml_name_value in delete_list:
                    xml_name = self._PSbook_current_sheet.search_xmlname_by_value(xml_name_value)
                    item_delete = QComparisonItem(xml_name)
                    item_delete.setCheckState(Qt.Unchecked)
                    item_delete.setCheckable(True)
                    self._comparison_delete_model.appendRow(item_delete)
                self.animation_progressBar(80)
        finally:
            self.refresh_comparison_append_list(self._comparison_append_model)
            self.refresh_comparison_delete_list(self._comparison_delete_model)
            self.refresh_message('comparison done')
            #self.refresh_msg('comparison done')
            self.animation_progressBar(100)
        
                
            
    @pyqtSlot()
    def comparison_delete(self):
        self.animation_progressBar(0)
        #########################
        #   Data operation
        #########################
        for delete_item in self.checked_delete():
            self._PSbook_current_sheet.delete_row(delete_item.cell.row,1)
            self._comparison_delete_model.removeRow(delete_item.row())
        self.animation_progressBar(80)
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_comparison_delete_list(self._comparison_delete_model)
        #self.store_ps_file('comparison delete',self._PSbook.virtual_workbook)
        self.store_ps_file('comparison delete')
        self.comparison_start()
        self.refresh_message('comparison delete done')
        self.refresh_msg('comparison delete done')
        self.animation_progressBar(100)
        self._PSbook_modified = True
        #print 'comparison delete done'
    @pyqtSlot()
    def comparison_append(self):
        self.animation_progressBar(0)
        #########################
        #   Data operation
        #########################
        if self._preview_selected_cell != None:
            self._PSbook_current_sheet.add_row(self._preview_selected_cell.row,self.checked_append_count(),PSsheet.DOWN)
            overwrite_row = self._preview_selected_cell.row
            for append_item in self.checked_append():
                overwrite_row += 1
                self._PSbook_current_sheet.cell_wr(overwrite_row,self._PSbook_current_sheet.xmlname.col).value = append_item.value
                self._comparison_append_model.removeRow(append_item.row())
        self.animation_progressBar(80)
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_comparison_append_list(self._comparison_append_model)
        #self.store_ps_file('comparison append',self._PSbook.virtual_workbook)
        self.store_ps_file('comparison append')
        self.comparison_start()
        self.refresh_message('comparison append done')
        self.refresh_msg('comparison append done')
        self.animation_progressBar(100)
        self._PSbook_modified = True
        #print 'comparison append done'
        #for append_item in self.checked_append():
        #    self._PSbook_current_sheet.add_row(
    @pyqtSlot(bool)
    def comparison_select_all_delete(self,state):
        for i in range(self._comparison_delete_model.rowCount()):
            item = self._comparison_delete_model.item(i)
            item.setCheckState(state)
        if state == Qt.Checked:
            self.refresh_msg('select all delete items')
        else:
            self.refresh_msg('unselect all delete items')
    @pyqtSlot(bool)
    def comparison_select_all_append(self,state):
        for i in range(self._comparison_append_model.rowCount()):
            item = self._comparison_append_model.item(i)
            item.setCheckState(state)
        if state == Qt.Checked:
            self.refresh_msg('select all append items')
        else:
            self.refresh_msg('unselect all append items')
    @pyqtSlot()
    def checked_delete(self):
        items = []
        for i in range(self._comparison_delete_model.rowCount()):
            item = self._comparison_delete_model.item(i)
            if item.checkState() == Qt.Checked:
                items.append(item)
        items.reverse()
        self.refresh_msg('checked delete items:%s'%str(map(lambda x:x.value,items)))
        return items
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
    @pyqtSlot()
    def preview_add(self):
        self.animation_progressBar(0)
        #########################
        #   Data operation
        #########################
        pt(4)
        if self._preview_selected_cell != None:
            self._PSbook_current_sheet.add_row(self._preview_selected_cell.row,1,PSsheet.DOWN)
        pt(5)
        self.animation_progressBar(80)
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        pt(6)
        #########################
        #   Refresh UI
        #########################
        pt(7)
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        pt(8)
        self.refresh_message('added one row below row %d'%self._preview_selected_cell.row)
        pt(9)
        self.refresh_msg('added one row below row %d'%self._preview_selected_cell.row)
        #self.store_ps_file('add',self._PSbook.virtual_workbook)
        self.store_ps_file('add')
        self.animation_progressBar(100)
        self._PSbook_modified = True
        pt(10)

    @pyqtSlot()
    def preview_delete(self):
        self.animation_progressBar(0)
        #########################
        #   Data operation
        #########################
        if self._preview_selected_cell != None:
            self._PSbook_current_sheet.delete_row(self._preview_selected_cell.row,1)
        self.animation_progressBar(80)
        #########################
        #   Update model   
        #########################
        self._PSbook_current_sheet.update_model()
        #########################
        #   Refresh UI
        #########################
        self.refresh_preview(self._PSbook_current_sheet.preview_model)
        self.refresh_message('deleted row %d'%self._preview_selected_cell.row)
        self.refresh_msg('deleted row %d'%self._preview_selected_cell.row)
        #self.store_ps_file('delete',self._PSbook.virtual_workbook)
        self.store_ps_file('delete')
        self.animation_progressBar(100)
        self._PSbook_modified = True

    @pyqtSlot()
    def preview_lock(self):
        self.animation_progressBar(0)
        #for item in self._PSbook_current_sheet.extended_preview_model():
        self._PSbook_current_sheet.unlock_sheet()
        self._PSbook_current_sheet.unlock_all_cells()
        self.animation_progressBar(20)
        for status in self._PSbook_current_sheet.status():
            if status.value == 'POR':
                self._PSbook_current_sheet.lock_row(status.row,True)
        self.animation_progressBar(80)
        self._PSbook_current_sheet.lock_sheet()
        #self.store_ps_file('lock',self._PSbook.virtual_workbook)
        self.store_ps_file('lock')
        self.refresh_message('lock sheet done')
        self.refresh_msg('lock sheet done')
        self.animation_progressBar(100)
        self._PSbook_modified = True
    ##################################################
    #       Recover
    ##################################################
    def recover_ps_sheet_selected(self):
        #print 'recover to sheet %d'%self._PSbook_current_sheet_idx
        #self._window.update_ps_header_selected(self._PSbook_current_sheet_idx)
        self.signal_refresh_ps_header_selected.emit(self._PSbook_current_sheet_idx)
        self.select_ps_sheet(self._PSbook_current_sheet_idx)
    def recover_cas_sheet_selected(self):
        #print 'recover to sheet %d'%self._CASbook_current_sheet_idx
        #self._window.update_cas_header_selected(self._CASbook_current_sheet_idx)
        self.signal_refresh_cas_header_selected.emit(self._CASbook_current_sheet_idx)
        self.select_cas_sheet(self._CASbook_current_sheet_idx)
    def store_ps_file(self,action):
        ps_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))
        self._PSbook.save_as(ps_file_name)
        #self._PSbook.workbook.save(ps_file_name)
        self._PSstack.push(PsPack(action,ps_file_name))
        self._PSbook_autosave_flag = True
        self.open_ps_by_bytesio(ps_file_name+r'.xlsx')
        self.recover_ps_sheet_selected()
    def store_ps_file_without_open(self,action):
        ps_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))
        self._PSbook.save_as(ps_file_name)
        self._PSstack.push(PsPack(action,ps_file_name))

#    def store_ps_file(self,action,file_content):
#        self._PSbook_autosave_flag = True
#        self._PSstack.push(FilePack(action,file_content))
#        self.open_ps_by_bytesio(self._PSstack.currentFile.fh)
#        self.recover_ps_sheet_selected()
#    def store_cas_file(self,action):
#        cas_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,8))
#        #print 'name: %s'%cas_file_name
#        bytesio = BytesIO()
#        self._CASbook.workbook_wr.save(cas_file_name)
#        with open(cas_file_name+r'.xlsx','rb') as cas_file:
#            bytesio.write(cas_file.read())
#        #self._CASbook_autosave_flag = True
#        self._CASstack.push(FilePack(action,bytesio.getvalue()))
#        bytesio.close()
    def store_cas_file(self,action):
        cas_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))
        #print 'name: %s'%cas_file_name
        self._CASbook.save_as(cas_file_name)
        #self._CASbook.workbook.save(cas_file_name)
        self._CASstack.push(CasPack(action,cas_file_name))
        self._CASbook_autosave_flag = True
        self.open_cas_by_bytesio(cas_file_name+r'.xlsx')
        self.recover_cas_sheet_selected()
    def store_cas_file_without_open(self,action):
        cas_file_name = 'tmp\\'+''.join(random.sample(string.ascii_letters,16))
        #print 'name: %s'%cas_file_name
        self._CASbook.save_as(cas_file_name)
        #self._CASbook.workbook.save(cas_file_name)
        self._CASstack.push(CasPack(action,cas_file_name))
    @pyqtSlot()
    def undo_ps(self):
        self.animation_progressBar(0)
        f = self._PSstack.pop()
        if f != None:
            self._PSbook_autosave_flag = True
            self.open_ps_by_bytesio(f[1].file_name+r'.xlsx')
            self.animation_progressBar(50)
            self.recover_ps_sheet_selected()
            self.animation_progressBar(100)
            self.refresh_message('revert action:%s'%f[0])
            self.refresh_msg('ps file revert action:%s'%f[0])
        else:
            self.refresh_message('Already at oldest change')
            self.animation_progressBar(100)
            self._PSbook_modified = False
            
    @pyqtSlot()
    def undo_cas(self):
        self.animation_progressBar(0)
        f = self._CASstack.pop()
        if f != None:
            self._CASbook_autosave_flag = True
            self.open_cas_by_bytesio(f[1].file_name+r'.xlsx')
            self.animation_progressBar(50)
            self.recover_cas_sheet_selected()
            self.animation_progressBar(100)
            self.refresh_message('revert action:%s'%f[0])
            self.refresh_msg('cas file revert action:%s'%f[0])
        else:
            self.refresh_message('Already at oldest change')
            self.animation_progressBar(100)
            self._CASbook_modified = False

    @pyqtSlot()
    def select_extended_preview(self):
        self._extended_preview = ExtendedPreview.ExtendedPreview(self._PSbook_current_sheet.extended_preview_model())
        self._extended_preview.show()


#    def refresh_cas_book_name(self,model):
#        self._window.update_cas_file(model)
#    def refresh_ps_book_name(self,model):
#        self._window.update_ps_file(model)
#    def refresh_cas_sheet_name(self,model):
#        self._window.update_cas_sheets(model)
#    def refresh_ps_sheet_name(self,model):
#        self._window.update_ps_sheets(model)
#    def refresh_preview(self,model):
#        self._window.update_preview(model)
#    def refresh_ps_header(self,model):
#        self._window.update_ps_header(model)
#    def refresh_cas_header(self,model):
#        self._window.update_cas_header(model)
#    def refresh_comparison_delete_list(self,model):
#        self._window.update_comparison_delete_list(model)
#    def refresh_comparison_append_list(self,model):
#        self._window.update_comparison_append_list(model)
#    def refresh_message(self,model):
#        self._window.update_message(model)
#    def refresh_msg(self,model):
#        self._window.update_msg(model)
#        logging.info(str(model))
#    def refresh_selected_cell(self,model):
#        self._window.update_selected_cell(model)
#    def refresh_progressBar(self,model):
#        self._window.update_progressBar(model)
#    def animation_progressBar(self,model):
#        if self._progressBar_status < model:
#            while self._progressBar_status < model:
#                self._progressBar_status += 0.002
#                self.refresh_progressBar(self._progressBar_status)
#        else:
#            self._progressBar_status = 0
#            self.refresh_progressBar(self._progressBar_status)
#            while self._progressBar_status < model:
#                self._progressBar_status += 0.002
#                self.refresh_progressBar(self._progressBar_status)
    def refresh_cas_book_name(self,model):
        print 'controller send cas book name'
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
    def refresh_message(self,model):
        self._queue_wr.put(('refresh_message',model))
    def refresh_msg(self,model):
        self._queue_wr.put(('refresh_msg',model))
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
#
#def newThread(func):
#    def wrapper(self,*args,**kw):
#        pass
class OperateThread(QThread):
    def __init__(self,func = None,parent = None):
        super(OperateThread,self).__init__(parent)
        self._func = func
    def run(self):
        if self._func != None:
            self._func()
            self.finished.emit()
class ProgressBarThread(QThread):
    def __init(self,func = None,parent = None):
        super(ProgressBarThread,self).__init(parent)
        self._refresh_func = func
        self._progressBar_status = 0
    def refresh_progressBar(self,model):
        if self._refresh_func != None:
            if self._progressBar_status < model:
                while self._progressBar_status < model:
                    self._progressBar_status += 0.002
                    self._refresh_func(self._progressBar_status)
            else:
                self._progressBar_status = 0
                self._refresh_func(self._progressBar_status)
                while self._progressBar_status < model:
                    self._progressBar_status += 0.002
                    self._refresh_func(self._progressBar_status)
       
        
            
            




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

