#coding:utf-8
from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
import WindowUI
import sys
from model import *
##################################################
#       class to handle the view
##################################################
class Window(QMainWindow):

    ##################################################
    #       Initial method
    ##################################################
    def __init__(self):
        super(Window,self).__init__()
        self.init_Window()
    def init_Window(self):
        self.ui = WindowUI.Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowState(Qt.WindowMaximized)

    ##################################################
    #       Bind event
    ##################################################
    def bind_open_cas(self,func):
        self.ui.open_cas.clicked.connect(func)
    def bind_open_ps(self,func):
        self.ui.open_ps.clicked.connect(func)
    def bind_save_cas(self,func):
        self.ui.save_cas.clicked.connect(func)
    def bind_save_ps(self,func):
        self.ui.save_ps.clicked.connect(func)
    def bind_saveas_cas(self,func):
        self.ui.saveas_cas.clicked.connect(func)
    def bind_saveas_ps(self,func):
        self.ui.saveas_ps.clicked.connect(func)
    def bind_select_cas_sheet(self,func):
        self.ui.sheets_cas.currentIndexChanged.connect(func)
    def bind_select_ps_sheet(self,func):
        self.ui.sheets_ps.currentIndexChanged.connect(func)
    def bind_select_preview(self,func):
        self.ui.preview.clicked.connect(func)
    def bind_sync_ps_to_cas(self,func):
        self.ui.sync_ps_to_cas.clicked.connect(func)
    def bind_sync_cas_to_ps(self,func):
        self.ui.sync_cas_to_ps.clicked.connect(func)
    def bind_sync_select_all_ps_headers(self,func):
        self.ui.sync_select_all_ps.stateChanged.connect(func)
    def bind_sync_select_all_cas_headers(self,func):
        self.ui.sync_select_all_cas.stateChanged.connect(func)
    def bind_comparison_start(self,func):
        self.ui.comparison_start.clicked.connect(func)
    def bind_comparison_delete(self,func):
        self.ui.comparison_delete.clicked.connect(func)
    def bind_comparison_append(self,func):
        self.ui.comparison_append.clicked.connect(func)
    def bind_comparison_select_all_delete(self,func):
        self.ui.comparison_select_all_delete.stateChanged.connect(func)
    def bind_comparison_select_all_append(self,func):
        self.ui.comparison_select_all_append.stateChanged.connect(func)
    def bind_preview_add(self,func):
        self.ui.preview_add.clicked.connect(func)
    def bind_preview_delete(self,func):
        self.ui.preview_delete.clicked.connect(func)
    def bind_preview_lock(self,func):
        self.ui.preview_lock.clicked.connect(func)
    def bind_undo_cas(self,func):
        self.ui.undo_cas.clicked.connect(func)
    def bind_undo_ps(self,func):
        self.ui.undo_ps.clicked.connect(func)
    def bind_select_extended_preview(self,func):
        self.ui.extended_preview.clicked.connect(func)
    def bind_ps_header_changed(self,func):
        self.ui.ps_header.clicked.connect(func)
    def bind_cas_header_changed(self,func):
        self.ui.cas_header.clicked.connect(func)
    def bind_comparison_append_list_changed(self,func):
        self.ui.comparison_append_list.clicked.connect(func)
    def bind_comparison_delete_list_changed(self,func):
        self.ui.comparison_delete_list.clicked.connect(func)
    ##################################################
    #       Custom slot
    ##################################################
    @pyqtSlot(str)
    def update_cas_file(self,filename):
        self.ui.name_cas.setText(filename)
    @pyqtSlot(str)
    def update_ps_file(self,filename):
        self.ui.name_ps.setText(filename)
    @pyqtSlot(list)
    def update_cas_sheets(self,sheetnames):
        model = QStandardItemModel()
        for sheetname in sheetnames:
            item = QStandardItem(sheetname)
            model.appendRow(item)
        self.ui.sheets_cas.setModel(model)
    @pyqtSlot(list)
    def update_ps_sheets(self,sheetnames):
        model = QStandardItemModel()
        for sheetname in sheetnames:
            item = QStandardItem(sheetname)
            model.appendRow(item)
        self.ui.sheets_ps.setModel(model)
    @pyqtSlot(list)
    def update_preview(self,itemss):
        model = QStandardItemModel()
        model.setColumnCount(4)
        model.setHeaderData(0,Qt.Horizontal,'status')
        model.setHeaderData(1,Qt.Horizontal,'subject matter')
        model.setHeaderData(2,Qt.Horizontal,'container name')
        model.setHeaderData(3,Qt.Horizontal,'xmlname')
        for items in itemss:
            item1 = QStandardItem(items[0] if items[0] is not None else '')
            item2 = QStandardItem(items[1] if items[1] is not None else '')
            item3 = QStandardItem(items[2] if items[2] is not None else '')
            item4 = QStandardItem(items[3] if items[3] is not None else '')
            model.appendRow((item1,item2,item3,item4))

        self.ui.preview.setModel(model)
        size = (self.ui.preview.width()-50)/4
        self.ui.preview.setColumnWidth(0,size)
        self.ui.preview.setColumnWidth(1,size)
        self.ui.preview.setColumnWidth(2,size)
        self.ui.preview.setColumnWidth(3,size)
#        self.ui.preview.setColumnWidth(0,100)
#        self.ui.preview.setColumnWidth(1,150)
#        self.ui.preview.setColumnWidth(2,200)
#        self.ui.preview.setColumnWidth(3,200)
        self.ui.preview.resizeRowsToContents()
    @pyqtSlot(list)
    def update_ps_header(self,headers):
        model = QStandardItemModel()
        if len(headers) != 0:
            for header in headers:
                item = QStandardItem(QString(header) if header is not None else '')
                item.setCheckState(Qt.Unchecked)
                item.setCheckable(True)
                model.appendRow(item)
        # change color for every two rows
        cnt = model.rowCount()
        for i in range(cnt):
            if i%2 == 0:
                model.item(i).setBackground(QBrush(QColor(217,217,217)))
        self.ui.ps_header.setModel(model)
    @pyqtSlot(list)
    def update_cas_header(self,headers):
        model = QStandardItemModel()
        if len(headers) != 0:
            for header in headers:
                if header == None:
                    header = ''
                item = QStandardItem(QString(header) if header is not None else '')
                item.setCheckState(Qt.Unchecked)
                item.setCheckable(True)
                model.appendRow(item)
        # change color for every two rows
        cnt = model.rowCount()
        for i in range(cnt):
            if i%2 == 0:
                model.item(i).setBackground(QBrush(QColor(217,217,217)))
        self.ui.cas_header.setModel(model)
    @pyqtSlot(int)
    def update_ps_header_selected(self,idx):
        self.ui.sheets_ps.setCurrentIndex(idx)
    @pyqtSlot(int)
    def update_cas_header_selected(self,idx):
        self.ui.sheets_cas.setCurrentIndex(idx)
    @pyqtSlot(list)
    def update_comparison_delete_list(self,deletes):
        model = QStandardItemModel()
        for delete in deletes:
            item = QStandardItem(delete if delete is not None else '')
            item.setCheckState(Qt.Unchecked)
            item.setCheckable(True)
            model.appendRow(item)
        # change color for every two rows
        cnt = model.rowCount()
        for i in range(cnt):
            if i%2 == 0:
                model.item(i).setBackground(QBrush(QColor(217,217,217)))
        self.ui.comparison_delete_list.setModel(model)
    @pyqtSlot(list)
    def update_comparison_append_list(self,appends):
        model = QStandardItemModel()
        for append in appends:
            item = QStandardItem(append if append is not None else '')
            item.setCheckState(Qt.Unchecked)
            item.setCheckable(True)
            model.appendRow(item)
        # change color for every two rows
        cnt = model.rowCount()
        for i in range(cnt):
            if i%2 == 0:
                model.item(i).setBackground(QBrush(QColor(217,217,217)))
        self.ui.comparison_append_list.setModel(model)
    @pyqtSlot(str)
    def update_message(self,model):
        self.ui.message.setText(str(model))
    @pyqtSlot(str)
    def update_msg(self,model):
        self.ui.msg.setText(str(model))
    @pyqtSlot(list)
    def update_selected_cell(self,model):
        self.ui.selected_row.setText(str(model[0]))
        self.ui.selected_col.setText(str(model[1]))
    @pyqtSlot(float)
    def update_progressBar(self,model):
        self.ui.progressBar.setValue(float(model))
    def open_file_confirm(self):
        choice = QMessageBox()
        choice.setWindowTitle('attention')
        choice.setText('The original file has been modified')
        choice.addButton(QString('Save'),QMessageBox.AcceptRole)
        choice.addButton(QString('Save as'),QMessageBox.AcceptRole)
        choice.addButton(QString('Don\'t save'),QMessageBox.RejectRole)
        ret = choice.exec_()
        return ret
    def open_action_confirm(self,action):
        choice = QMessageBox()
        choice.setWindowTitle('attention')
        choice.setText('%s'%action) 
        choice.addButton(QString('Ok'),QMessageBox.AcceptRole)
        choice.addButton(QString('Cancel'),QMessageBox.RejectRole)
        ret = choice.exec_()
        return ret
    @pyqtSlot(str)
    def pop_up_message(self,msg):
        ret = QMessageBox()#QMessageBox.Warning,'Warning',msg)
        ret.setWindowIcon(self.windowIcon())
        ret.setIcon(QMessageBox.Warning)
        ret.setWindowTitle('Attention')
        ret.setText(msg)
        ret.exec_()
        

def open_file_dialog():
    filedialog = QFileDialog()
    fileName = filedialog.getOpenFileName(filter = '*.xlsx')
    return fileName
def save_file_dialog():
    filedialog = QFileDialog()
    fileName = filedialog.getSaveFileName(filter = '*.xlsx')
    return fileName

