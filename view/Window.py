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
        #self.init_model(CASbook,PSbook)
    def init_Window(self):
        self.ui = WindowUI.Ui_MainWindow()
        self.ui.setupUi(self)
        self.timer = QBasicTimer()
    def init_model(self,CASbook,PSbook):
        self.ui.name_cas.setText(CASbook)
        self.ui.name_ps.setText(PSbook)

        #       sheets_cas_model
        #self._sheets_cas_model = QStandardItemModel()
        #self.ui.sheets_cas.setModel(self._sheets_cas_model)

        ##       sheets_ps_model
        #self._sheets_ps_model = QStandardItemModel()
        #self.ui.sheets_ps.setModel(self._sheets_ps_model)

        ##       preview_model
        #self._preview_model = QStandardItemModel()
        #self.ui.preview.setModel(self._preview_model)
    
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
#    def update_cas_sheets(self,sheets_name):
#        self._sheets_cas_model.clear()
#        for sheet_name in sheets_name:
#            item = QStandardItem(sheet_name)
#            self._sheets_cas_model.appendRow(item)
    @pyqtSlot(list)
    def update_cas_sheets(self,sheetnames):
        model = QStandardItemModel()
        for sheetname in sheetnames:
            item = QStandardItem(sheetname)
            model.appendRow(item)
        self.ui.sheets_cas.setModel(model)
#    def update_ps_sheets(self,sheets_name):
#        self._sheets_ps_model.clear()
#        for sheet_name in sheets_name:
#            item = QStandardItem(sheet_name)
#            self._sheets_ps_model.appendRow(item)
    @pyqtSlot(list)
    def update_ps_sheets(self,sheetnames):
        model = QStandardItemModel()
        for sheetname in sheetnames:
            item = QStandardItem(sheetname)
            model.appendRow(item)
        self.ui.sheets_ps.setModel(model)
#    def update_preview(self):
#        for i in range(3):
#            items =[]
#            for j in range(50):
#                item = QStandardItem(str(j))
#                if i == 0:
#                    item.setCheckState(False)
#                    item.setCheckable(True)
#                items.append(item)
#            
#            self._preview_model.appendColumn(items)
#        print self._preview_model
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
        self.ui.preview.setColumnWidth(0,100)
        self.ui.preview.setColumnWidth(1,150)
        self.ui.preview.setColumnWidth(2,200)
        self.ui.preview.setColumnWidth(3,200)
        self.ui.preview.resizeRowsToContents()
    @pyqtSlot(list)
    def update_ps_header(self,headers):
        model = QStandardItemModel()
        if len(headers) != 0:
            for header in headers:
                item = QStandardItem(header)
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
                item = QStandardItem(header)
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
            item = QStandardItem(delete)
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
            item = QStandardItem(append)
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
    @pyqtSlot(int)
    def update_progressBar(self,model):
        self.ui.progressBar.setValue(int(model))
    def open_file_confirm(self):
        #choice = QMessageBox.question(None,'Attention','File has been modified, sure to leave?',QMessageBox.Save|'Save as'|QMessageBox.Discard)
        choice = QMessageBox()
        choice.setWindowTitle('attention')
        choice.setText('File has been modified, sure to leave?')
        choice.addButton(QString('Save'),QMessageBox.AcceptRole)
        choice.addButton(QString('Save as'),QMessageBox.AcceptRole)
        choice.addButton(QString('Don\'t save'),QMessageBox.RejectRole)
        if choice == QMessageBox.Yes:
            return 0 
        elif choice == QMessageBox.Help:
            return 1
        elif choice == QMessageBox.Discard:
            return 2
#class SaveConfirm(QMessageBox):
#    def __init__(self,parent=None):
#        super(SaveConfirm,self).__init__(parent)
        

def open_file_dialog():
    filedialog = QFileDialog()
    #filedialog.setNameFilter('*.xlsx')
    fileName = filedialog.getOpenFileName(filter = '*.xlsx')
    return fileName
def save_file_dialog():
    filedialog = QFileDialog()
    fileName = filedialog.getSaveFileName(filter = '*.xlsx')
    return fileName
def open_file_confirm(self):
    choice = QMessageBox.question(None,'Attention','File has been modified, sure to leave?',QMessageBox.Save|'Save as'|QMessageBox.Discard)
    if choice == QMessageBox.Yes:
        return True
    else:
        return False
