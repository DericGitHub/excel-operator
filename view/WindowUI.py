# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created: Mon Aug 28 17:36:31 2017
#      by: PyQt4 UI code generator 4.6.2
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1600, 850)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtGui.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtGui.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 20, 31, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.open_cas = QtGui.QPushButton(self.centralwidget)
        self.open_cas.setGeometry(QtCore.QRect(50, 60, 51, 27))
        self.open_cas.setObjectName("open_cas")
        self.sheets_cas = QtGui.QComboBox(self.centralwidget)
        self.sheets_cas.setGeometry(QtCore.QRect(460, 20, 241, 31))
        self.sheets_cas.setObjectName("sheets_cas")
        self.save_cas = QtGui.QPushButton(self.centralwidget)
        self.save_cas.setGeometry(QtCore.QRect(110, 60, 51, 27))
        self.save_cas.setObjectName("save_cas")
        self.saveas_cas = QtGui.QPushButton(self.centralwidget)
        self.saveas_cas.setGeometry(QtCore.QRect(170, 60, 51, 27))
        self.saveas_cas.setObjectName("saveas_cas")
        self.undo_cas = QtGui.QPushButton(self.centralwidget)
        self.undo_cas.setGeometry(QtCore.QRect(230, 60, 51, 27))
        self.undo_cas.setObjectName("undo_cas")
        self.undo_ps = QtGui.QPushButton(self.centralwidget)
        self.undo_ps.setGeometry(QtCore.QRect(230, 150, 51, 27))
        self.undo_ps.setObjectName("undo_ps")
        self.saveas_ps = QtGui.QPushButton(self.centralwidget)
        self.saveas_ps.setGeometry(QtCore.QRect(170, 150, 51, 27))
        self.saveas_ps.setObjectName("saveas_ps")
        self.sheets_ps = QtGui.QComboBox(self.centralwidget)
        self.sheets_ps.setGeometry(QtCore.QRect(460, 100, 241, 31))
        self.sheets_ps.setObjectName("sheets_ps")
        self.label_2 = QtGui.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(10, 110, 31, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.save_ps = QtGui.QPushButton(self.centralwidget)
        self.save_ps.setGeometry(QtCore.QRect(110, 150, 51, 27))
        self.save_ps.setObjectName("save_ps")
        self.open_ps = QtGui.QPushButton(self.centralwidget)
        self.open_ps.setGeometry(QtCore.QRect(50, 150, 51, 27))
        self.open_ps.setObjectName("open_ps")
        self.label_3 = QtGui.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(40, 500, 51, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtGui.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(40, 520, 71, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.sync_select_all_ps = QtGui.QCheckBox(self.centralwidget)
        self.sync_select_all_ps.setGeometry(QtCore.QRect(80, 770, 92, 22))
        self.sync_select_all_ps.setObjectName("sync_select_all_ps")
        self.name_cas = QtGui.QLineEdit(self.centralwidget)
        self.name_cas.setGeometry(QtCore.QRect(50, 20, 361, 25))
        self.name_cas.setReadOnly(True)
        self.name_cas.setObjectName("name_cas")
        self.name_ps = QtGui.QLineEdit(self.centralwidget)
        self.name_ps.setGeometry(QtCore.QRect(50, 110, 361, 25))
        self.name_ps.setReadOnly(True)
        self.name_ps.setObjectName("name_ps")
        self.preview = QtGui.QTableView(self.centralwidget)
        self.preview.setGeometry(QtCore.QRect(830, 0, 700, 691))
        self.preview.setEditTriggers(QtGui.QAbstractItemView.DoubleClicked)
        self.preview.setWordWrap(True)
        self.preview.setObjectName("preview")
        self.ps_header = QtGui.QListView(self.centralwidget)
        self.ps_header.setGeometry(QtCore.QRect(30, 551, 281, 211))
        self.ps_header.setObjectName("ps_header")
        self.message = QtGui.QTextBrowser(self.centralwidget)
        self.message.setGeometry(QtCore.QRect(630, 240, 181, 91))
        font = QtGui.QFont()
        font.setFamily("Helvetica [Adobe]")
        font.setPointSize(14)
        self.message.setFont(font)
        self.message.setObjectName("message")
        self.sync_ps_to_cas = QtGui.QPushButton(self.centralwidget)
        self.sync_ps_to_cas.setGeometry(QtCore.QRect(210, 770, 75, 23))
        self.sync_ps_to_cas.setObjectName("sync_ps_to_cas")
        self.sync_cas_to_ps = QtGui.QPushButton(self.centralwidget)
        self.sync_cas_to_ps.setGeometry(QtCore.QRect(500, 770, 75, 23))
        self.sync_cas_to_ps.setObjectName("sync_cas_to_ps")
        self.sync_select_all_cas = QtGui.QCheckBox(self.centralwidget)
        self.sync_select_all_cas.setGeometry(QtCore.QRect(370, 770, 92, 22))
        self.sync_select_all_cas.setObjectName("sync_select_all_cas")
        self.cas_header = QtGui.QListView(self.centralwidget)
        self.cas_header.setGeometry(QtCore.QRect(325, 551, 281, 211))
        self.cas_header.setObjectName("cas_header")
        self.label_6 = QtGui.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(330, 520, 81, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.label_5 = QtGui.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(40, 190, 121, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.label_7 = QtGui.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(40, 210, 71, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.comparison_delete_list = QtGui.QListView(self.centralwidget)
        self.comparison_delete_list.setGeometry(QtCore.QRect(30, 241, 281, 211))
        self.comparison_delete_list.setObjectName("comparison_delete_list")
        self.comparison_append_list = QtGui.QListView(self.centralwidget)
        self.comparison_append_list.setGeometry(QtCore.QRect(325, 241, 281, 211))
        self.comparison_append_list.setObjectName("comparison_append_list")
        self.comparison_select_all_delete = QtGui.QCheckBox(self.centralwidget)
        self.comparison_select_all_delete.setGeometry(QtCore.QRect(80, 460, 92, 22))
        self.comparison_select_all_delete.setObjectName("comparison_select_all_delete")
        self.comparison_append = QtGui.QPushButton(self.centralwidget)
        self.comparison_append.setGeometry(QtCore.QRect(500, 460, 75, 23))
        self.comparison_append.setObjectName("comparison_append")
        self.comparison_select_all_append = QtGui.QCheckBox(self.centralwidget)
        self.comparison_select_all_append.setGeometry(QtCore.QRect(370, 460, 92, 22))
        self.comparison_select_all_append.setObjectName("comparison_select_all_append")
        self.label_8 = QtGui.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(330, 210, 81, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.comparison_delete = QtGui.QPushButton(self.centralwidget)
        self.comparison_delete.setGeometry(QtCore.QRect(210, 460, 75, 23))
        self.comparison_delete.setObjectName("comparison_delete")
        self.comparison_start = QtGui.QPushButton(self.centralwidget)
        self.comparison_start.setGeometry(QtCore.QRect(170, 190, 90, 27))
        self.comparison_start.setObjectName("comparison_start")
        self.preview_add = QtGui.QPushButton(self.centralwidget)
        self.preview_add.setGeometry(QtCore.QRect(930, 710, 90, 27))
        self.preview_add.setObjectName("preview_add")
        self.preview_delete = QtGui.QPushButton(self.centralwidget)
        self.preview_delete.setGeometry(QtCore.QRect(1070, 710, 90, 27))
        self.preview_delete.setObjectName("preview_delete")
        self.preview_lock = QtGui.QPushButton(self.centralwidget)
        self.preview_lock.setGeometry(QtCore.QRect(1230, 710, 90, 27))
        self.preview_lock.setObjectName("preview_lock")
        self.extended_preview = QtGui.QPushButton(self.centralwidget)
        self.extended_preview.setGeometry(QtCore.QRect(1360, 710, 90, 27))
        self.extended_preview.setObjectName("extended_preview")
        self.label_9 = QtGui.QLabel(self.centralwidget)
        self.label_9.setGeometry(QtCore.QRect(1000, 760, 81, 31))
        font = QtGui.QFont()
        font.setFamily("Abyssinica SIL")
        font.setPointSize(18)
        self.label_9.setFont(font)
        self.label_9.setObjectName("label_9")
        self.label_10 = QtGui.QLabel(self.centralwidget)
        self.label_10.setGeometry(QtCore.QRect(1120, 760, 81, 31))
        font = QtGui.QFont()
        font.setFamily("Abyssinica SIL")
        font.setPointSize(18)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.selected_row = QtGui.QLabel(self.centralwidget)
        self.selected_row.setGeometry(QtCore.QRect(1060, 760, 81, 31))
        font = QtGui.QFont()
        font.setFamily("Abyssinica SIL")
        font.setPointSize(18)
        self.selected_row.setFont(font)
        self.selected_row.setObjectName("selected_row")
        self.selected_col = QtGui.QLabel(self.centralwidget)
        self.selected_col.setGeometry(QtCore.QRect(1180, 760, 81, 31))
        font = QtGui.QFont()
        font.setFamily("Abyssinica SIL")
        font.setPointSize(18)
        self.selected_col.setFont(font)
        self.selected_col.setObjectName("selected_col")
        self.msg = QtGui.QLabel(self.centralwidget)
        self.msg.setGeometry(QtCore.QRect(1250, 780, 321, 17))
        self.msg.setObjectName("msg")
        self.progressBar = QtGui.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(1250, 750, 321, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtGui.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1600, 25))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtGui.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QtGui.QApplication.translate("MainWindow", "Excel Operator", None, QtGui.QApplication.UnicodeUTF8))
        self.label.setText(QtGui.QApplication.translate("MainWindow", "CAS", None, QtGui.QApplication.UnicodeUTF8))
        self.open_cas.setText(QtGui.QApplication.translate("MainWindow", "Open", None, QtGui.QApplication.UnicodeUTF8))
        self.save_cas.setText(QtGui.QApplication.translate("MainWindow", "Save", None, QtGui.QApplication.UnicodeUTF8))
        self.saveas_cas.setText(QtGui.QApplication.translate("MainWindow", "Saveas", None, QtGui.QApplication.UnicodeUTF8))
        self.undo_cas.setText(QtGui.QApplication.translate("MainWindow", "Undo", None, QtGui.QApplication.UnicodeUTF8))
        self.undo_ps.setText(QtGui.QApplication.translate("MainWindow", "Undo", None, QtGui.QApplication.UnicodeUTF8))
        self.saveas_ps.setText(QtGui.QApplication.translate("MainWindow", "Saveas", None, QtGui.QApplication.UnicodeUTF8))
        self.label_2.setText(QtGui.QApplication.translate("MainWindow", "PS", None, QtGui.QApplication.UnicodeUTF8))
        self.save_ps.setText(QtGui.QApplication.translate("MainWindow", "Save", None, QtGui.QApplication.UnicodeUTF8))
        self.open_ps.setText(QtGui.QApplication.translate("MainWindow", "Open", None, QtGui.QApplication.UnicodeUTF8))
        self.label_3.setText(QtGui.QApplication.translate("MainWindow", "Sync", None, QtGui.QApplication.UnicodeUTF8))
        self.label_4.setText(QtGui.QApplication.translate("MainWindow", "PS file", None, QtGui.QApplication.UnicodeUTF8))
        self.sync_select_all_ps.setText(QtGui.QApplication.translate("MainWindow", "select all", None, QtGui.QApplication.UnicodeUTF8))
        self.sync_ps_to_cas.setText(QtGui.QApplication.translate("MainWindow", "PS->CAS", None, QtGui.QApplication.UnicodeUTF8))
        self.sync_cas_to_ps.setText(QtGui.QApplication.translate("MainWindow", "CAS->PS", None, QtGui.QApplication.UnicodeUTF8))
        self.sync_select_all_cas.setText(QtGui.QApplication.translate("MainWindow", "select all", None, QtGui.QApplication.UnicodeUTF8))
        self.label_6.setText(QtGui.QApplication.translate("MainWindow", "CAS file", None, QtGui.QApplication.UnicodeUTF8))
        self.label_5.setText(QtGui.QApplication.translate("MainWindow", "Comparison", None, QtGui.QApplication.UnicodeUTF8))
        self.label_7.setText(QtGui.QApplication.translate("MainWindow", "PS file", None, QtGui.QApplication.UnicodeUTF8))
        self.comparison_select_all_delete.setText(QtGui.QApplication.translate("MainWindow", "select all", None, QtGui.QApplication.UnicodeUTF8))
        self.comparison_append.setText(QtGui.QApplication.translate("MainWindow", "append", None, QtGui.QApplication.UnicodeUTF8))
        self.comparison_select_all_append.setText(QtGui.QApplication.translate("MainWindow", "select all", None, QtGui.QApplication.UnicodeUTF8))
        self.label_8.setText(QtGui.QApplication.translate("MainWindow", "CAS file", None, QtGui.QApplication.UnicodeUTF8))
        self.comparison_delete.setText(QtGui.QApplication.translate("MainWindow", "delete", None, QtGui.QApplication.UnicodeUTF8))
        self.comparison_start.setText(QtGui.QApplication.translate("MainWindow", "Comparison", None, QtGui.QApplication.UnicodeUTF8))
        self.preview_add.setText(QtGui.QApplication.translate("MainWindow", "Add", None, QtGui.QApplication.UnicodeUTF8))
        self.preview_delete.setText(QtGui.QApplication.translate("MainWindow", "Delete", None, QtGui.QApplication.UnicodeUTF8))
        self.preview_lock.setText(QtGui.QApplication.translate("MainWindow", "Lock", None, QtGui.QApplication.UnicodeUTF8))
        self.extended_preview.setText(QtGui.QApplication.translate("MainWindow", "Preview", None, QtGui.QApplication.UnicodeUTF8))
        self.label_9.setText(QtGui.QApplication.translate("MainWindow", "Row:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_10.setText(QtGui.QApplication.translate("MainWindow", "Col:", None, QtGui.QApplication.UnicodeUTF8))
        self.msg.setText(QtGui.QApplication.translate("MainWindow", "TextLabel", None, QtGui.QApplication.UnicodeUTF8))

import resource_rc
