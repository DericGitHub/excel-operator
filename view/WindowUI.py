# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Window.ui'
#
# Created: Wed Oct 11 16:00:00 2017
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
        self.gridLayout_7 = QtGui.QGridLayout(self.centralwidget)
        self.gridLayout_7.setObjectName("gridLayout_7")
        self.horizontalLayout = QtGui.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.verticalLayout = QtGui.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.gridLayout = QtGui.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.label = QtGui.QLabel(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.name_cas = QtGui.QLineEdit(self.centralwidget)
        self.name_cas.setReadOnly(True)
        self.name_cas.setObjectName("name_cas")
        self.gridLayout.addWidget(self.name_cas, 0, 1, 1, 4)
        self.sheets_cas = QtGui.QComboBox(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Minimum, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sheets_cas.sizePolicy().hasHeightForWidth())
        self.sheets_cas.setSizePolicy(sizePolicy)
        self.sheets_cas.setMinimumSize(QtCore.QSize(150, 0))
        self.sheets_cas.setObjectName("sheets_cas")
        self.gridLayout.addWidget(self.sheets_cas, 0, 5, 1, 1)
        self.open_cas = QtGui.QPushButton(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.open_cas.sizePolicy().hasHeightForWidth())
        self.open_cas.setSizePolicy(sizePolicy)
        self.open_cas.setObjectName("open_cas")
        self.gridLayout.addWidget(self.open_cas, 1, 1, 1, 1)
        self.save_cas = QtGui.QPushButton(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.save_cas.sizePolicy().hasHeightForWidth())
        self.save_cas.setSizePolicy(sizePolicy)
        self.save_cas.setObjectName("save_cas")
        self.gridLayout.addWidget(self.save_cas, 1, 2, 1, 1)
        self.saveas_cas = QtGui.QPushButton(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.saveas_cas.sizePolicy().hasHeightForWidth())
        self.saveas_cas.setSizePolicy(sizePolicy)
        self.saveas_cas.setObjectName("saveas_cas")
        self.gridLayout.addWidget(self.saveas_cas, 1, 3, 1, 1)
        self.undo_cas = QtGui.QPushButton(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.undo_cas.sizePolicy().hasHeightForWidth())
        self.undo_cas.setSizePolicy(sizePolicy)
        self.undo_cas.setObjectName("undo_cas")
        self.gridLayout.addWidget(self.undo_cas, 1, 4, 1, 1)
        self.label_2 = QtGui.QLabel(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_2.sizePolicy().hasHeightForWidth())
        self.label_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 2, 0, 1, 1)
        self.name_ps = QtGui.QLineEdit(self.centralwidget)
        self.name_ps.setReadOnly(True)
        self.name_ps.setObjectName("name_ps")
        self.gridLayout.addWidget(self.name_ps, 2, 1, 1, 4)
        self.sheets_ps = QtGui.QComboBox(self.centralwidget)
        self.sheets_ps.setMinimumSize(QtCore.QSize(150, 0))
        self.sheets_ps.setObjectName("sheets_ps")
        self.gridLayout.addWidget(self.sheets_ps, 2, 5, 1, 1)
        self.open_ps = QtGui.QPushButton(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.open_ps.sizePolicy().hasHeightForWidth())
        self.open_ps.setSizePolicy(sizePolicy)
        self.open_ps.setObjectName("open_ps")
        self.gridLayout.addWidget(self.open_ps, 3, 1, 1, 1)
        self.save_ps = QtGui.QPushButton(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.save_ps.sizePolicy().hasHeightForWidth())
        self.save_ps.setSizePolicy(sizePolicy)
        self.save_ps.setObjectName("save_ps")
        self.gridLayout.addWidget(self.save_ps, 3, 2, 1, 1)
        self.saveas_ps = QtGui.QPushButton(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.saveas_ps.sizePolicy().hasHeightForWidth())
        self.saveas_ps.setSizePolicy(sizePolicy)
        self.saveas_ps.setObjectName("saveas_ps")
        self.gridLayout.addWidget(self.saveas_ps, 3, 3, 1, 1)
        self.undo_ps = QtGui.QPushButton(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.undo_ps.sizePolicy().hasHeightForWidth())
        self.undo_ps.setSizePolicy(sizePolicy)
        self.undo_ps.setObjectName("undo_ps")
        self.gridLayout.addWidget(self.undo_ps, 3, 4, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout)
        self.gridLayout_3 = QtGui.QGridLayout()
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.label_5 = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.gridLayout_3.addWidget(self.label_5, 0, 0, 1, 2)
        self.label_7 = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.gridLayout_3.addWidget(self.label_7, 1, 0, 1, 1)
        self.label_8 = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.gridLayout_3.addWidget(self.label_8, 1, 3, 1, 1)
        self.comparison_delete_list = QtGui.QListView(self.centralwidget)
        self.comparison_delete_list.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.comparison_delete_list.setVerticalScrollMode(QtGui.QAbstractItemView.ScrollPerPixel)
        self.comparison_delete_list.setObjectName("comparison_delete_list")
        self.gridLayout_3.addWidget(self.comparison_delete_list, 2, 0, 1, 3)
        self.comparison_append_list = QtGui.QListView(self.centralwidget)
        self.comparison_append_list.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.comparison_append_list.setVerticalScrollMode(QtGui.QAbstractItemView.ScrollPerPixel)
        self.comparison_append_list.setObjectName("comparison_append_list")
        self.gridLayout_3.addWidget(self.comparison_append_list, 2, 3, 1, 2)
        self.comparison_delete = QtGui.QPushButton(self.centralwidget)
        self.comparison_delete.setObjectName("comparison_delete")
        self.gridLayout_3.addWidget(self.comparison_delete, 3, 2, 1, 1)
        self.comparison_select_all_append = QtGui.QCheckBox(self.centralwidget)
        self.comparison_select_all_append.setObjectName("comparison_select_all_append")
        self.gridLayout_3.addWidget(self.comparison_select_all_append, 3, 3, 1, 1)
        self.comparison_append = QtGui.QPushButton(self.centralwidget)
        self.comparison_append.setObjectName("comparison_append")
        self.gridLayout_3.addWidget(self.comparison_append, 3, 4, 1, 1)
        self.comparison_select_all_delete = QtGui.QCheckBox(self.centralwidget)
        self.comparison_select_all_delete.setObjectName("comparison_select_all_delete")
        self.gridLayout_3.addWidget(self.comparison_select_all_delete, 3, 0, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_3)
        self.gridLayout_6 = QtGui.QGridLayout()
        self.gridLayout_6.setObjectName("gridLayout_6")
        self.label_3 = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.gridLayout_6.addWidget(self.label_3, 0, 0, 1, 1)
        self.label_4 = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.gridLayout_6.addWidget(self.label_4, 1, 0, 1, 1)
        self.label_6 = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.gridLayout_6.addWidget(self.label_6, 1, 3, 1, 1)
        self.ps_header = QtGui.QListView(self.centralwidget)
        self.ps_header.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.ps_header.setVerticalScrollMode(QtGui.QAbstractItemView.ScrollPerPixel)
        self.ps_header.setObjectName("ps_header")
        self.gridLayout_6.addWidget(self.ps_header, 2, 0, 1, 3)
        self.cas_header = QtGui.QListView(self.centralwidget)
        self.cas_header.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.cas_header.setVerticalScrollMode(QtGui.QAbstractItemView.ScrollPerPixel)
        self.cas_header.setObjectName("cas_header")
        self.gridLayout_6.addWidget(self.cas_header, 2, 3, 1, 2)
        self.sync_ps_to_cas = QtGui.QPushButton(self.centralwidget)
        self.sync_ps_to_cas.setObjectName("sync_ps_to_cas")
        self.gridLayout_6.addWidget(self.sync_ps_to_cas, 3, 2, 1, 1)
        self.sync_select_all_cas = QtGui.QCheckBox(self.centralwidget)
        self.sync_select_all_cas.setObjectName("sync_select_all_cas")
        self.gridLayout_6.addWidget(self.sync_select_all_cas, 3, 3, 1, 1)
        self.sync_cas_to_ps = QtGui.QPushButton(self.centralwidget)
        self.sync_cas_to_ps.setObjectName("sync_cas_to_ps")
        self.gridLayout_6.addWidget(self.sync_cas_to_ps, 3, 4, 1, 1)
        self.sync_select_all_ps = QtGui.QCheckBox(self.centralwidget)
        self.sync_select_all_ps.setObjectName("sync_select_all_ps")
        self.gridLayout_6.addWidget(self.sync_select_all_ps, 3, 0, 1, 1)
        self.verticalLayout.addLayout(self.gridLayout_6)
        self.horizontalLayout.addLayout(self.verticalLayout)
        self.verticalLayout_2 = QtGui.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.gridLayout_4 = QtGui.QGridLayout()
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.preview = QtGui.QTableView(self.centralwidget)
        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.preview.sizePolicy().hasHeightForWidth())
        self.preview.setSizePolicy(sizePolicy)
        self.preview.setMinimumSize(QtCore.QSize(700, 0))
        self.preview.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.preview.setVerticalScrollMode(QtGui.QAbstractItemView.ScrollPerPixel)
        self.preview.setHorizontalScrollMode(QtGui.QAbstractItemView.ScrollPerPixel)
        self.preview.setWordWrap(True)
        self.preview.setObjectName("preview")
        self.gridLayout_4.addWidget(self.preview, 0, 0, 1, 6)
        spacerItem = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.gridLayout_4.addItem(spacerItem, 1, 0, 1, 1)
        self.preview_add = QtGui.QPushButton(self.centralwidget)
        self.preview_add.setObjectName("preview_add")
        self.gridLayout_4.addWidget(self.preview_add, 1, 1, 1, 1)
        self.preview_delete = QtGui.QPushButton(self.centralwidget)
        self.preview_delete.setObjectName("preview_delete")
        self.gridLayout_4.addWidget(self.preview_delete, 1, 2, 1, 1)
        self.preview_lock = QtGui.QPushButton(self.centralwidget)
        self.preview_lock.setObjectName("preview_lock")
        self.gridLayout_4.addWidget(self.preview_lock, 1, 3, 1, 1)
        self.extended_preview = QtGui.QPushButton(self.centralwidget)
        self.extended_preview.setObjectName("extended_preview")
        self.gridLayout_4.addWidget(self.extended_preview, 1, 4, 1, 1)
        spacerItem1 = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.gridLayout_4.addItem(spacerItem1, 1, 5, 1, 1)
        self.verticalLayout_2.addLayout(self.gridLayout_4)
        self.gridLayout_5 = QtGui.QGridLayout()
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.label_9 = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Abyssinica SIL")
        font.setPointSize(18)
        self.label_9.setFont(font)
        self.label_9.setObjectName("label_9")
        self.gridLayout_5.addWidget(self.label_9, 0, 0, 1, 1)
        self.selected_row = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Abyssinica SIL")
        font.setPointSize(18)
        self.selected_row.setFont(font)
        self.selected_row.setObjectName("selected_row")
        self.gridLayout_5.addWidget(self.selected_row, 0, 1, 1, 1)
        self.label_10 = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Abyssinica SIL")
        font.setPointSize(18)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.gridLayout_5.addWidget(self.label_10, 0, 2, 1, 1)
        self.selected_col = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Abyssinica SIL")
        font.setPointSize(18)
        self.selected_col.setFont(font)
        self.selected_col.setObjectName("selected_col")
        self.gridLayout_5.addWidget(self.selected_col, 0, 3, 1, 1)
        spacerItem2 = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.gridLayout_5.addItem(spacerItem2, 0, 4, 1, 1)
        self.progressBar = QtGui.QProgressBar(self.centralwidget)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.gridLayout_5.addWidget(self.progressBar, 0, 5, 1, 1)
        spacerItem3 = QtGui.QSpacerItem(40, 20, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        self.gridLayout_5.addItem(spacerItem3, 1, 4, 1, 1)
        self.msg = QtGui.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setFamily("Abyssinica SIL")
        font.setPointSize(18)
        self.msg.setFont(font)
        self.msg.setObjectName("msg")
        self.gridLayout_5.addWidget(self.msg, 1, 5, 1, 1)
        self.verticalLayout_2.addLayout(self.gridLayout_5)
        self.horizontalLayout.addLayout(self.verticalLayout_2)
        self.gridLayout_7.addLayout(self.horizontalLayout, 0, 0, 1, 1)
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
        self.label_2.setText(QtGui.QApplication.translate("MainWindow", "PS", None, QtGui.QApplication.UnicodeUTF8))
        self.open_ps.setText(QtGui.QApplication.translate("MainWindow", "Open", None, QtGui.QApplication.UnicodeUTF8))
        self.save_ps.setText(QtGui.QApplication.translate("MainWindow", "Save", None, QtGui.QApplication.UnicodeUTF8))
        self.saveas_ps.setText(QtGui.QApplication.translate("MainWindow", "Saveas", None, QtGui.QApplication.UnicodeUTF8))
        self.undo_ps.setText(QtGui.QApplication.translate("MainWindow", "Undo", None, QtGui.QApplication.UnicodeUTF8))
        self.label_5.setText(QtGui.QApplication.translate("MainWindow", "Comparison", None, QtGui.QApplication.UnicodeUTF8))
        self.label_7.setText(QtGui.QApplication.translate("MainWindow", "PS file", None, QtGui.QApplication.UnicodeUTF8))
        self.label_8.setText(QtGui.QApplication.translate("MainWindow", "CAS file", None, QtGui.QApplication.UnicodeUTF8))
        self.comparison_delete.setText(QtGui.QApplication.translate("MainWindow", "Delete", None, QtGui.QApplication.UnicodeUTF8))
        self.comparison_select_all_append.setText(QtGui.QApplication.translate("MainWindow", "select all", None, QtGui.QApplication.UnicodeUTF8))
        self.comparison_append.setText(QtGui.QApplication.translate("MainWindow", "Append", None, QtGui.QApplication.UnicodeUTF8))
        self.comparison_select_all_delete.setText(QtGui.QApplication.translate("MainWindow", "select all", None, QtGui.QApplication.UnicodeUTF8))
        self.label_3.setText(QtGui.QApplication.translate("MainWindow", "Sync", None, QtGui.QApplication.UnicodeUTF8))
        self.label_4.setText(QtGui.QApplication.translate("MainWindow", "PS file", None, QtGui.QApplication.UnicodeUTF8))
        self.label_6.setText(QtGui.QApplication.translate("MainWindow", "CAS file", None, QtGui.QApplication.UnicodeUTF8))
        self.sync_ps_to_cas.setText(QtGui.QApplication.translate("MainWindow", "PS->CAS", None, QtGui.QApplication.UnicodeUTF8))
        self.sync_select_all_cas.setText(QtGui.QApplication.translate("MainWindow", "select all", None, QtGui.QApplication.UnicodeUTF8))
        self.sync_cas_to_ps.setText(QtGui.QApplication.translate("MainWindow", "CAS->PS", None, QtGui.QApplication.UnicodeUTF8))
        self.sync_select_all_ps.setText(QtGui.QApplication.translate("MainWindow", "select all", None, QtGui.QApplication.UnicodeUTF8))
        self.preview_add.setText(QtGui.QApplication.translate("MainWindow", "Add", None, QtGui.QApplication.UnicodeUTF8))
        self.preview_delete.setText(QtGui.QApplication.translate("MainWindow", "Delete", None, QtGui.QApplication.UnicodeUTF8))
        self.preview_lock.setText(QtGui.QApplication.translate("MainWindow", "Lock POR", None, QtGui.QApplication.UnicodeUTF8))
        self.extended_preview.setText(QtGui.QApplication.translate("MainWindow", "Preview", None, QtGui.QApplication.UnicodeUTF8))
        self.label_9.setText(QtGui.QApplication.translate("MainWindow", "Row:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_10.setText(QtGui.QApplication.translate("MainWindow", "Col:", None, QtGui.QApplication.UnicodeUTF8))

import resource_rc
