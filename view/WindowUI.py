# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created: Thu Aug  3 13:23:59 2017
#      by: PyQt4 UI code generator 4.6.2
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(981, 790)
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
        self.redo_cas = QtGui.QPushButton(self.centralwidget)
        self.redo_cas.setGeometry(QtCore.QRect(290, 60, 51, 27))
        self.redo_cas.setObjectName("redo_cas")
        self.undo_ps = QtGui.QPushButton(self.centralwidget)
        self.undo_ps.setGeometry(QtCore.QRect(230, 150, 51, 27))
        self.undo_ps.setObjectName("undo_ps")
        self.saveas_ps = QtGui.QPushButton(self.centralwidget)
        self.saveas_ps.setGeometry(QtCore.QRect(170, 150, 51, 27))
        self.saveas_ps.setObjectName("saveas_ps")
        self.redo_ps = QtGui.QPushButton(self.centralwidget)
        self.redo_ps.setGeometry(QtCore.QRect(290, 150, 51, 27))
        self.redo_ps.setObjectName("redo_ps")
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
        self.label_3.setGeometry(QtCore.QRect(540, 280, 51, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtGui.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(540, 320, 51, 31))
        font = QtGui.QFont()
        font.setFamily("Lohit Devanagari")
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.checkBox = QtGui.QCheckBox(self.centralwidget)
        self.checkBox.setGeometry(QtCore.QRect(570, 690, 92, 22))
        self.checkBox.setObjectName("checkBox")
        self.radioButton = QtGui.QRadioButton(self.centralwidget)
        self.radioButton.setGeometry(QtCore.QRect(490, 620, 107, 22))
        self.radioButton.setObjectName("radioButton")
        self.pushButton_11 = QtGui.QPushButton(self.centralwidget)
        self.pushButton_11.setGeometry(QtCore.QRect(390, 680, 111, 27))
        self.pushButton_11.setObjectName("pushButton_11")
        self.name_cas = QtGui.QLineEdit(self.centralwidget)
        self.name_cas.setGeometry(QtCore.QRect(50, 20, 361, 25))
        self.name_cas.setReadOnly(True)
        self.name_cas.setObjectName("name_cas")
        self.name_ps = QtGui.QLineEdit(self.centralwidget)
        self.name_ps.setGeometry(QtCore.QRect(50, 110, 361, 25))
        self.name_ps.setReadOnly(True)
        self.name_ps.setObjectName("name_ps")
        self.preview = QtGui.QTableView(self.centralwidget)
        self.preview.setGeometry(QtCore.QRect(40, 190, 361, 421))
        self.preview.setEditTriggers(QtGui.QAbstractItemView.DoubleClicked)
        self.preview.setObjectName("preview")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtGui.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 981, 25))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtGui.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QtGui.QApplication.translate("MainWindow", "MainWindow", None, QtGui.QApplication.UnicodeUTF8))
        self.label.setText(QtGui.QApplication.translate("MainWindow", "CAS", None, QtGui.QApplication.UnicodeUTF8))
        self.open_cas.setText(QtGui.QApplication.translate("MainWindow", "Open", None, QtGui.QApplication.UnicodeUTF8))
        self.save_cas.setText(QtGui.QApplication.translate("MainWindow", "Save", None, QtGui.QApplication.UnicodeUTF8))
        self.saveas_cas.setText(QtGui.QApplication.translate("MainWindow", "Saveas", None, QtGui.QApplication.UnicodeUTF8))
        self.undo_cas.setText(QtGui.QApplication.translate("MainWindow", "Undo", None, QtGui.QApplication.UnicodeUTF8))
        self.redo_cas.setText(QtGui.QApplication.translate("MainWindow", "Redo", None, QtGui.QApplication.UnicodeUTF8))
        self.undo_ps.setText(QtGui.QApplication.translate("MainWindow", "Undo", None, QtGui.QApplication.UnicodeUTF8))
        self.saveas_ps.setText(QtGui.QApplication.translate("MainWindow", "Saveas", None, QtGui.QApplication.UnicodeUTF8))
        self.redo_ps.setText(QtGui.QApplication.translate("MainWindow", "Redo", None, QtGui.QApplication.UnicodeUTF8))
        self.label_2.setText(QtGui.QApplication.translate("MainWindow", "PS", None, QtGui.QApplication.UnicodeUTF8))
        self.save_ps.setText(QtGui.QApplication.translate("MainWindow", "Save", None, QtGui.QApplication.UnicodeUTF8))
        self.open_ps.setText(QtGui.QApplication.translate("MainWindow", "Open", None, QtGui.QApplication.UnicodeUTF8))
        self.label_3.setText(QtGui.QApplication.translate("MainWindow", "Sync", None, QtGui.QApplication.UnicodeUTF8))
        self.label_4.setText(QtGui.QApplication.translate("MainWindow", "PS file", None, QtGui.QApplication.UnicodeUTF8))
        self.checkBox.setText(QtGui.QApplication.translate("MainWindow", "CheckBox", None, QtGui.QApplication.UnicodeUTF8))
        self.radioButton.setText(QtGui.QApplication.translate("MainWindow", "RadioButton", None, QtGui.QApplication.UnicodeUTF8))
        self.pushButton_11.setText(QtGui.QApplication.translate("MainWindow", "add checkitem", None, QtGui.QApplication.UnicodeUTF8))

