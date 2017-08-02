# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\untitled.ui'
#
# Created: Wed Aug 02 22:29:11 2017
#      by: PyQt4 UI code generator 4.10
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(981, 790)
        self.centralwidget = QtGui.QWidget(MainWindow)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.label = QtGui.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 20, 31, 31))
        font = QtGui.QFont()
        font.setFamily(_fromUtf8("Lohit Devanagari"))
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName(_fromUtf8("label"))
        self.open_cas = QtGui.QPushButton(self.centralwidget)
        self.open_cas.setGeometry(QtCore.QRect(50, 60, 51, 27))
        self.open_cas.setObjectName(_fromUtf8("open_cas"))
        self.sheets_cas = QtGui.QComboBox(self.centralwidget)
        self.sheets_cas.setGeometry(QtCore.QRect(460, 20, 241, 31))
        self.sheets_cas.setObjectName(_fromUtf8("sheets_cas"))
        self.save_cas = QtGui.QPushButton(self.centralwidget)
        self.save_cas.setGeometry(QtCore.QRect(110, 60, 51, 27))
        self.save_cas.setObjectName(_fromUtf8("save_cas"))
        self.saveas_cas = QtGui.QPushButton(self.centralwidget)
        self.saveas_cas.setGeometry(QtCore.QRect(170, 60, 51, 27))
        self.saveas_cas.setObjectName(_fromUtf8("saveas_cas"))
        self.undo_cas = QtGui.QPushButton(self.centralwidget)
        self.undo_cas.setGeometry(QtCore.QRect(230, 60, 51, 27))
        self.undo_cas.setObjectName(_fromUtf8("undo_cas"))
        self.redo_cas = QtGui.QPushButton(self.centralwidget)
        self.redo_cas.setGeometry(QtCore.QRect(290, 60, 51, 27))
        self.redo_cas.setObjectName(_fromUtf8("redo_cas"))
        self.undo_ps = QtGui.QPushButton(self.centralwidget)
        self.undo_ps.setGeometry(QtCore.QRect(230, 150, 51, 27))
        self.undo_ps.setObjectName(_fromUtf8("undo_ps"))
        self.saveas_ps = QtGui.QPushButton(self.centralwidget)
        self.saveas_ps.setGeometry(QtCore.QRect(170, 150, 51, 27))
        self.saveas_ps.setObjectName(_fromUtf8("saveas_ps"))
        self.redo_ps = QtGui.QPushButton(self.centralwidget)
        self.redo_ps.setGeometry(QtCore.QRect(290, 150, 51, 27))
        self.redo_ps.setObjectName(_fromUtf8("redo_ps"))
        self.sheets_ps = QtGui.QComboBox(self.centralwidget)
        self.sheets_ps.setGeometry(QtCore.QRect(460, 100, 241, 31))
        self.sheets_ps.setObjectName(_fromUtf8("sheets_ps"))
        self.label_2 = QtGui.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(10, 110, 31, 31))
        font = QtGui.QFont()
        font.setFamily(_fromUtf8("Lohit Devanagari"))
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName(_fromUtf8("label_2"))
        self.save_ps = QtGui.QPushButton(self.centralwidget)
        self.save_ps.setGeometry(QtCore.QRect(110, 150, 51, 27))
        self.save_ps.setObjectName(_fromUtf8("save_ps"))
        self.open_ps = QtGui.QPushButton(self.centralwidget)
        self.open_ps.setGeometry(QtCore.QRect(50, 150, 51, 27))
        self.open_ps.setObjectName(_fromUtf8("open_ps"))
        self.label_3 = QtGui.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(540, 280, 51, 31))
        font = QtGui.QFont()
        font.setFamily(_fromUtf8("Lohit Devanagari"))
        font.setPointSize(12)
        self.label_3.setFont(font)
        self.label_3.setObjectName(_fromUtf8("label_3"))
        self.label_4 = QtGui.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(540, 320, 51, 31))
        font = QtGui.QFont()
        font.setFamily(_fromUtf8("Lohit Devanagari"))
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setObjectName(_fromUtf8("label_4"))
        self.checkBox = QtGui.QCheckBox(self.centralwidget)
        self.checkBox.setGeometry(QtCore.QRect(570, 690, 92, 22))
        self.checkBox.setObjectName(_fromUtf8("checkBox"))
        self.radioButton = QtGui.QRadioButton(self.centralwidget)
        self.radioButton.setGeometry(QtCore.QRect(490, 620, 107, 22))
        self.radioButton.setObjectName(_fromUtf8("radioButton"))
        self.pushButton_11 = QtGui.QPushButton(self.centralwidget)
        self.pushButton_11.setGeometry(QtCore.QRect(390, 680, 111, 27))
        self.pushButton_11.setObjectName(_fromUtf8("pushButton_11"))
        self.name_cas = QtGui.QLineEdit(self.centralwidget)
        self.name_cas.setGeometry(QtCore.QRect(50, 20, 361, 25))
        self.name_cas.setReadOnly(True)
        self.name_cas.setObjectName(_fromUtf8("name_cas"))
        self.name_ps = QtGui.QLineEdit(self.centralwidget)
        self.name_ps.setGeometry(QtCore.QRect(50, 110, 361, 25))
        self.name_ps.setReadOnly(True)
        self.name_ps.setObjectName(_fromUtf8("name_ps"))
        self.preview = QtGui.QTableView(self.centralwidget)
        self.preview.setGeometry(QtCore.QRect(40, 190, 361, 421))
        self.preview.setObjectName(_fromUtf8("preview"))
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtGui.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 981, 26))
        self.menubar.setObjectName(_fromUtf8("menubar"))
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtGui.QStatusBar(MainWindow)
        self.statusbar.setObjectName(_fromUtf8("statusbar"))
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow", None))
        self.label.setText(_translate("MainWindow", "CAS", None))
        self.open_cas.setText(_translate("MainWindow", "Open", None))
        self.save_cas.setText(_translate("MainWindow", "Save", None))
        self.saveas_cas.setText(_translate("MainWindow", "Saveas", None))
        self.undo_cas.setText(_translate("MainWindow", "Undo", None))
        self.redo_cas.setText(_translate("MainWindow", "Redo", None))
        self.undo_ps.setText(_translate("MainWindow", "Undo", None))
        self.saveas_ps.setText(_translate("MainWindow", "Saveas", None))
        self.redo_ps.setText(_translate("MainWindow", "Redo", None))
        self.label_2.setText(_translate("MainWindow", "PS", None))
        self.save_ps.setText(_translate("MainWindow", "Save", None))
        self.open_ps.setText(_translate("MainWindow", "Open", None))
        self.label_3.setText(_translate("MainWindow", "Sync", None))
        self.label_4.setText(_translate("MainWindow", "PS file", None))
        self.checkBox.setText(_translate("MainWindow", "CheckBox", None))
        self.radioButton.setText(_translate("MainWindow", "RadioButton", None))
        self.pushButton_11.setText(_translate("MainWindow", "add checkitem", None))

