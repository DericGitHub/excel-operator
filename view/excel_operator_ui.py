# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled1.ui'
#
# Created: Fri Jul 14 16:45:57 2017
#      by: PyQt4 UI code generator 4.6.2
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(397, 215)
        self.comboBox = QtGui.QComboBox(Form)
        self.comboBox.setGeometry(QtCore.QRect(10, 10, 331, 21))
        self.comboBox.setObjectName("comboBox")
        self.comboBox_2 = QtGui.QComboBox(Form)
        self.comboBox_2.setGeometry(QtCore.QRect(10, 40, 331, 21))
        self.comboBox_2.setObjectName("comboBox_2")
        self.pushButton = QtGui.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(350, 10, 41, 21))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtGui.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(350, 40, 41, 21))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtGui.QPushButton(Form)
        self.pushButton_3.setGeometry(QtCore.QRect(100, 80, 171, 111))
        self.pushButton_3.setObjectName("pushButton_3")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        Form.setWindowTitle(QtGui.QApplication.translate("Form", "Excel Operator", None, QtGui.QApplication.UnicodeUTF8))
        self.pushButton.setText(QtGui.QApplication.translate("Form", "PS", None, QtGui.QApplication.UnicodeUTF8))
        self.pushButton_2.setText(QtGui.QApplication.translate("Form", "CAS", None, QtGui.QApplication.UnicodeUTF8))
        self.pushButton_3.setText(QtGui.QApplication.translate("Form", "Transmit", None, QtGui.QApplication.UnicodeUTF8))

