# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled1.ui'
#
# Created: Mon Sep 11 14:34:31 2017
#      by: PyQt4 UI code generator 4.6.2
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1600, 850)
        self.extended_preview = QtGui.QTableView(Form)
        self.extended_preview.setGeometry(QtCore.QRect(0, 0, 1600, 800))
        self.extended_preview.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        self.extended_preview.setTabKeyNavigation(True)
        self.extended_preview.setObjectName("extended_preview")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        Form.setWindowTitle(QtGui.QApplication.translate("Form", "Form", None, QtGui.QApplication.UnicodeUTF8))

