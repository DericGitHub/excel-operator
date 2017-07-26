# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created: Fri Jul 14 14:47:36 2017
#      by: PyQt4 UI code generator 4.6.2
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui
from PyQt4.QtCore import Qt
import sys

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtGui.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtGui.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(280, 230, 90, 27))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.aaa)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtGui.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 25))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtGui.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
    def aaa(self):
        self.filedialog = QtGui.QFileDialog()
        self.filedialog.show()
        #lineEdit.setText(QFileDialog.getOpenFileName())

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QtGui.QApplication.translate("MainWindow", "MainWindow", None, QtGui.QApplication.UnicodeUTF8))
        self.pushButton.setText(QtGui.QApplication.translate("MainWindow", "PushButton", None, QtGui.QApplication.UnicodeUTF8))

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    w = QtGui.QMainWindow()
    w.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)

    a = Ui_MainWindow()
    a.setupUi(w)
    w.show()

    sys.exit(app.exec_())
