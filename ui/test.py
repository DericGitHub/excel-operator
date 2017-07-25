import main_window
import sys
from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *

class My_MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.init_My_MainWindow()
    def init_My_MainWindow(self):
        self.ui = main_window.Ui_MainWindow()
        self.ui.setupUi(self)
        self.model = QStandardItemModel()
        item = QStandardItem('asdf')
        item.setCheckState(False)
        item.setCheckable(True)
        self.model.appendRow(item)
        self.ui.listView.setModel(self.model)

#app = QApplication(sys.argv)
#w = QMainWindow()
#a = main_window.Ui_MainWindow()
#a.setupUi(w)
#w.show()
#sys.exit(app.exec_())
app = QApplication(sys.argv)
w = My_MainWindow()
w.show()
sys.exit(app.exec_())
