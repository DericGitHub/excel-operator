import main_window
import sys
from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *


app = QApplication(sys.argv)
w = QMainWindow()
a = main_window.Ui_MainWindow()
a.setupUi(w)
w.show()
sys.exit(app.exec_())
