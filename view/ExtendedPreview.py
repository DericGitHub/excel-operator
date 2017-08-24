from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
import ExtendedPreviewUI
class ExtendedPreview(QMainWindow):
    def __init__(self,model = None):
        super(ExtendedPreview,self).__init__()
        self.init_Form()
        if model != None:
            self.init_model(model)
    def init_Form(self):
        self.ui = ExtendedPreviewUI.Ui_Form()
        self.ui.setupUi(self)
    def init_model(self,model):
        self.ui.extended_preview.setModel(model)
        self.ui.extended_preview.resizeRowsToContents()
