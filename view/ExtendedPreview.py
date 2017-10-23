from PyQt4.QtGui import *
from PyQt4.Qt import *
from PyQt4.QtCore import *
import ExtendedPreviewUI
import datetime
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
    def update_extended_preview(self,array):
        model = QStandardItemModel()
        for row in array:
            line = []
            for column in row:
                #item = QStandardItem(QString(column) if column is not None else '')
                try:
                    item = QStandardItem('' if column is None else QString('%s-%s-%s'%((column+datetime.timedelta(days=1)).timetuple().tm_year,(column+datetime.timedelta(days=1)).timetuple().tm_mon,(column+datetime.timedelta(days=1)).timetuple().tm_mday)) if type(column) is datetime.datetime else QString(column))
                except BaseException as e:
                    print 'error :%s'%e
                    item = QStandardItem('Parse failed')
                line.append(item)
            model.appendRow(line)
        self.ui.extended_preview.setModel(model)
        self.ui.extended_preview.resizeRowsToContents()
