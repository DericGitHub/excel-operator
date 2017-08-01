import PSbook
import openpyxl as xl



class test(object):
    def __init__(self):
        pass
    def load(self,file_path):
        book = xl.load_workbook(file_path)
        self._workbook = PSbook.PSbook(book)
    

    @property
    def workbook(self):
        return self._workbook
