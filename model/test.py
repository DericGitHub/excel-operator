import PSbook
import openpyxl as xl



class test(object):
    def __init__(self):
        pass
    def load(file_path):
        book = xl.load(file_path)
        self._workbook = PSbook(book)
    

    @property
    def workbook(self):
        return self._workbook
