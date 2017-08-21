import xlwings as xw
from Workbook import Workbook
from CASsheet import CASsheet

##################################################
#       class for CAS book handling
##################################################
class CASbook(Workbook):
    
    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,file_name = None,app = None):
        super(CASbook,self).__init__(file_name,app)
        self._workbook_wr = self._workbook
        if file_name != None:
            #if app != None:
            #    self._workbook_wr = app.books.open(file_name)
            self.init_cas_book()
    def __del__(self):
        '''
            remove closing xlwings book
        '''
        #if self._workbook_wr != None:
        #    self._workbook_wr.close()
    def init_cas_book(self):
        #self.load_sheets(CASsheet,self._workbook.worksheets,self._workbook_wr.sheets)
        self.load_sheets(CASsheet,self._workbook.sheets)
        #self.load_sheets_name(self._workbook.sheetnames)
#    def load_sheets(self,sheet_cls,sheets,sheets_wr):
#        sheet_cnt = 0
#        for sheet in sheets:
#            self._sheets[sheet.title] = sheet_cls(sheet,sheets_wr[sheet.title])
#            sheet_cnt += 1
#        self._sheets_cnt = sheet_cnt
    def save_as(self,path_name):
        self._workbook.save(path_name)
        

    @property
    def workbook_wr(self):
        return self._workbook_wr
