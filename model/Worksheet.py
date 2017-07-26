

##################################################
#       Abstract class to handle worksheet
##################################################
class Worksheet(object):

    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,sheet = None):
        self._worksheet = sheet
        self._rows = []
        self._row_min = 0
        self._row_max = None
        self._cols = []
        self._col_min = 0
        self._col_max = None

    def load_rows(self,rows):
        row_cnt = 0
        for row in rows:
            self._rows.append(row)
            row_cnt += 1
        self._row_max = row_cnt

    def load_cols(self,cols):
        col_cnt = 0
        for col in cols:
            self._cols.append(col)
            col_cnt += 1
        self._col_max = col_cnt

        



    ##################################################
    #       User interface
    ##################################################
    def del_row(self,row_index):
        pass
    def append_row(self,row_index):
        pass
    def title(self):
        pass

    ##################################################
    #       Data
    ##################################################
    @property
    def rows(self):
        return self._rows
    @property
    def cols(self):
        return self._cols
