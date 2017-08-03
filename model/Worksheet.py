

##################################################
#       Abstract class to handle worksheet
##################################################
class Worksheet(object):

    ##################################################
    #       Initial method
    ##################################################
    def __init__(self,sheet = None,xmlname_coordinate = None):
        self._worksheet = sheet
        self._rows = []
        self._min_row = 0
        self._max_row = None
        self._cols = []
        self._min_col = 0
        self._max_col = None
        self._xmlname_coordinate = xmlname_coordinate
        self._title = {}
        

    def locate_xmlname(self):
        for row in self._rows:
            for cell in row:
                if cell.value == 'xmlname':
                    return cell.co
    def load_rows(self,rows):
        for row in rows:
            self._rows.append(row)

    def load_cols(self,cols):
        for col in cols:
            self._cols.append(col)
    def load_title(self):
        pass
        

        



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
    @property
    def max_row(self):
        return self._max_row
    @property
    def min_row(self):
        return self._min_row
    @property
    def max_col(self):
        return self._max_col
    @property
    def min_col(self):
        return self._min_col
