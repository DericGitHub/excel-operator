

class Workcell(object):
    def __init__(self,cell = None,sheet_wr = None,header = None,xmlname = None):
        self._cell = cell
        self._header = header
        self._xmlname = xmlname
        self._sheet_wr = sheet_wr




    @property
    def cell(self):
        return self._cell
    @property 
    def row(self):
        return self._cell.row
    @property
    def col(self):
        return self._cell.col_idx
    @property
    def column(self):
        return self._cell.col_idx
    @property
    def col_letter(self):
        return self._cell.column
    @property
    def value(self):
        return self._cell.value
    @value.setter
    def value(self,value):
        self._sheet_wr.api.Range(self.row,self.column).value = value
    @property
    def header(self):
        return self._header
    @property
    def xmlname(self):
        return self._xmlname
    @property
    def sheet_wr(self):
        return self._sheet_wr
 
class XmlName(Workcell):
    def __init__(self,cell = None,sheet_wr = None):
        super(XmlName,self).__init__(cell,sheet_wr)
    def get_item_by_header(self,header):
        return Workcell(self._cell.parent.cell(row = self.row,column = header.col),self.sheet_wr,header,self)


#    @property
#    def value(self):
#        if self._cell.value != None:
#            return self._cell.value
#        else:
#            return 'None'
    @property
    def subject_matter_value(self):
        pass
        #self._cell.parent.cell(row = self.row,col = 

class Header(Workcell):
    def __init(self,cell = None,sheet_wr = None):
        super(Header,self).__init__(cell,sheet_wr)

    def get_item_by_xmlname(self,xmlname):
        return Workcell(self._cell.parent.cell(row = xmlname.row,column = self.col),self.sheet_wr,self,xmlname)
        #return Workcell(self._cell.sheet.range(xmlname.row,self._cell.column),self,xmlname)
class Status(Workcell):
    def __init(self,cell = None,sheet_wr = None):
        super(Status,self).__init__(cell,sheet_wr)

    def get_item_by_xmlname(self,xmlname):
        return Workcell(self._cell.parent.cell(row = xmlname.row,column = self.col),self.sheet_wr,None,xmlname)
        #return Workcell(self._cell.sheets.range(xmlname.row,self._cell.column),None,xmlname)
       


    
 
