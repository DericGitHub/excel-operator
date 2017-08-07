

class Workcell(object):
    def __init__(self,cell = None,header = None,xmlname = None):
        self._cell = cell
        self._header = header
        self._xmlname = xmlname




    @property 
    def row(self):
        return self._cell.row
    @property
    def col(self):
        return self._cell.col_idx
    @property
    def value(self):
        return self._cell.value
    @property
    def header(self):
        return self._header
    @property
    def xmlname(self):
        return self._xmlname
 
class XmlName(Workcell):
    def __init__(self,cell = None):
        super(XmlName,self).__init__(cell)
    def get_item_by_header(self,header):
        return Workcell(self._cell.parent.cell(row = self.row,column = header.col),header,self)


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
    def __init(self,cell = None):
        super(Header,self).__init__(cell)

    def get_item_by_xmlname(self,xmlname):
        return Workcell(self._cell.parent.cell(row = xmlname.row,column = self.col),self,xmlname)
        


    
 
