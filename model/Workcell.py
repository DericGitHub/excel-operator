

class Workcell(object):
    def __init__(self,cell = None):
        self._cell = cell




    @property 
    def row(self):
        return self._cell.row
    @property
    def col(self):
        return self._cell.col_idx
    @property
    def value(self):
        return self._cell.value
 
class XmlName(Workcell):
    def __init__(self,cell = None):
        super(XmlName,self).__init__(cell)



    @property
    def value(self):
        if self._cell.value != None:
            return self._cell.value
        else:
            return 'None'
    @property
    def subject_matter_value(self):
        pass

class Header(Workcell):
    def __init(self,cell = None):
        super(Header,self).__init__(cell)


    
 
