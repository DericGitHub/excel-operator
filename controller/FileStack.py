
##################################################
#       class for storing the historical file
##################################################
class FileStack(object):
    def __init__(self,max_depth = 10):
        self._max_depth = max_depth
        self._file_stack = []
        self._current_depth = 0
    

    def push(self,pack):
        if self.is_stack_full() == True:
            
        
        pass
    def pop(self,index):
        pass


    def is_stack_full(self):
        pass
    def is_stack_empty(self):
        pass
    def 


class FilePack(object):
    def __init__(self,fh = None,occupied = False,next_pack = None,last_pack = None):
        self._fh = fh
        self._occupied = occupied
        self._next_pack = next_pack
        self._last_pack = last_pack

    @property
    def fh(self):
        return self._fh
    @property
    def occupied(self):
        return self._occupied
    @property
    def next_pack(self):
        if self._next_pack.occupied = True:
            return 

        
