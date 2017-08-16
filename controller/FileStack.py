from io import BytesIO,StringIO
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
            self._file_stack.pop(0)
        self._file_stack.append(pack)
            
        
        pass
    def pop(self):
        if self.is_stack_empty() != True:
            if self.len != 1:
                return (self._file_stack.pop().action,self._file_stack[-1])
            else:
                return None
        else:
            return None
            


    def is_stack_full(self):
        if len(self._file_stack) >= self._max_depth:
            return True
        else:
            return False
    def is_stack_empty(self):
        if len(self._file_stack) == 0:
            return True
        else:
            return False
    
    @property
    def fileStack(self):
        return self._file_stack
    @property
    def len(self):
        return len(self._file_stack)
    @property
    def currentFile(self):
        return self._file_stack[-1]


class FilePack(object):
    def __init__(self,action = None,content = None):
        self._action = action
        self._fh = BytesIO()
        if content != None:
            self._fh.write(content)
    def __del__(self):
        self._fh.close()

    @property
    def action(self):
        return self._action
    @property
    def fh(self):
        return self._fh
class CasPack(object):
    def __init__(self,action = None,fileName = None):
        self._action = action
        self._file_name = fileName
    @property
    def action(self):
        return self._action
    @property
    def file_name(self):
        return self._file_name

        
