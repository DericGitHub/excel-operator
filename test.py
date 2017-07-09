#!/usr/bin/env py

import openpyxl as xl
import time
from threading import Timer
#wb = xl.load_workbook('Swan_M227_CAS_final.xlsx')
#print wb.sheetnames
#print wb.defined_name
#ws = wb['Swan M227']
#print ws.cell(row=5,column=7).value


class jxl():
    def __init__(self,file_path):
        self.file_path = file_path
        self.workbook = xl.load_workbook(file_path)
        self.autosave_on = False
        self.save_period = None
        result ='''Inited jxl instance.
        file_path = %s
        autosave_on = %s
        save_period = %s
        '''%(self.file_path,self.autosave_on,self.save_period)
        print(result)

    
    ##################################################
    #       save 
    ##################################################
    def save(self):
        self.workbook.save(self.file_path)
    
    def save_as(self,file_path):
        self.workbook.save(file_path)
        
    ##################################################
    #       fetch 
    ##################################################
    ##################################################
    #       modification
    ##################################################
    def rewrite(self,sheet_name,start_col,iterable):
        self.workbook[sheet_name].rewrite(start_col,iterable)

    def copy_col(self,sheet_name,src,dst):
        pass
    def copy_row(self,sheet_name,src,dst):
        col = next(self.workbook[sheet_name].iter_rows(min_row=src,max_row=src))
        for i in col:
            print(i.value)
    def add_col(self,sheet_name,orientation,count):
        pass
    def del_col(self,sheet_name,orientation,count):
        pass
    def add_row(self,sheet_name,orientation,count):
        if orientation == -1:
            remain_cols = self.workbook[sheet_name].iter_rows(min_row=count,max_row=self.workbook[sheet_name].max_row)
            for i in remain_cols:
                for n in i:
                    print(n.value)
        for i in self.workbook[sheet_name].colums:
            pass

    def del_row(self,sheet_name,orientation,count):
        pass
    def add_sheet(self,sheet_name,orientation,count):
        pass
    def del_sheet(self,sheet_name,orientation,count):
        pass

    ##################################################
    #       lock
    ##################################################
    def lock(self,cells):
        pass
    def unlock(self,cells):
        pass
    
    ##################################################
    #       comparison
    ##################################################


    ##################################################
    #       autosave
    ##################################################
    def autosave(self):
        if self.autosave_on == True:
            pass 

if __name__ == "__main__":
    handler = jxl("faxtel.xlsx")
    handler.rewrite('Sheet1',1,handler.workbook['Sheet1'].iter_rows(min_row=5,max_row=5))
    handler.save()
