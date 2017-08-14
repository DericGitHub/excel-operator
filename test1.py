# -*- coding: utf-8 -*-
import openpyxl as xl

a = xl.load_workbook('Product Specification.xlsx')
b = xl.writer.excel.save_virtual_workbook(a)
f = open('1.txt','w')
f.write(b)
f.close()
