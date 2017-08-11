import openpyxl as xl

a = xl.load_workbook('Product Specification.xlsx')
b = xl.writer.excel.save_virtual_workbook(a)
f = open('2.xlsx','w')
f.write(b)
f.close()
