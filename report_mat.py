#Python Script to Convert Text Based Report into Excel Matrix Sheet

import xlwt 
from xlwt import Workbook 

wb = Workbook()

sheet1 = wb.add_sheet('Sheet 1')


sheet1.write(0, 0, 'First Cell of Matrix') 
sheet1.write(1, 1, 'Test Report File')

wb.save('example.xls') 
