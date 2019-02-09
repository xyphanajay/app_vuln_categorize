#Python Script to Convert Text Based Report into Excel Matrix Sheet

import xlwt 
from xlwt import Workbook 

class sheet:
	row = 0
	col = 0
	name = ''
	wb = ''
	sheet1 = ''
	style = xlwt.easyxf('font: bold 0') 
	
	def __init__(self):
		print("Sheet class called!")
	
	def create_wb(self):
		self.wb = Workbook()
	
	def add_sheet(self):
		self.sheet1 = self.wb.add_sheet('Sheet 1')
	
	
	def add_into(self, r, c, data):
		self.sheet1.write(r, c, data, self.style)


	def save_sheet(self, name = 'example'):	
		self.wb.save(name + '.xls')
		
	def style(self, bold = '0', color = 'black'):
		self.style = xlwt.easyxf('font: bold ' + bold + ', color ' + color + ';')
		
class text:
	txt = ''
	line = ''
	lines = []
	
	def __init__(self):
		print("text class")
	
	def open_text(self, name):
		self.txt = open(name + '.txt')

	def print_txt(self):
		for self.line in self.txt:
			print(self.line)	
	
	def give_txt(self):
									#### this for loop is unable to read lines from file
		for self.line in self.txt:
			print("reading line -> ")
			print(self.line)
			self.lines.append(self.line)
		print("giving text -> ")
		print(self.lines)
		return self.lines
			
		
def main():
	report = text()
	excel = sheet()
	excel.create_wb()
	excel.add_sheet()
	report_name = input('Enter report name: ')
	report.open_text(report_name)
	report.print_txt()
	txt_data = report.give_txt()
	print("getting txt_data -> ")
	print(txt_data)
	i = 0
	j = 0
	for txt in txt_data:
		if i == 10:
			j += 1
			i = 0
		print("adding -> " + txt)
		excel.add_into(i, j, txt)
	excel.save_sheet(input('Excel file name: '))
	
main()	
	