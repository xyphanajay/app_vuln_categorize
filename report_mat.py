#Python Script to Convert Text Based Report into Excel Matrix Sheet
import os

import xlwt 
from xlwt import Workbook 

class sheet:
	row = 0
	col = 0
	name = ''
	wb = ''
	sheet1 = ''
	s_no = 1
	style = xlwt.easyxf('font: bold 0, color black;') 
	
	def __init__(self):
		print("Sheet class called!")
		self.wb = Workbook()
		self.sheet1 = self.wb.add_sheet('Sheet 1', cell_overwrite_ok=True)
	
	def add_sheet(self):
		self.s_no += 1
		self.sheet = self.wb.add_sheet('Sheet ' + str(self.s_no), cell_overwrite_ok=True)
	
	def add_into(self, r, c, data):
		self.sheet1.write(r, c, data)

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
		try:
			self.txt = open(name + '.txt')
			return 1
		except IOError:
			try:
				self.txt = open(name)
				return 1
			except IOError:
				print("File not found!")
				return 0

	def print_txt(self):
		self.txt.seek(0)
		for self.line in self.txt:
			print(self.line)	
	
	def give_txt(self):
									#### this for loop is unable to read lines from file
		self.txt.seek(0)
		for self.line in self.txt:
			#print("reading line -> ")
			#print(self.line)
			self.lines.append(self.line)
		#print("giving text -> ")
		#print(self.lines)
		return self.lines
		
class reports:
	dirc = []
	def __init__(self, dirc = ''):
		self.dirc = os.listdir(dirc)
		
	def show_files(self):
		for l in self.dirc:
			print (l)

		
def main():
	excel = sheet()
	i = 0
	j = 0
	print("List file in report ->")
	rep = reports("./report")
	rep.show_files()
	f = text()
	for r in rep.dirc:
		#print(r)
		if (f.open_text('./report/' + r)):
			pass
		else:
			print("Filed to open!")
			exit()
		#f.print_txt()
		txt_data = f.give_txt()
	
	#exit()
		for txt in txt_data:
			if i == 10:
				j += 1
				i = 0
			#print("adding -> " + txt)
			excel.add_into(i, j, txt)
			i += 1
		excel.save_sheet(r)
		print("Sheet saved: " + r)
main()
	