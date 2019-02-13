#Python Script to Convert Text Based Report into Excel Matrix Sheet
import os
import re
import datetime

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
		self.txt.seek(0)
		self.lines = []
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
			print (rename(l))
class picker:
	ct = '[Critical]'
	wn = '[Warning]'
	ct_l = []
	wn_l = []
	def pick_lines(self, txt):
		self.ct_l = []
		self.wn_l = []
		for line in txt:
			if self.ct in line:
				self.ct_l.append(line)
			if self.wn in line:
				self.wn_l.append(line)
		return self.ct_l, self.wn_l
			
	def split_cat(self, data):
		# case 1: [Critical]  Accessing the Internet Checking: ---> ['', 'Critical', ' AndroidManifest ContentProvider Exported Checking', '']
		# case 2: [Critical] <Implicit_Intent> Implicit Service Checking: ---> ['', 'Critical', '', 'Implicit_Intent', ' Implicit Service Checking', '']
		# case 3: [Critical] <WebView><Remote Code Execution><#CVE-2013-4710#> WebView RCE Vulnerability Checking: ---> ['', 'Critical', '', 'WebView', '', 'Remote Code Execution', '', '#CVE-2013-4710#', ' WebView RCE Vulnerability Checking', '']
		# dealing: Tag1: Critical Tag2: None | (WebView) Descri: (Getting Android_id)
		# lenth: 1 tag: 4; 2 tags: 6; 3 tags: 8(rare); 4 tags: 10
		# picks: s[3] of all tags > 1; s[n - 2] for n tag
		split = re.split("\[|\] | <|<|>|> |:", data)
		if len(split) == 4:
			return split[2], None
		else:
			return split[3], split[len(split) - 2]
			
def rename(name):
	return re.split(r"_", name)[0]
		
def main():
	excel = sheet()						# creating sheet
	count = 0
	cnt = []
	app_no = 0
	vul_no = 0
	temp = ''
	ct_l = []							# critical lines
	wn_l = []							# warning lines
	critical = []
	critical_all = []
	warning = []
	print("List file in report ->")
	rep = reports("./report")			# creating obj for report files
	pick = picker()						# creating obj for line picker
	f = text()							# creating text obj
	#rep.show_files()					# showing list of report files
	for r in rep.dirc:
	
		if (f.open_text('./report/' + r)):
			print("Report accessing: " + rename(r))
			app_no += 1
			excel.add_into(0, app_no, rename(r))
			pass
		else:
			print("Filed to open!")
			exit()
		critical = []							# clearing critical old report info
		txt_data = f.give_txt()					# getting report text
		ct_l = []
		ct_l, wn_l = pick.pick_lines(txt_data)	# getting critical and warning of each report
		for data in ct_l:
			tag, des = pick.split_cat(data)
			if des is None:
				temp = tag
			else:
				temp = tag + ":" + des
			critical.append(temp)
			# adding and checking vul in critical_main
			if temp not in critical_all:
				critical_all.append(temp)
				vul_no = critical_all.index(temp) + 1
				excel.add_into(vul_no, 0, temp)
		cnt.append(len(critical_all))
		# marking 0 1 in app report
		for vuln in critical_all:
			vul_no = critical_all.index(vuln) + 1
			if vuln in critical:
				excel.add_into(vul_no, app_no, 1)
			else:
				excel.add_into(vul_no, app_no, 0)
	# filling rest with zeros
	for n in cnt:
		count += 1
		while n < len(critical_all):
			n += 1
			excel.add_into(n, count, 0)
			
	print("=" * 70)
	print("Total Critical cases: " + str(len(critical_all)))
	for item in critical_all:
		print(item)
	excel.save_sheet("Final_report " + str(datetime.datetime.now()))
main()
	