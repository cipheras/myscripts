import xlsxwriter
import ctypes
import os
import datetime
import itertools

class expense(object):

	def __init__(self):
		self.sheet_num = 0
		# self.excelfilename = None
		self.workbook = None
		self.worksheet = None
		self.filename=input('Enter file name: ')
		self.names = input('Enter name of the persons: ').split(',')
		self.all = []

	def createFile(self):
		self.sheet_num += 1
		try:
			self.workbook = xlsxwriter.Workbook(self.filename)
			self.worksheet = self.workbook.add_worksheet('sheet_'+str(self.sheet_num))
		except Exception as e:
			ctypes.windll.user32.MessageBoxW(0,e,'Error!!',0)

	def who(self):
		for n in range(1,len(self.names)+1):
			for comb in itertools.permutations(self.names,n):
				print(comb)
				self.all.append(comb)
		print(self.all)

	def header(self,hdata,row,col,size):
		format = self.workbook.add_format({
    		'font_name' : 'Arial' ,
			'font_size' : 12 ,
			'text_wrap' : False ,
			'border'	: True,
			'border_color':'black',
			'bold'		: True,
			'align' 	: 'center',
			'valign'	: 'vcenter',
			'fg_color'	: 'green'                                                                     
			})
		
		self.worksheet.write(row,col,hdata,format)
		self.worksheet.set_column(col,col,size)
		
	
	def headerData(self):
		self.header('S.no',0,0,8)
		self.header('Item',0,1,35)
		self.header('Amount',0,2,12)
		self.header('Paid by',0,3,15)
		self.header('x->y',0,4,20)
		self.header('x->z',0,5,20)
		self.header('y->z',0,6,20)
		self.header('y->x',0,7,20)
		self.header('z->y',0,8,20)
		self.header('z->x',0,9,20)
		

	def createDirectory(self):  
		newpath = r'\Expense Acc'
		self.filename = newpath+'\\'+self.filename+'.xlsx'
		try:
			if not os.path.exists(newpath):
				os.makedirs(newpath)
		except Exception as e:
			ctypes.windll.user32.MessageBoxW(0,e,'Error!!',0)

	def closeWorkbook(self):
		self.workbook.close()

if __name__ == '__main__':
	obj = expense()
	obj.createDirectory()
	obj.createFile()
	obj.who()
	obj.headerData()
	obj.closeWorkbook()
