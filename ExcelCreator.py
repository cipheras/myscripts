import xlsxwriter
import ctypes
import datetime
import os
#from xlsxwriter.utility import xl_rowcol_to_cell
#CreateExcelDoc.workbook = xlsxwriter.CreateExcelDoc.workbook('demo.xlsx')
#CreateExcelDoc.worksheet = CreateExcelDoc.workbook.add_CreateExcelDoc.worksheet()

class CreateExcelDoc :
	objerrordetailslist = []
	sheet_num = 0
	exceldocrowcounter = 0
	exceldocserialnumber = 1
	excelfilename = None
	workbook = None
	worksheet = None
	#global sheet_num, CreateExcelDoc.exceldocrowcounter, CreateExcelDoc.exceldocserialnumber,CreateExcelDoc.excelfilename
	
	def createDoc(): 
		CreateExcelDoc.sheet_num+= 1
		try:
			CreateExcelDoc.workbook = xlsxwriter.Workbook(CreateExcelDoc.excelfilename)
			CreateExcelDoc.worksheet = CreateExcelDoc.workbook.add_worksheet('Sheet_'+str(CreateExcelDoc.sheet_num))
			CreateExcelDoc.createExcelHeader()
		except Exception as e:
			ctypes.windll.user32.MessageBoxW(0,e,'Error!!',0)

		
	
	def appendDoc():
		#incrementing sheet number every time append function is called
		CreateExcelDoc.sheet_num+= 1
		try:
			CreateExcelDoc.worksheet = CreateExcelDoc.workbook.add_worksheet('Sheet_'+str(CreateExcelDoc.sheet_num))
			CreateExcelDoc.createExcelHeader()
		except Exception as e:
			ctypes.windll.user32.MessageBoxW(0,e,'Error!!',0)
		

	def createHeaders(htext,f_row,f_col,l_row,l_col,mergecolumns,worksheetinteriorcolor,font,size,fcolor):
		#f_row,f_col,l_row,l_col are used instead of cell1 and cell2
		#f_row,f_col used instead of row,col
		merge_format = CreateExcelDoc.workbook.add_format({
			'font_name' : 'Arial' ,
			'font_size' : 14 ,
			'text_wrap' : True , 
			'border'	: True,
			'border_color':'black',
			'bold'		: font,
			'align' 	: 'center',
			'valign'	: 'vcenter',
			'fg_color'	: worksheetinteriorcolor                                                                     
			})
		
		if f_row is not l_row or f_col is not l_col :  #0,0,0,5
			CreateExcelDoc.worksheet.merge_range(f_row,f_col,l_row,l_col,htext,merge_format)#htext inplace of mergecolumns
		else:
			CreateExcelDoc.worksheet.write(f_row,f_col,htext,merge_format)
			CreateExcelDoc.worksheet.set_column(f_col,l_col,size)
		
		''''#background colors
		if(CreateExcelDoc.worksheetinteriorcolor == 'CORNFLOWERBLUE')
			merge_format.set_fg_color('#05ECE3')
		elif(CreateExcelDoc.worksheetinteriorcolor =='LIGHTGRAY'):
			merge_format.set_fg_color('cdc9c9')
		elif(CreateExcelDoc.worksheetinteriorcolor =='yellow'):
			merge_format.set_fg_color('yellow')
		elif(CreateExcelDoc.worksheetinteriorcolor =='grey'):
			merge_format.set_fg_color('#5E6464')
		elif(CreateExcelDoc.worksheetinteriorcolor =='gainsboro'):
			merge_format.set_fg_color('#D4DEDE')
		elif(CreateExcelDoc.worksheetinteriorcolor =='blue'):
			merge_format.set_fg_color('blue')
		elif(CreateExcelDoc.worksheetinteriorcolor =='peachpuff'):
			merge_format.set_fg_color('#F0DD7C')
		elif(CreateExcelDoc.worksheetinteriorcolor =='sienna'):
			merge_format.set_fg_color('#E57713')
		'''

		if(fcolor is ""):
			merge_format.set_font_color('white')
		else:
			merge_format.set_font_color('black')


	def createExcelHeader():
		#(con_row,con_col) = xl_cell_to_rowcol('A1')
		CreateExcelDoc.exceldocrowcounter = 0
		CreateExcelDoc.createHeaders('Test Results',CreateExcelDoc.exceldocrowcounter,0,CreateExcelDoc.exceldocrowcounter,5,6,'#F0DD7C',False,100,'n')
		CreateExcelDoc.exceldocrowcounter+=1

		CreateExcelDoc.createHeaders('S.No',CreateExcelDoc.exceldocrowcounter,0,CreateExcelDoc.exceldocrowcounter,0, 0,'cdc9c9',True,7,'')
		CreateExcelDoc.createHeaders('APPLICATION NAME & VERSION',CreateExcelDoc.exceldocrowcounter,1,CreateExcelDoc.exceldocrowcounter,1, 0,'cdc9c9',True,30,'')
		CreateExcelDoc.createHeaders('USER NAME',CreateExcelDoc.exceldocrowcounter,2,CreateExcelDoc.exceldocrowcounter,2, 0,'cdc9c9',True,30,'')
		CreateExcelDoc.createHeaders('DESCRIPTION',CreateExcelDoc.exceldocrowcounter,3,CreateExcelDoc.exceldocrowcounter,3, 0,'cdc9c9',True,60,'')
		CreateExcelDoc.createHeaders('STATUS',CreateExcelDoc.exceldocrowcounter,4,CreateExcelDoc.exceldocrowcounter,4, 0,'cdc9c9',True,45,'')
		CreateExcelDoc.createHeaders('DATE TESTED',CreateExcelDoc.exceldocrowcounter,5,CreateExcelDoc.exceldocrowcounter,5, 0,'cdc9c9',True,15,'')

		CreateExcelDoc.exceldocrowcounter+=1

	
	def addData(data,f_row,f_col,l_row,l_col,n_format,fail):
		#f_row,f_col used instead of row,col
		data_format = CreateExcelDoc.workbook.add_format({
			'border'	: True,
			'border_color' : 'black',
			'align' 	: 'center',
			'valign'	: 'vcenter',
			'num_format' : n_format,
			'text_wrap' : True,
			'font_name' : 'Arial',
			'font_size' : 10
			})
		CreateExcelDoc.worksheet.write(f_row,f_col,data,data_format)
		'''if not f_row==l_row and f_col==l_col:	
			CreateExcelDoc.worksheet.merge_range(f_row,f_col,l_row,l_col,data,data_format)

		if (col==0 or col==1 or col==2 or col==3 or col==4 or col==5):
			data_format.set_align('center')
			data_format.set_align('top')
		'''
		if (fail):
			data_format.set_font_color('#DB490E')#tomato color
		else:
			data_format.set_font_color('#37890C')#green color

	
	def addExcelData(testurl,testedurl,username,teststatus,fail=False):
		#bool fail
		#fail = False
		datevalue = ''+ str(datetime.date.today().strftime('%m-%d-%Y'))+''
		CreateExcelDoc.addData(CreateExcelDoc.exceldocserialnumber,CreateExcelDoc.exceldocrowcounter,0,CreateExcelDoc.exceldocrowcounter,0,"",fail)
		CreateExcelDoc.addData(username,CreateExcelDoc.exceldocrowcounter,1,CreateExcelDoc.exceldocrowcounter,1,"",fail)
		CreateExcelDoc.addData(testurl,CreateExcelDoc.exceldocrowcounter,2,CreateExcelDoc.exceldocrowcounter,2,"",fail)
		CreateExcelDoc.addData(testedurl,CreateExcelDoc.exceldocrowcounter,3,CreateExcelDoc.exceldocrowcounter,3,"",fail)
		CreateExcelDoc.addData(teststatus,CreateExcelDoc.exceldocrowcounter,4,CreateExcelDoc.exceldocrowcounter,4,"",fail)
		CreateExcelDoc.addData(datevalue,CreateExcelDoc.exceldocrowcounter,5,CreateExcelDoc.exceldocrowcounter,5,"",fail)

		CreateExcelDoc.exceldocrowcounter+=1
		CreateExcelDoc.exceldocserialnumber+=1

		if (fail):
			CreateExcelDoc.addFailTestCasesToArrayList(None,None,None,datevalue,teststatus,testurl)

	
	def createDirectory(username,dtype,filename): #dtype is used instead of type
		# filename = ''   
		outputfilenametimeflag = str(datetime.datetime.now().strftime('%H-%M-%S'))+"_"+username
		newpath = r'C:\\Users\\abhishek.singh2\\Desktop\\'+filename
		
		if dtype.lower() == 'info':
			strdirname = newpath +'\\'+filename+'_OutputResult_'+ str(datetime.date.today().strftime('%m-%d-%Y'))

		if len(filename) <= 0 :
			if (dtype.lower() == 'info'):
				filename =  'NoName_'
		else:
			filename = filename + '_'

		CreateExcelDoc.excelfilename = strdirname + '\\' + filename + outputfilenametimeflag + '.xlsx'
		
		try:
			if not os.path.exists(strdirname):
				os.makedirs(strdirname)
		except Exception as e:
			ctypes.windll.user32.MessageBoxW(0,e,'Error!!',0)

		# CreateExcelDoc.saveworkbook(filename,username,dtype)

	'''def saveworkbook(filename,username,dtype):
		#filename = ''
		#username = ''
		#dtype = ''
		CreateExcelDoc.createDirectory(username,dtype,filename)
		#CreateExcelDoc.workbook.close()
	'''
	def closeWorkbook():
		CreateExcelDoc.workbook.close()

	
	def addFailTestCasesToArrayList(testscriptname,mappingcolumnnumber,functionalitytested,datevalue,teststatus,testurl=None):  
		if testurl:	
			failtestcases = testurl+'#'+teststatus+'#'+datevalue
		else:
			failtestcases = testscriptname+'#'+mappingcolumnnumber+'#'+functionalitytested+'#'+datevalue+'#'+teststatus
			CreateExcelDoc.objerrordetailslist.append(failtestcases)
		# print(CreateExcelDoc.objerrordetailslist)
			

	def addFailTestCases():
		counter = 0 
		exceldocumentrowcounter = CreateExcelDoc.exceldocrowcounter
		exceldocumentserialnumber = CreateExcelDoc.exceldocserialnumber
		while counter < len(CreateExcelDoc.objerrordetailslist):
			failtestcasedata = CreateExcelDoc.objerrordetailslist[counter].split('#')
			if len(failtestcasedata)>=5 :
				testscriptname = failtestcasedata[0]
				mappingcolumnnumber = failtestcasedata[1]
				functionalitytested = failtestcasedata[2]
				datevalue = failtestcasedata[3]
				teststatus = failtestcasedata[4]

				CreateExcelDoc.addData(exceldocumentserialnumber,exceldocumentrowcounter,0,exceldocumentrowcounter,0,'',True)
				CreateExcelDoc.addData(testscriptname,exceldocumentrowcounter,1,exceldocumentrowcounter,1,'',True)
				CreateExcelDoc.addData(mappingcolumnnumber,exceldocumentrowcounter,2,exceldocumentrowcounter,2,'',True)
				CreateExcelDoc.addData(functionalitytested,exceldocumentrowcounter,3,exceldocumentrowcounter,3,'',True)
				CreateExcelDoc.addData(datevalue,exceldocumentrowcounter,4,exceldocumentrowcounter,4,'',True)
				CreateExcelDoc.addData(teststatus,exceldocumentrowcounter,5,exceldocumentrowcounter,5,'',True)

			exceldocumentrowcounter+=1
			exceldocumentserialnumber+=1
			counter+=1
			 
	

if __name__ == '__main__':
	CreateExcelDoc.createDoc()