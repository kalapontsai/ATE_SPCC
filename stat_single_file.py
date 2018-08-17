#-*- coding: utf-8 -*-  
from win32com.client import Dispatch, constants
import os

def get_yield( filepath = "d:\\kadela\\git\\ATE_SPCC\\sample\\", filename = "YS11-17110027.csv"):
	#filepath = "d:\\kadela\\git\\ATE_SPCC\\sample\\" #區域網路要先設網路磁碟機代號
	#filename = "YS11-17110027.csv"
	if os.path.isdir(filepath):
		print ("path ok:", filepath)
		xlsfile = filepath + filename
		if os.path.isfile(xlsfile):
			print ("file ok :", xlsfile)

			xlsApp = Dispatch("Excel.Application")
			xlsApp.Visible = 0                  #顯示 Excel
			xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
			sheetname = filename.upper()[:-4]
			print ("sheetname :", sheetname)
			xlsSheet = xlsBook.Worksheets(sheetname)  

			row = 16   #列
			col = 2 #欄
			total = 0
			good = 0
			result = []
			while (xlsSheet.Cells(row,col).Value is not None) and (total < 100):
				result.append(xlsSheet.Cells(row,col).Value)
				if 'PASS' in result[total]:
					good += 1
				row += 1
				total += 1

			print ("Total is :", total)
			print ("Pass is  :", good)
			#print (result)
			#print ('ABORT' in result[0])
			xlsBook.Close()
			xlsApp.Quit()                #結束 Excel
			del xlsApp
		else:
			print ("filename fail:", xlsfile)
	else:
		print ("path fail:", filepath)

#get_yield(filepath = "d:\\kadela\\git\\ATE_SPCC\\sample\\", filename = "YS11-17110027.csv")
