#-*- coding: utf-8 -*-  
from win32com.client import Dispatch, constants
import os
import os.path

filepath = "d:\\kadela\\git\\ATE_SPCC\\Temp\\" #區域網路要先設網路磁碟機代號
if os.path.isdir(filepath):
	print ("path ok:", filepath)
	filename = "YS11-17110027.csv"
	xlsfile = filepath + filename
	if os.path.isfile(xlsfile):
		print ("file ok :", xlsfile)

		xlsApp = Dispatch("Excel.Application")
		xlsApp.Visible = 1                  #顯示 Excel
		xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
		xlsSheet = xlsBook.Worksheets(u'YS11-17110027')  
		result1 = xlsSheet.Cells(16,2).Value
		print ("result is :",result1)
		xlsBook.Close()
		xlsApp.Quit()                #結束 Excel
		del xlsApp
	else:
		print ("filename fail:", xlsfile)
else:
	print ("path fail:", filepath)

#xlApp.Workbooks(1).Sheets(1).Cells(1,1).Value = "Python Rules!"