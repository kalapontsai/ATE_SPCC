#-*- coding: utf-8 -*-  
from win32com.client import Dispatch, constants
import os


def get_yield( filepath = "d:\\kadela\\git\\ATE_SPCC\\sample\\", filename = "YS11-17110027.csv",sampling = 0):
	#filepath = "d:\\kadela\\git\\ATE_SPCC\\sample\\" #區域網路要先設網路磁碟機代號
	#filename = "YS11-17110027.csv"
	if os.path.isdir(filepath):
		#print ("path ok:", filepath)
		xlsfile = os.path.join(filepath,filename)
		if os.path.isfile(xlsfile):
			#print ("file ok :", xlsfile)

			xlsApp = Dispatch("Excel.Application")
			xlsApp.Visible = 0                  #顯示 Excel
			xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
			lot_id = os.path.splitext(filename)[0]
			#print ("sheetname :", lot_id)
			xlsSheet = xlsBook.Worksheets(lot_id)  

			row = 16   #列
			col = 2 #欄
			total = 0
			good = 0
			result = []
			if sampling == 0 : sampling = 999999
			while (xlsSheet.Cells(row,col).Value is not None) and (total < sampling):
				result.append(xlsSheet.Cells(row,col).Value)
				if 'PASS' in result[total]:
					good += 1
				row += 1
				total += 1

			#print ("Total is :", total)
			#print ("Pass is  :", good)
			#print (result)
			#print ('ABORT' in result[0])
			xlsBook.Close()
			xlsApp.Quit()                #結束 Excel
			del xlsApp
		else:
			print ("filename fail:", xlsfile)
	else:
		print ("path fail:", filepath)
	return (lot_id,total,good)

#get_yield(filepath = "d:\\kadela\\git\\ATE_SPCC\\sample\\", filename = "YS11-17110027.csv")
