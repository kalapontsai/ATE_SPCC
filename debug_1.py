import os
from win32com.client import Dispatch, constants

xlsfile = "D:\\Kadela\\git\\ATE_SPCC\\sample\\YS11-17110099.csv"
xlsApp = Dispatch("Excel.Application")
xlsApp.Visible = 0                  #顯示 Excel
xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
lot_id = "YS11-17110099"     
#print ("sheetname :", lot_id)
xlsSheet = xlsBook.Worksheets(lot_id)  

row = 16   #列
col = 2 #欄
total = 0
good = 0
#result = []
answer = ['PASS','ABORT','FAIL']
print (xlsSheet.Cells(16,2))
print (xlsSheet.Cells(5,4))
print (type(xlsSheet.Cells(16,2).Value) == type(None))
