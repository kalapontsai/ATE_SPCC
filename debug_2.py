# try to write the stat value into SPCC templated file

import os
from win32com.client import Dispatch, constants

xlsfile = "D:\\Kadela\\git\\ATE_SPCC\\sample\\XR.xlsx"
xlsApp = Dispatch("Excel.Application")
xlsApp.Visible = 0                  #dont show Excel
xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
xlsSheet = xlsBook.Worksheets("p-chart")  

total_row = 6   #列
sample_row = 7
col = 3 #欄

xlsSheet.Cells(total_row,col).Value = '100'
xlsSheet.Cells(sample_row,col).Value = '5'
xlsBook.Save()
xlsBook.Close()
xlsApp.Quit()                #結束 Excel
del xlsApp
