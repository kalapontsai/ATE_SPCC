#-*- coding: utf-8 -*-  
from win32com.client import Dispatch, constants
import os

xlsApp = Dispatch("Excel.Application")
xlsApp.Visible = 1                  #顯示 Excel
xlsBook = xlsApp.Workbooks.Add()    #新增一工作簿
xlsSheet = xlsBook.Worksheets(u'工作表1')  #新增的工作簿預設含三個工作表

for i in range(5):                    #新增兩列資料
    xlsSheet.Cells(i+1,1).Value = (i+1)
    xlsSheet.Cells(i+1,2).Value = (i+1)

xlsSheet.Cells(1,1).Font.Color = 0xff0000

testfile = 'D:\\kadela\\git\\ATE_SPCC\\1.xlsx'
if os.path.isfile(testfile):
    os.remove(testfile)
xlsBook.SaveAs(testfile)      #存檔 xlsBook.SaveAs(testfile,FileFormat= 56)
xlsBook.Close()                #關閉工作簿
xlsApp.Quit()                #結束 Excel
del xlsApp