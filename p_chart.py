#對指定目錄內所有測試檔進行統計, 紀錄lot id, 總數, 良品數
#再寫入p-chart樣本檔

import os
from win32com.client import Dispatch, constants

def get_yield( filepath = "d:\\kadela\\git\\ATE_SPCC\\sample\\", filename = "YS11-17110027.csv",sampling = 0):
	xlsfile = os.path.join(filepath,filename)
	xlsApp = Dispatch("Excel.Application")
	xlsApp.Visible = 0                  #顯示 Excel
	xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
	lot_id = os.path.splitext(filename)[0]      #分離前檔名及副檔名
	xlsSheet = xlsBook.Worksheets(lot_id)  
	row = 16   #列
	col = 2 #欄
	total = 0
	good = 0
	answer = ['PASS','ABORT','FAIL']
	if sampling == 0 : sampling = 999999
	while (str.strip(xlsSheet.Cells(row,col).Value) in answer) and (total < sampling):
		if str.strip(xlsSheet.Cells(row,col).Value) == 'PASS' :  
			good += 1
		row += 1
		total += 1
		if type(xlsSheet.Cells(row,col).Value) == type(None):
			break  # 偵測下一個若是空白則結束
	xlsBook.Close()
	xlsApp.Quit()
	del xlsApp
	return (lot_id,total,good)

yield_data = []
curr_dir = "d:\\kadela\\git\\ATE_SPCC\\sample_2\\"
for Path, Folder, FileName in os.walk(curr_dir):
	for i in FileName:
		tmp = Path + "\\" +i
		if i[:2] == "YS" and i[-3:] == "csv" :
			print (Path,i)
			t = get_yield(filepath = Path, filename = i, sampling = 0)
			yield_data += [t[0],t[1],t[2]]

# 放入管制圖範本
r_qty = len(yield_data) / 3
#print (type(yield_data))
#print (yield_data)
#print ("lot =", r_qty)
total_row = 6   #總數
sample_row = 7  #不良數
col = 3 #欄
chartfile = "d:\\kadela\\git\\ATE_SPCC\\sample\\XR.xlsx"
xlsApp = Dispatch("Excel.Application")
xlsApp.Visible = 0 
xlsBook = xlsApp.Workbooks.open(chartfile)
xlsSheet = xlsBook.Worksheets('p-chart')
i = 1
while i <= r_qty :
	xlsSheet.Cells(total_row,col).Value = yield_data[3*i-2]
	xlsSheet.Cells(sample_row,col).Value = yield_data[3*i-1]
	col += 1
	i += 1
xlsBook.Save()
xlsBook.Close()
xlsApp.Quit()
del xlsApp
