#os.path.dirname('d:\\library\\book.txt') 去掉文件名,返回路徑 'd:\\library'
#os.path.basename('d:\\library\\book.txt') 返回文件名 'book.txt'
#os.path.split('d:\\library\\book.txt') 返回路徑和文件名 ('d:\\library', 'book.txt')
#os.path.splitdrive('d:\\library\\book.txt') 返回磁碟機代號和路徑字元組 ('d:', '\\library\\book.txt')
#os.path.splitext('d:\\library\\book.txt') 返回文件名和副檔名 ('d:\\library\\book', '.txt')
#os.listdir() 返回目錄下所有文件和目錄名稱
#os.getcwd() 目前工作目錄
#os.path.isfile() os.path.isdir() 檢驗是檔案還是目錄
#os.path.exists() 檢驗路徑是否存在(不含檔案)
#os.path.splitext() 分離檔案名稱與副檔名
#array_name = [[0 for i in range(m)] for j in range(n)] 建立m*n的二維陣列

import os
from win32com.client import Dispatch, constants

def get_yield( filepath = "d:\\kadela\\git\\ATE_SPCC\\sample\\", filename = "YS11-17110027.csv",sampling = 0):
	#filepath = "d:\\kadela\\git\\ATE_SPCC\\sample\\" #區域網路要先設網路磁碟機代號
	#filename = "YS11-17110027.csv"
	xlsfile = os.path.join(filepath,filename)
	xlsApp = Dispatch("Excel.Application")
	xlsApp.Visible = 0                  #顯示 Excel
	xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
	lot_id = os.path.splitext(filename)[0]      #分離前檔名及副檔名
	#print ("sheetname :", lot_id)
	xlsSheet = xlsBook.Worksheets(lot_id)  

	row = 16   #列
	col = 2 #欄
	total = 0
	good = 0
	#result = []
	answer = ['PASS','ABORT','FAIL']
	if sampling == 0 : sampling = 999999
	while (str.strip(xlsSheet.Cells(row,col).Value) in answer) and (total < sampling): #xls表格內有多餘的空格
		if str.strip(xlsSheet.Cells(row,col).Value) == 'PASS' :  
			good += 1
		row += 1
		total += 1
		if type(xlsSheet.Cells(row,col).Value) == type(None):
			break  # 偵測下一個若是空白則結束
	xlsBook.Close()
	xlsApp.Quit()                #結束 Excel
	del xlsApp
	return (lot_id,total,good)

yield_data = []
curr_dir = "d:\\kadela\\git\\ATE_SPCC\\sample_2\\"
for Path, Folder, FileName in os.walk(curr_dir):
	for i in FileName:
		tmp = Path + "\\" +i
		if i[:2] == "YS" and i[-3:] == "csv" :
			#print (Path,i)
			t = get_yield(filepath = Path, filename = i, sampling = 10)
			yield_data += [t[0],t[1],t[2]]
print (yield_data)
