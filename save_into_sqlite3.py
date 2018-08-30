# -*- coding: utf-8 -*-
#對指定目錄內所有測試檔進行統計, 紀錄[檔案修改時間], [lot id], [總數], [良品數]
#再寫入Django的資料庫,以及log紀錄

import os,time
from datetime import datetime, timezone
from win32com.client import Dispatch, constants
import sqlite3
import csv

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
#Django資料庫路徑
conn = sqlite3.connect('d:\\kadela\\git\\mysite\\db.sqlite3')
#測試資料存檔路徑
curr_dir = "d:\\kadela\\git\\ATE_SPCC\\sample_2\\"
c = conn.cursor()

#設定log的路徑與檔名
dt = str(datetime.now().strftime("%Y%m%d%H%M%S"))
log_file = "d:\\kadela\\git\\ATE_SPCC\\log\\atelog-" + dt + ".csv"
f = open(log_file, 'w', newline='')
w = csv.writer(f)

for Path, Folder, FileName in os.walk(curr_dir):
	for i in FileName:
		tmp = Path + "\\" +i
		if i[:2] == "YS" and i[-3:] == "csv" :
			m_time = os.path.getmtime(os.path.join(Path,i))
			f_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(m_time))
			print (Path,i,f_time)
			t = get_yield(filepath = Path, filename = i, sampling = 0)
			yie = int((t[2] / t[1]) * 100)
			sql = "INSERT INTO ate_spcc_test_static (create_date,lotid,total,good,yield_g) \
			VALUES ('" + f_time + "', '"+ t[0][5:] + "', '" + str(t[1]) + "', '" + str(t[2]) + "', '"  + str(yie) + "')"
			print (sql)
			c.execute(sql)
			w.writerow([sql]) #以list放入,每個字元不會被逗號分開

c.close()
conn.commit()
conn.close()
f.close()
