# -*- coding: utf-8 -*-
import os,time
import re
import pyodbc
from datetime import datetime, timezone
from win32com.client import Dispatch, constants
import shutil #檔案處理套件

if __name__=='__main__':
	dt = str(datetime.now().strftime("%Y%m%d%H%M%S"))
	conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=192.168.1.4;DATABASE=ate;UID=sa;PWD=yds6f')
	c = conn.cursor()

	xlsApp = Dispatch("Excel.Application")
	xlsApp.Visible = 0 
	xlsApp.DisplayAlerts = False 
	spcc_dir = "d:\\temp\\ate\\SPCC\\"
	if not os.path.isdir(spcc_dir):
		print ('管制圖目錄[SPCC]目錄不存在 !!')
		os._exit()
	xlsfile = os.path.join(spcc_dir,'template-XR.xlsx')

	#取測項單位及名稱
	unit = []
	qry = 'SELECT col,unit,name FROM TestUnit ORDER BY col'
	c.execute(qry)
	row = c.fetchone()
	while row:
		unit.append(row)
		row = c.fetchone()
#	for tmp_unit in unit:
	print (unit)

#unit[0] : ['1', 'ma', '空載輸入電流_L']
#unit[2] : ['3', 'ma', '空載輸入電流_H']
#unit[6] : ['7', 'V', '空載輸出電壓CH1_L']

	#取device list
	device = []
	qry = 'SELECT DISTINCT device FROM LotTitle ORDER BY device'
	c.execute(qry)
	row = c.fetchone()
	while row:
		device.append(row)
		row = c.fetchone()

	for tmp_device in device:
		print (tmp_device[0])
		
		#取產品最後一個檔案的上下限
		t_std = []
		qry = "SELECT TOP(1) * FROM LotTitle WHERE device = '" + tmp_device[0] + "' ORDER BY lotdt_idx DESC"
		c.execute(qry)
		row = c.fetchone()
		for tmp_std in row:
			t_std.append(tmp_std)
		print (t_std)
		
		test_item = [0,2]
		for tmp_item in test_item:
			title = []
			#title_item = unit[test_item][2]
			title.append(unit[tmp_item][2])
			#title_unit = unit[test_item][1].upper()
			title.append(unit[tmp_item][1])
			usl = t_std[tmp_item * 2 + 4]
			title.append(usl)
			lsl = t_std[tmp_item * 2 + 5]
			title.append(lsl)
			cl = (usl + lsl) / 2
			title.append(cl)
			tester = t_std[3]
			title.append(t_std[3])
			t_date = str(t_std[0])[:4] + '/' + str(t_std[0])[4:6] + '/' + str(t_std[0])[6:8]
			title.append(t_date)
			print (title)
			#print ('title_item:%s unit:%s usl=%s lsl=%s cl=%s tester:%s date:%s' % (title[0],title[1],title[2],title[3],title[4],title[5],title[6]))
			
			#取得每個產品,測試時間最後25批的lotdt_idx
			lotidx = []
			qry = "SELECT TOP(25) lotdt_idx FROM LotTitle WHERE device = '" + tmp_device[0] + "' ORDER BY lotdt_idx DESC"
			c.execute(qry)
			row = c.fetchone()
			while row:
				lotidx.append(row)
				row = c.fetchone()
			print ('%s 共有 %s 批' % (tmp_device[0], len(lotidx)))

			#取每個lotidx的前5筆tmp_item測項的值
			test_item = "col_" + str(tmp_item)
			test_data = []
			for tmp_idx in lotidx:
				qry = "SELECT TOP(5) col_" + str(tmp_item + 1) + " FROM TestData WHERE lotdt_idx = '" + str(tmp_idx[0]) + "' and t_result = 1"
				#print (qry)
				c.execute(qry)
				row_test_data = c.fetchone()
				data_account = 1
				while row_test_data:
					test_data.append(row_test_data)
					row_test_data = c.fetchone()
					data_account += 1
				while data_account <= 5: #少於5筆自動補0
					test_data.append(0)
					data_account += 1
			print ('%s 共有 %s 筆' % (tmp_device[0], len(test_data)))

			xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
			xlsSheet = xlsBook.Worksheets('X-R')

			xlsSheet.Cells(3,3).Value = str(tmp_device[0]) #device
			xlsSheet.Cells(5,3).Value = title[0]   #_test item
			xlsSheet.Cells(6,3).Value = title[1]   #_unit
			xlsSheet.Cells(4,11).Value = title[2]  #usl
			xlsSheet.Cells(5,11).Value = title[4]  #cl
			xlsSheet.Cells(6,11).Value = title[3]  #lsl
			xlsSheet.Cells(5,23).Value = title[5]  #tester
			xlsSheet.Cells(6,29).Value = title[6]  #date

			pos_x = range(3, 3 + int(len(test_data) / 5))
			pos_y = [9,10,11,12,13]
			test_data_idx = 0
			for x in pos_x:
				for y in pos_y:
					xlsSheet.Cells(y,x).Value = test_data[test_data_idx]
					test_data_idx += 1

			dt_xls = title[6].replace("/","")
			xls_saveas = 'XR-' + str(tmp_device[0]) + '-' + title[0] + '-' + dt_xls + '.xlsx'
			xls_saveas = os.path.join(spcc_dir,xls_saveas)
			#print (xls_saveas)




			xlsSheet.SaveAs(xls_saveas)
			xlsBook.Close()

	xlsApp.Quit()
	del xlsApp
			








