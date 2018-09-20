# -*- coding: utf-8 -*-
# 啟動指令 : query_to_XR_chart.py 輸出檔案的目錄
# ex: query_to_XR_chart.py D:\spcc\
import os,sys,time
import pyodbc
import csv
from datetime import datetime, timezone
from win32com.client import Dispatch, constants
import shutil #檔案處理套件

if __name__=='__main__':
	#讀取存檔路徑
	if len(sys.argv) > 1 :
		if not (os.path.isdir(sys.argv[1])):
			print ('參數非正確路徑 !! 預定為 D:\\temp\\ate\\spcc\\')
			spcc_dir = "d:\\temp\\ate\\spcc\\"
		else :
			if sys.argv[1][:-1] != '\\' :
				spcc_dir = sys.argv[1] + '\\'
			else :
				spcc_dir = sys.argv[1]
	else :
		print ('參數未設路徑 !! 預定為 D:\\temp\\ate\\spcc\\')
		spcc_dir = "d:\\temp\\ate\\spcc\\"
	spcc_config = spcc_dir + "config\\"

	xlsfile = os.path.join(spcc_config,'template-XR.xlsx')
	configfile = os.path.join(spcc_config,'spcc_config.ini')

	#讀取參數檔
	csvfile = open(configfile, 'r')
	csvCursor = csv.reader(csvfile, delimiter=':')
	device = []
	test_item = []
	for row in csvCursor:
		if any(row):  #過濾空行 避免錯誤訊息 IndexError: list index out of range
			if row[0] == 'server' : server = row[1]
			if row[0] == 'database' : database = row[1]
			if row[0] == 'uid' : uid = row[1]
			if row[0] == 'pwd' : pwd = row[1]
			if row[0] == 'device' :device.append(row[1])
			if row[0] == 'test_item' :
				item = str.split(row[1],',')
				#print ('test_item: %s' % item)
				for i in item:
					test_item.append(int(i))
	if test_item == []:
		print ('未指定測項,預設為[空載輸入電流_L]')
		test_item.append(0)
	csvfile.close()

	dt = str(datetime.now().strftime("%Y%m%d%H%M%S"))
	if server and database and uid and pwd :
		odbc = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database
		odbc += ';UID=' + uid + ';PWD=' + pwd
		conn = pyodbc.connect(odbc)
		c = conn.cursor()
	else:
		print ('ODBC connect error !!')
		os._exit()

	xlsApp = Dispatch("Excel.Application")
	xlsApp.Visible = 0 
	xlsApp.DisplayAlerts = False 

	if not (os.path.isdir(spcc_dir) or os.path.isdir(spcc_config)):
		print ('管制圖目錄[SPCC] 或 [SPCC/config]錯誤 !!')
		os._exit()

	#取測項單位及名稱
	unit = []
	qry = 'SELECT col,unit,name FROM TestUnit ORDER BY col'
	c.execute(qry)
	row = c.fetchone()
	while row:
		unit.append(row)
		row = c.fetchone()
	#unit[0] : ['1', 'ma', '空載輸入電流_L']
	#unit[2] : ['3', 'ma', '空載輸入電流_H']
	#unit[6] : ['7', 'V', '空載輸出電壓CH1_L']

	#取device list
	if device == [] :
		qry = 'SELECT DISTINCT device FROM LotTitle ORDER BY device'
		c.execute(qry)
		row = c.fetchone()
		while row:
			device.append(row)
			row = c.fetchone()

	for tmp_device in device:
		print ('Device name : %s is processing ....' % tmp_device)
		
		#取產品最後一個檔案的上下限
		t_std = []
		qry = "SELECT TOP(1) * FROM LotTitle WHERE device = '" + tmp_device[0] + "' ORDER BY lotdt_idx DESC"
		#print (qry)
		c.execute(qry)
		row = c.fetchone()
		for tmp_std in row:
			t_std.append(tmp_std)

		for tmp_item in test_item:
			#print ('tmp_item : ' + str(tmp_item))
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
			print ('==========')
			print ('測項: %s' % (title[0]))
			print ('單位: %s usl=%s lsl=%s cl=%s 機台號碼:%s 日期:%s' % (title[1],title[2],title[3],title[4],title[5],title[6]))
			
			#取得每個產品,測試時間最後25批的lotdt_idx
			lotidx = []
			qry = "SELECT TOP(25) lotdt_idx FROM LotTitle WHERE device = '" + tmp_device[0] + "' ORDER BY lotdt_idx DESC"
			c.execute(qry)
			row = c.fetchone()
			while row:
				lotidx.append(row)
				row = c.fetchone()

			#取每個lotidx的前5筆tmp_item測項的值
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
			print ('%s 共有 %s 批 %s 筆' % (tmp_device, len(lotidx), len(test_data)))

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
			xlsSheet.SaveAs(xls_saveas)
			xlsBook.Close()
			print ('儲存 %s .....' % (xls_saveas))
			print ('- - - - - - - - - - - - - - - - - - -')

	xlsApp.Quit()
	del xlsApp
			








