# -*- coding: utf-8 -*-
#對指定目錄內所有測試檔進行統計
#先抓取報告內表單屬性[檔案修改時間], [lot name], [測試機台號碼], [各測項上下限]
#後以批次讀取各測項讀值,共75項
#再寫入MS-SQL的資料庫,以及log紀錄
#處理過的檔案進行歸檔,原目錄清空

import os,time
import re
import pyodbc
from datetime import datetime, timezone
from win32com.client import Dispatch, constants
import csv
import shutil #檔案處理套件

def title_filter (string):
	if re.findall(r"\-\d+\.?\d*",string) == [] :
		t = re.findall(r"\d+\.?\d*",string)
	else :
		t = re.findall(r"\-\d+\.?\d*",string)

	if t == []:
		t = "0.0"
	return t

def get_title(filepath, filename):
	xlsfile = os.path.join(filepath,filename)
	xlsApp = Dispatch("Excel.Application")
	xlsApp.Visible = 0                  #顯示 Excel
	xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
	xlsSheet = xlsBook.Worksheets(os.path.splitext(filename)[0])#分離前檔名及副檔名
	#找出各欄位的值
	lotname = str.strip(xlsSheet.Cells(5,1).Value)[4:]
	device = str.strip(xlsSheet.Cells(2,1).Value)[41:]
	#test_date = str.strip(xlsSheet.Cells(12,1).Value)[5:] #實際儲存值抓檔案最後時間 變數:f_time
	tester = '1'
	#print ('lotname: %s / device: %s / tester: %s' % (lotname,device,tester))
	#input ('抓取上下限')
	col_low = []
	col_high = []
	pos_x = 3
	pos_y = 14
	while pos_x <= 77:
		s_low = title_filter(str.strip(xlsSheet.Cells(pos_y,pos_x).Value))[0]
		s_high = title_filter(str.strip(xlsSheet.Cells(pos_y+1,pos_x).Value))[0]
		col_low.append(s_low)
		col_high.append(s_high)
		pos_x += 1
	xlsBook.Close()
	xlsApp.Quit()
	del xlsApp
	return (lotname,device,tester,col_low,col_high)


def get_yield( filepath, filename,sampling = 999999):
	xlsfile = os.path.join(filepath,filename)
	xlsApp = Dispatch("Excel.Application")
	xlsApp.Visible = 0                  #顯示 Excel
	xlsBook = xlsApp.Workbooks.open(xlsfile)    #開啟一工作簿
	lot_id = os.path.splitext(filename)[0]      #分離前檔名及副檔名
	xlsSheet = xlsBook.Worksheets(lot_id)  
	row = 16   #列
	col = 2 #欄
	#y_total 內 y_data為單筆list
	y_total = []
	idx = 0 #筆數
	answer = ['PASS','ABORT','FAIL']
	while (str.strip(xlsSheet.Cells(row,col).Value) in answer) and (idx < sampling):
		y_data = []
		if str.strip(xlsSheet.Cells(row,col).Value) == 'PASS' :  
			y_data.append('1')
		elif str.strip(xlsSheet.Cells(row,col).Value) == 'ABORT' :
			y_data.append('2')
		else :
			y_data.append('3')
		pos_x = 3
		while pos_x <= 77 :
			y_data.append(xlsSheet.Cells(row,pos_x).Value)
			pos_x += 1
		#print (y_data)
		y_total.append(y_data)
		row += 1
		idx += 1
		if type(xlsSheet.Cells(row,3).Value) == type(None):
			break  # 偵測下一個若是空白則結束
	xlsBook.Close()
	xlsApp.Quit()
	del xlsApp
	#print (y_total)
	return (y_total)

def check_record(lotdt):
	qry = 'SELECT TOP (1) lotdt_idx FROM LotTitle WHERE lotdt_idx = ' + lotdt
	c.execute(qry)
	if c.fetchone() :
		return True
	else:
		return False



if __name__=='__main__':
	dt = str(datetime.now().strftime("%Y%m%d%H%M%S"))
	#連接資料庫
	conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=192.168.1.4;DATABASE=ate;UID=sa;PWD=yds6f')
	c = conn.cursor()
	sql_pre = 'INSERT INTO TestData (lotdt_idx,t_result,'
	for tmp_fieldname in range(1,76):
		sql_pre += 'col_' + str(tmp_fieldname) + ','
	sql_pre = sql_pre[:-1] + ') VALUES '

	#測試資料原始檔根目錄
	curr_dir = "d:\\temp\\ate\\Dc-DcTestDataRecode\\"
	if not os.path.isdir(curr_dir):
		print ('測試資料目錄不存在 !!')
		os._exit()

	#測試資料歸檔根目錄
	save_dir = "d:\\temp\\ate\\Dc-DcTestDataRecode-Save\\"
	save_dir += dt
	if not os.path.isdir(save_dir): #檢查儲存目錄是否存在
		print ('測試資料存檔目錄不存在, 正在新建....')
		os.mkdir(save_dir)

	log_file = save_dir + "\\atelog-" + dt + ".csv"
	f = open(log_file, 'w', newline='')
	w = csv.writer(f)

	for Path, Folder, FileName in os.walk(curr_dir):
		w.writerow(['= = = = = = = = = = = = ='])
		for i in FileName:
			m_time = os.path.getmtime(os.path.join(Path,i))
			f_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(m_time))
			lot_dt = str(time.strftime("%Y%m%d%H%M%S", time.localtime(m_time)))
			if i[:2] != "YS" or i[-3:] != "csv" :
				w.writerow([('%s%s ----------檔名不符,不予儲存: ' % (Path,i))])
				print ('=================================')
				print ('%s%s ----------檔名不符,不予儲存: ' % (Path,i))
				continue
			if check_record(lot_dt):
				w.writerow([('%s%s ----------檔案重複,不予儲存: ' % (Path,i))])
				print ('=================================')
				print ('%s%s ----------檔案重複,不予儲存: ' % (Path,i))
				continue
			print ('=================================')
			print ('路徑: %s' % (Path))
			print ('檔案: %s' % (i))
			w.writerow(['讀取:%s%s...' % (Path,i)])

			title = get_title(filepath = Path, filename = i)
			sql = 'INSERT INTO LotTitle (lotdt_idx,lotname,device,tester,'
			for tmp_title in range(1,76):
				sql += 'col_' + str(tmp_title) + '_l, col_' + str(tmp_title) + '_h,'
			sql = sql[:-1] + ') VALUES ('
			#input (sql)
			sql += lot_dt + ", '" + title[0] + "', '" + title[1] + "', " + title[2] + ","
			#input (sql)
			for tmpStd in range(len(title[3])):
				sql += title[3][tmpStd] + ',' + title[4][tmpStd] + ','
			sql = sql[:-1] + ')'
			w.writerow([sql])
			c.execute(sql)
			conn.commit()
			
			#抓量測數據
			yield_data = get_yield(filepath = Path, filename = i, sampling = 999999)
			#SQL限制最大筆數1000
			batch = 100
			yield_acc = int(len(yield_data))
			loop_mod = yield_acc % batch
			#判斷是否有餘數需再加一次迴圈
			if loop_mod > 0:
				loop = int(yield_acc / batch) + 1
			else:
				loop = int(yield_acc / batch)
			print ('total %s / loop:%s / 餘數:%s' % (yield_acc,loop,loop_mod))
			#紀錄目前的位置
			pos = 0
			#每批次的終點
			batch_end = 0
			for loop_i in range(1,loop+1) :	
				if (loop_i * batch) > yield_acc :
					batch_end = yield_acc					
				else:
					batch_end = loop_i * batch
				print ('目前迴圈:%s / 位置:%s / 單批終點:%s' % (loop_i,pos,batch_end))
				w.writerow(['----------目前迴圈:%s / 位置:%s / 單批終點:%s' % (loop_i,pos,batch_end)])

				sql = ""
				for tmp_data1 in range(pos, batch_end):
					sql += '(' + lot_dt + ','
					for tmp_data2 in range(len(yield_data[0])):
						sql += str(yield_data[tmp_data1][tmp_data2]) + ','
					sql = sql[:-1] + '),'
					pos += 1
				sql = sql[:-1]
				sql = sql_pre + sql
				w.writerow([sql]) #以list放入,每個字元不會被逗號分開
				c.execute(sql)
				conn.commit()
			newfile = i	
			while os.path.isfile(save_dir + "\\" + newfile): #相同檔名在副檔名加字
				newfile += ".1"
			tmp = os.path.join(Path,i)
			shutil.move(tmp,save_dir + "\\" + newfile) #將處理過的檔案歸檔
	c.close()
	conn.close()
	f.close()
	


