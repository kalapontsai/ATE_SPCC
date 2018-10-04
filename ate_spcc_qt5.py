# -*- coding: utf-8 -*-
import os, sys, time
from datetime import datetime, timezone
import pyodbc
from win32com.client import Dispatch, constants
import shutil #檔案處理套件
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QDesktopWidget, QAction, qApp,\
 QLabel, QHBoxLayout, QLineEdit, QPushButton, QFormLayout, QDialog, QFileDialog, QComboBox, QMessageBox
from PyQt5.QtGui import QIcon, QFont

class Export():
	def __init__(self):
		super(self.__class__, self).__init__()
		self.cwd = os.getcwd()
	
	def output(Path,tmp_xls,device,testitem):
		if len(Path) > 1 :
			if Path[-1:] != '\\' :
				savepath = Path + '\\'
			else :
				savepath = Path
		else :
			print ('參數未設路徑 !! 預定為' + self.cwd)
			savepath = self.cwd

		dt = str(datetime.now().strftime("%Y%m%d%H%M%S"))
		conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=192.168.1.4;DATABASE=ate;UID=sa;PWD=yds6f')
		c = conn.cursor()

		xlsApp = Dispatch("Excel.Application")
		xlsApp.Visible = 0 
		xlsApp.DisplayAlerts = False 

		if not (os.path.isdir(Path) or os.path.isdir(tmp_xls)):
			QMessageBox.information(self,"錯誤","參數非正確路徑 !!", QMessageBox.Ok ,QMessageBox.Ok)
			return ('path error')

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
		msg = 'Device name : ' + device + 'is processing ....'
		MainWindow.statusBar().showMessage(msg)
		print ('Device name : %s is processing ....' % device)
		
		#取產品最後一個檔案的上下限
		t_std = []
		qry = "SELECT TOP(1) * FROM LotTitle WHERE device = '" + device + "' ORDER BY lotdt_idx DESC"
		print (qry)
		c.execute(qry)
		row = c.fetchone()
		for tmp_std in row:
			t_std.append(tmp_std)
		title = []
		title.append(unit[testitem][2])
		title.append(unit[testitem][1])
		usl = t_std[testitem * 2 + 4]
		title.append(usl)
		lsl = t_std[testitem * 2 + 5]
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
		msg = '單位:' + title[1]
		msg += ',usl=' + str(title[2])
		msg += ',lsl=' + str(title[3])
		msg += ',製作日期:' + str(title[6])
		MainWindow.statusBar().showMessage(msg)

		#取得每個產品,測試時間最後25批的lotdt_idx
		lotidx = []
		qry = "SELECT TOP(25) lotdt_idx FROM LotTitle WHERE device = '" + device + "' ORDER BY lotdt_idx DESC"
		c.execute(qry)
		row = c.fetchone()
		while row:
			lotidx.append(row)
			row = c.fetchone()

		#取每個lotidx的前5筆tmp_item測項的值
		test_data = []
		for tmp_idx in lotidx:
			qry = "SELECT TOP(5) col_" + str(testitem + 1) + " FROM TestData WHERE lotdt_idx = '" + str(tmp_idx[0]) + "' and t_result = 1"
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
		print ('%s 共有 %s 批 %s 筆' % (device, len(lotidx), len(test_data)))
		MainWindow.statusBar().showMessage('%s 共有 %s 批 %s 筆' % (device, len(lotidx), len(test_data)))
		try:
			xlsBook = xlsApp.Workbooks.open(tmp_xls)    #開啟一工作簿
		except:
			print ('範本讀取發生異常 !!')
			return ('template error')
		xlsSheet = xlsBook.Worksheets('X-R')

		xlsSheet.Cells(3,3).Value = str(device) #device
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

		#dt_xls = title[6].replace("/","") #改為報告產生日期
		xls_saveas = str(device) + '-' + '平均全距圖'+ '-' + title[0] + '-' + dt + '.xlsx'
		xls_saveas = os.path.join(Path,xls_saveas)
		print ('儲存 %s .....' % (xls_saveas))
		MainWindow.statusBar().showMessage('儲存 %s .....' % (xls_saveas))
		xlsSheet.SaveAs(xls_saveas)
		xlsBook.Close()

		xlsApp.Quit()
		del xlsApp
		conn.close()
		return ('output ok')

class MainWindow(QMainWindow):
	def __init__(self):
		super(self.__class__, self).__init__()
		self.setWindowTitle("報表產出")
		self.statusBar().showMessage('Reday')
		self.cwd = os.getcwd()
		menubar = self.menuBar()
		menubar.setNativeMenuBar(False)
		fileMenu = menubar.addMenu('File')
		#给menu创建一个Action
		exitAction = QAction(QIcon('exit.png'), 'Exit', self)
		exitAction.setShortcut('Ctr+Q')
		exitAction.setStatusTip('Exit Application')
		exitAction.triggered.connect(qApp.quit)
		#将这个Action添加到fileMenu上
		fileMenu.addAction(exitAction)

		self.setupUi()
		self.setFixedSize(640, 250)
		#self.resize(640, 250) #若要可變視窗大小
		self.center()
		self.show()
	
	def click_changepath(self,Path):
		self.statusBar().showMessage('Reday')
		if os.path.exists(Path):
			path = QFileDialog.getExistingDirectory(self,"選取資料夾",Path)
		else:
			path = QFileDialog.getExistingDirectory(self,"選取資料夾",self.cwd)
			#path = QFileDialog.getOpenFileName(self,"Open File Dialog","/","Excel files(*.xlsx)")
		if path != '' :
			self.lineedit_path.setText(str(path))
		print(path)

	def click_change_temp_path(self,Path):
		self.statusBar().showMessage('Reday')
		if os.path.exists(Path):
			path = QFileDialog.getOpenFileName(self,"選取範本檔",Path,"Excel files(*.xlsx)")
		else:
			path = QFileDialog.getOpenFileName(self,"選取範本檔",self.cwd,"Excel files(*.xlsx)")
		if path[0] != '' :
			self.lineedit_temp_path.setText(str(path[0]))
		print(path)


	def click_go(self):
		if len(self.lineedit_path.text()) > 1 :
			if not (os.path.isdir(self.lineedit_path.text())):
				print ('參數非正確路徑 !!')
				QMessageBox.information(self,"錯誤","參數非正確路徑 !!", QMessageBox.Ok ,QMessageBox.Ok)
				return
			else :
				msg = '存檔路徑:'+self.lineedit_path.text() + '\n'
				msg += '範本路徑:' + self.lineedit_temp_path.text() + '\n'
				msg += '機種名稱:' + self.ComboBox_device.currentText() + '\n'
				msg += '測項名稱:' + self.ComboBox_item.currentText() + '\n\n'
				msg += '是否繼續?'
				reply = QMessageBox.question(self,'確認',msg, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

			if reply == QMessageBox.Yes:
				print ('click_go : Yes')
				self.statusBar().showMessage('開始製作報表........')
				result = Export.output(self.lineedit_path.text(),self.lineedit_temp_path.text(),self.ComboBox_device.currentText(),self.ComboBox_item.currentIndex())
				if result == 'output ok' :
					self.statusBar().showMessage('報表產出完成 !!')
				elif result == 'path error':
					self.statusBar().showMessage('存檔路徑發生問題, 轉換中止 !!')
				elif result == 'template error':
					self.statusBar().showMessage('範本讀取發生問題, 轉換中止 !!')
			else:
				print ('click_go : No')
				self.statusBar().showMessage('Ready')

	def center(self):
		qr = self.frameGeometry()
		cp = QDesktopWidget().availableGeometry().center()
		qr.moveCenter(cp)
		self.move(qr.topLeft())

	def setupUi(self):
		central_widget = QWidget()
		label_path = QLabel()
		label_path.setText("存檔路徑")
		label_path.setFont(QFont('SansSerif', 12))
		label_temp_path = QLabel()
		label_temp_path.setText("範本路徑")
		label_temp_path.setFont(QFont('SansSerif', 12))
		label_device = QLabel()
		label_device.setText("機種:")
		label_device.setFont(QFont('SansSerif', 14))
		label_item = QLabel()
		label_item.setText("測項:")
		label_item.setFont(QFont('SansSerif', 14))
		self.lineedit_path = QLineEdit()
		self.lineedit_path.setFont(QFont('SansSerif', 14))
		self.lineedit_path.setFixedWidth(400)
		self.lineedit_path.setText(self.cwd)
		self.lineedit_temp_path = QLineEdit()
		self.lineedit_temp_path.setFont(QFont('SansSerif', 14))
		self.lineedit_temp_path.setFixedWidth(400)
		self.lineedit_temp_path.setText(os.path.join(self.cwd,'template-XR.xlsx'))
		button_path = QPushButton('變更', self)
		button_path.setFont(QFont('SansSerif', 14))
		button_path.clicked.connect(lambda:self.click_changepath(self.lineedit_path.text()))
		button_temp_path = QPushButton('變更', self)
		button_temp_path.setFont(QFont('SansSerif', 14))
		button_temp_path.clicked.connect(lambda:self.click_change_temp_path(self.lineedit_temp_path.text()))
		self.ComboBox_device = QComboBox(self)
		self.ComboBox_device.setFont(QFont('SansSerif', 12))
		self.ComboBox_item = QComboBox(self)
		self.ComboBox_item.setFont(QFont('SansSerif', 12))
		conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=192.168.1.4;DATABASE=ate;UID=sa;PWD=yds6f')
		c = conn.cursor()
		qry = 'SELECT DISTINCT device FROM LotTitle ORDER BY device'
		c.execute(qry)
		row = c.fetchone()
		#print (row)
		while row:
			self.ComboBox_device.addItem(row[0])
			row = c.fetchone()

		qry = 'SELECT DISTINCT name,col FROM TestUnit ORDER BY col'
		c.execute(qry)
		row = c.fetchone()
		#print (row)
		while row:
			self.ComboBox_item.addItem(row[0],row[1])
			row = c.fetchone()
		conn.close()

		button_go = QPushButton('開始製作', self)
		button_go.setFont(QFont('SansSerif', 14))
		button_go.clicked.connect(self.click_go)

		#另一種button排列方式
		#pybutton = QPushButton('Click me', self)
		#pybutton.clicked.connect(self.clickMethod)
		#pybutton.resize(100,32)
		#pybutton.move(50, 50)
		
		main_layout = QFormLayout()
		main_layout.setVerticalSpacing(20)
		h1_layout = QHBoxLayout()
		h1_layout.addWidget(label_path)
		h1_layout.addSpacing(5)
		h1_layout.addWidget(self.lineedit_path)
		h1_layout.addSpacing(20)
		h1_layout.addWidget(button_path)
		h2_layout = QHBoxLayout()
		h2_layout.addWidget(label_temp_path)
		h2_layout.addSpacing(5)
		h2_layout.addWidget(self.lineedit_temp_path)
		h2_layout.addSpacing(20)
		h2_layout.addWidget(button_temp_path)
		h3_layout = QHBoxLayout()
		h3_layout.addWidget(label_device)
		h3_layout.addSpacing(10)
		h3_layout.addWidget(self.ComboBox_device)
		h3_layout.addSpacing(10)
		h3_layout.addWidget(label_item)
		h3_layout.addSpacing(10)
		h3_layout.addWidget(self.ComboBox_item)
		h3_layout.addSpacing(100)

		h4_layout = QHBoxLayout()
		h4_layout.addWidget(button_go)

		#h_layout.setContentsMargins(0, 0, 0, 0)
		main_layout.addRow(h1_layout)
		main_layout.addRow(h2_layout)
		main_layout.addRow(h3_layout)
		main_layout.addRow(h4_layout)
		central_widget.setLayout(main_layout)
		self.setCentralWidget(central_widget) # new central widget

if __name__ == "__main__":
	app = QApplication(sys.argv)
	MainWindow = MainWindow()
	sys.exit(app.exec_())
