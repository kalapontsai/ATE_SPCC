## ATE_SPCC 2018/10/04 ##
*	catch_csv_report_into_SQL.py :
	> ATE tester will create a result file as csv format
	> use Python.win32 to catch the following data : 
		1.test report title include device name, testing date, lot number,tester machine number
		2.use the last modified datetime as index to identifiy the same lot has many test report
		3.catch USL/LSL value from report and keep it in table[LotTitle]
		4.catch test value by each item and refer the index to [Lottitle].[Lotdx_idx]

	> program also identify the test reslt by 3 keyword : "PASS", "ABORT", "FAIL", the refernce table is [TestResult]
	> save into MS SQL database
	> save log and move the csv file to a specified directory

*	query_to_XR_chart.py :
	> query criteria : collect 10 newest lot by each device. get the test value from 5ea pass sample.
	> use MS excel templated file to fill in all data it need.
	> refer the specified test item (test_item[]) then query the data from [TestData]
	> save as new file to specified directory

*	catch_csv_report_into_sqlite3.py :
	> phase out due to new file:catch_csv_report_into_SQL.py

*	ate_spcc_qt5.py :
	> use PYQT5 to create user UI instead of open .ini file for parameter setting


