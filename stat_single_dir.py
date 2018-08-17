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

import os.path
import numpy as np
import stat_single_file as sf

listdir = []
listfile = []

curr_dir = "d:\\kadela\\git\\ATE_SPCC\\sample\\"
tmp_list = os.listdir(curr_dir)
for i in tmp_list:
	#tmp_i = os.path.join(curr_dir,i)
	if os.path.isdir(os.path.join(curr_dir,i)):
		listdir.append(i)
	else:
		listfile.append(i)
print (listdir,listfile)

yield_data = []
for i in listfile:
	t = sf.get_yield(filepath = curr_dir, filename = i)
	print ('---------')
	print ('lotid:',t[0])
	print ('total:',t[1])
	print ('good :',t[2])
	yield_data += [t[0],t[1],t[2]]

print('------')
print(yield_data)


