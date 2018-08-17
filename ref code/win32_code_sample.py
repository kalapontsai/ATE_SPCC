#!/usr/bin/python
# coding=utf-8
'''
@author: humingwei
'''
import os
import sys
import win32com.client as win32
from threading import current_thread, _MainThread


# 下划线种类: XlUnderlineStyle 枚举 (Word)
# xlUnderlineStyleNone -4142      无下划线。
# xlUnderlineStyleSingle 2  单下划线。
# xlUnderlineStyleDouble -4119  粗双下划线。
# xlUnderlineStyleSingleAccounting 4  不支持。
# xlUnderlineStyleDoubleAccounting 5      彼此靠近的两条细下划线。

# ColorIndex: XlColorIndex 枚举 (Excel)
# xlColorIndexNone    -4142    无颜色。
# xlColorIndexAutomatic    -4105    自动配色。
# 1-黑、2-白、3-红、4-绿、5-青、6-黄、7-紫、8-蓝


class Excel(object):
    def __init__(self, file_path=''):
        '''初始'''
        if not isinstance(current_thread(), _MainThread):
            from pythoncom import CoInitialize
            CoInitialize()  # 为调用线程初始化COM库。
        self.xlapp = win32.gencache.EnsureDispatch('Excel.Application')
        self.xlapp.Visible = False
        self.xlapp.DisplayAlerts = False  # 不警告
        if os.path.isfile(file_path):
            self.workbook = self.xlapp.Workbooks.Open(file_path)
        else:
            self.workbook = self.xlapp.Workbooks.Add()
            self.workbook.SaveAs(file_path)
    
#     def __new__(self, file_path):
#         ''''''
#         return 
#         self.workbook = self.xlapp.Workbooks.Open(file_path)
    
    def save(self, new_file_name=None):
        '''保存'''
        if new_file_name:
            self.workbook.SaveAs(new_file_name)
        else:
            self.workbook.Save()
    
    def get_sheet(self, sheet_name=1):
        '''获取页签'''
        try:
            sheet = self.workbook.Worksheets(sheet_name)
        except:
            if isinstance(sheet_name, str):
                sheet = self.workbook.Worksheets.Add()
                sheet.Name = sheet_name
            else:
                raise Exception('The number of sheet is out of range!\n')
        return Sheet(sheet)
    
    def close(self):
        '''保存并退出'''
        self.save()
        self.xlapp.Application.Quit()
        
    def __del__(self):
        ''''''
        self.close()
        del self.xlapp


class Sheet(object):
    def __init__(self, sheet):
        ''''''
        self.sheet = sheet
    
    def get_cell_rows(self):
        '''获取已使用行数'''
        return self.sheet.UsedRange.Rows.Count
    
    def get_cell_cols(self):
        '''获取已使用列数'''
        return self.sheet.UsedRange.Columns.Count
    
    def get_cell(self, row, col):
        '''获取单元格数据'''
        return self.sheet.Cells(row, col).Value
    
    def set_cell(self, row, col, value, bold=False, under_line=-4142, color_index=-4142):
        '''设置单元格数据'''
        the_cell = self.sheet.Cells(row, col)
        the_cell.Value = value
        the_cell.Font.Bold = bold
        the_cell.Font.Underline = under_line
        the_cell.Font.ColorIndex = color_index
    
    def get_range(self, row1=1, col1=1, row2=None, col2=None):
        '''获取块数据'''
        if not row2:
            row2 = self.get_cell_rows()
        if not col2:
            col2 = self.get_cell_cols()
        return self.sheet.Range(self.sheet.Cells(row1, col1), self.sheet.Cells(row2, col2)).Value
    
    def set_range(self, row1=1, col1=1, value=[], bold=False, under_line=-4142, color_index=-4142):
        '''块赋值'''
        rows = len(value)
        cols = len(value[0])
        the_range = self.sheet.Range(self.sheet.Cells(row1, col1), self.sheet.Cells(row1 + rows - 1, col1 + cols - 1))
        the_range.Value = value
        the_range.Font.Bold = bold
        the_range.Font.Underline = under_line
        the_range.Font.ColorIndex = color_index
        
#         cols = None # 按行写，有利于节约内存，不利于过滤出问题数据
#         for row_value in value:
#             if not cols:
#                 cols = col1 + len(row_value) - 1
#             self.sheet.Range(self.sheet.Cells(row1, col1), self.sheet.Cells(row1, cols)).Value = row_value
#             row1 += 1  
    
    def merge_cells(self, row1, col1, row2, col2, value='', bold=False):
        '''合并单元格'''
        the_range = self.sheet.Range(self.sheet.Cells(row1, col1), self.sheet.Cells(row2, col2))
        the_range.MergeCells = True
        if value:
            the_range.Value = value
        the_range.Font.Bold = bold
    
    def add_hyperlink(self, row, col, address, sub_address='', value='', tip=''):
        '''
        @todo: 添加超链接
        @param address: 超链接的地址
        @param sub_address: 超链接的子地址
        @param value: 要显示的超链接的文本
        @param tip: 要显示的超链接的文本
        '''
        the_cell = self.sheet.Cells(row, col)
        self.sheet.Hyperlinks.Add(Anchor=self.sheet.Range(the_cell, the_cell),
                                  Address=address,
                                  SubAddress=sub_address,
                                  ScreenTip=tip,
                                  TextToDisplay=value)
    
    def add_picture(self, picture_path, left, top, width, height, link_file=False, save_with_document=True):
        '''
        @todo: 插入图片
        @param picture_path: 要创建的 OLE 对象的源文件
        @param left: 相对于文档的左上角，以磅为单位给出图片左上角的位置
        @param top: 相对于文档的顶部，以磅为单位给出图片左上角的位置
        @param width: 以磅为单位给出图片的宽度
        @param height: 以磅为单位给出图片的高度
        @param link_file: 要链接至的文件
        @param save_with_document: 将图片与文档一起保存
        '''
        self.sheet.Shapes.AddPicture(picture_path, link_file, save_with_document, left, top, width, height)
    

if __name__ == '__main__':
    file_name = 'test.xlsx'
    file_path = os.path.join(os.path.split(os.path.realpath(__file__))[0], file_name)
    excel = Excel(file_path)
    sheet = excel.get_sheet('hmw')
    for i in range(1, 20):
        sheet.set_cell(i, i, i, color_index=i)
    print(sheet.get_cell_rows())
    print(sheet.get_cell_cols())
    print(sheet.get_cell(1, 1))
#     sheet.set_range(1, 1, [[1, 2, 3, 4], [21, 22, 23, 24], ['31', '32', '33', '34'],
#                            ['d1', 'd2', 'd3', 'd4']], True, 2, 5)
    print(sheet.get_cell_rows())
    print(sheet.get_cell_cols())
    print(sheet.get_range())
    sheet.merge_cells(1, 1, 2, 2, 'merge_cells')
#     sheet.add_hyperlink(5, 5, '鱼欲渔')
#     sheet.add_hyperlink(6, 6, file_name, "'Sheet1'!A1")
#     sheet.add_picture(r'E:\Img\000001.png', 7, 1, 160, 40)
    del excel
    print('__END__')
    sys.exit()