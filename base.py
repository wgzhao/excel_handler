# -*- coding:utf-8 -*-
import xlrd
class PmsBase():
    '''
    电力校验程序基类，其他具体的校验程序均基于该类
    '''
    def __init__(self,filepath):
        '''
        打开给定的Excel文件，并获得当前文件的行数和列数
        '''
        self.wb = xlrd.open_workbook(filepath,formatting_info=True)
        self.ws = self.wb.sheet_by_index(0)
        self.nrows = self.ws.nrows
        self.ncols = self.ws.ncols       
        
    def _getcell(self,row,col):
        '''
        获得指定单元格的内容
        '''
        return  self.ws.cell(row,col).value.strip()

    def _isnumber(self,s):
        #一个字符串是否是数字
        try:
            k = float(s)
            return True
        except ValueError:
            return False
    
    def validate(self):
        '''
        虚函数，需要在具体的子类中实现
        '''
        pass