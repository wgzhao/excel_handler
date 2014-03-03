#!/usr/bin/env python
# -*- coding:utf-8 -*-
__author__ = 'wgzhao<wgzhao@gmail.com>'

from Tkinter import *
import tkFileDialog
import os
import win32com.client as win32
import sys
codec=sys.getfilesystemencoding()


ERROR_COLOR=27
pms={'1':[u'#1主变',u'#2主变',u'#3主变'],'2':[u'#1主变压器',u'#2主变压器',u'#3主变压器'],'6':u'省（直辖市、自治区）公司',
    '7':[u'国网湖南省电力公司',u'湖南省电力公司'],'9':[u'交流220kV',u'交流110kV',u'交流35kV'],'相数':u'三相','相别':u'ABC相',
    '自造国家':u'中国','使用环境':u'户外式','绝缘耐热等级':[u'A',u'B'],'用途':u'降压变压器',
    '结构形式':u'芯式','冷却方式':[u'强迫油循环导向风冷(ODAF)',u'自然油循环风冷(ONAF)',u'自然冷却(ONAN'],'调压方式':[u'有载调压',u'无励磁调压']}

def _pms_validate(filename):
    #Excel表格应该共有56列
    cell_error = easyxf('pattern: pattern solid, back_colour yellow')
    workbook =xlrd.open_workbook(filename)
    pms_sheet = workbook.sheet_by_index(0)
    ncols = pms_sheet.ncols
    nrows = pms_sheet.nrows
    i = 1 #跳过表头
    
    
    return (True,None)
    
class PmsValidate():
    def __init__(self,filename):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.wb = self.excel.Workbooks.Open(filename)
        self.excel.Visible = False
        self.ws = self.wb.Worksheets(1)
        self.ws.Activate()
        self.nrows = self.ws.UsedRange.Rows.Count
        self.ncols = self.ws.UsedRange.Columns.Count
        
    def _getcell(self,row,col):
        cellValue = self.ws.Cells(row,col).Value
        return cellValue

    
    def pms_validate(self):
        err_cells = []
        #skip header
        for row in range(3,self.nrows + 1):
            #1. 设备名称  
            cell11=self._getcell(row,1)
            if cell11 not in pms['1']:
                err_cells.append((row,1))
            #2. 运行编号
            cell12 = self._getcell(row,2)
            if cell12 not in pms['2']:
                err_cells.append((row,2))
            elif cell12 != cell11[:4] + u'压器':
                err_cells.append((row,2))
            
            #6. 资产性质
            if self._getcell(row,6)!= pms['6']:
                err_cells.append((row,6))
        
            #7. 资产单位
            if self._getcell(row,7) not in  pms['7']:
                err_cells.append((row,7))
       
            #9. 电压等级    
            if self._getcell(row,9) not in  pms['9']:
                err_cells.append((row,9))
        #highligh error cells
        if len(err_cells):
            for cell in err_cells:
                self.ws.Cells(cell[0],cell[1]).Interior.ColorIndex = ERROR_COLOR
            self.wb.Save()
        self.excel.Application.Quit()
    
class ExcelHandler():

    def __init__(self,geometry='800x600'):
        self.win = Tk()
        self.win.title('PMS data validation')
        self.win.geometry(geometry)


        self.menubar = Menu(self.win)

        #file menu
        filemenu = Menu(self.menubar,tearoff=0)
        filemenu.add_command(label='Open',command=self.openexcel)
        filemenu.add_separator()
        filemenu.add_command(label='Exit',command=self.win.quit)
        self.menubar.add_cascade(label='File',menu=filemenu)
        #help menu
        helpmenu = Menu(self.menubar,tearoff=0)
        helpmenu.add_command(label="About..",command=self.about)
        helpmenu.add_command(label='Help',command=self.help)
        self.menubar.add_cascade(label='Help',menu=helpmenu)
        self.win.config(menu=self.menubar)
        self.win.mainloop()

    def hello(self):
        print('hello,world!')

    def openexcel(self):
        #import a excel file
        ftypes=[("97-2003 Excel files",'*.xls'),('2007 Excel files','*.xlsx')]
        fn = tkFileDialog.askopenfilename(parent=self.win,title='Choose a file',filetypes=ftypes)

        if fn != '':
            print(fn)
            pms_validate(fn)
        else:
            print("%s open failed" % fn)

    def about(self):
        win_about = Toplevel(self.win)
        lb_about = Label(win_about,text="this is a excel handler program")
        lb_about.pack()

    def help():
        print("usage:")


if __name__ == '__main__':
    #handler= ExcelHandler('600x400')
    #test code
    curdir=os.getcwd()
    pmshandler = PmsValidate(os.path.join(curdir,u'主变压器.xls'))
    pmshandler.pms_validate()

