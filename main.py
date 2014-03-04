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
    '7':[u'国网湖南省电力公司',u'湖南省电力公司'],'9':[u'交流220kV',u'交流110kV',u'交流35kV'],'11':u'三相','12':u'ABC相',
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
        cellValue = self.ws.Cells(row,col).Value.strip()
        return cellValue

    
    def pms_validate(self):
        
        err_cells = []
        total_error_lines = 0
        #skip header
        fd = open('validate.log','w')
        for row in range(2,self.nrows + 1):
            fd.write("validating %d/%d" % (row,self.nrows))
            #print "validating %d/%d" % (row,self.nrows)
            this_line_is_error = False
            #1. 设备名称  "#1主变、#2主变、#3主变 以外为不合格数据
            cell11=self._getcell(row,1)
            if cell11 not in pms['1']:
                err_cells.append((row,1))
                this_line_is_error = True
            #2. 运行编号
            ##1、#1主变压器、#2主变压器、#3主变压器以外为不合格数据；
            ##2、与设备名称不一致为不合格数据"
            cell12 = self._getcell(row,2)
            if cell12 not in pms['2']:
                err_cells.append((row,2))
                this_line_is_error = True
                
            elif cell12 != cell11[:4] + u'压器':
                err_cells.append((row,2))
                this_line_is_error = True
                
            #6. 资产性质  "省（直辖市、自治区）公司" 以外为不合格数据；
            if self._getcell(row,6)!= pms['6']:
                err_cells.append((row,6))
                this_line_is_error = True
        
            #7. 资产单位  "国网湖南省电力公司、湖南省电力公司" 以外为不合格数据
            if self._getcell(row,7) not in  pms['7']:
                err_cells.append((row,7))
                this_line_is_error = True
       
            #9. 电压等级   "交流220kV、交流110kV、交流35kV" 以外为不合格数据
            dydj = self._getcell(row,9)
            if dydj not in  pms['9']:
                err_cells.append((row,9))
                this_line_is_error = True
                
            #11. 相数 三相以外为不合格数据
            if self._getcell(row,11) != pms['11']:
                err_cells.append((row,11))
                this_line_is_error = True
            #12. 相别  ABC相以外为不合格数据
            if self._getcell(row,12) != pms['12']:
                err_cells.append((row,12))
                this_line_is_error = True
            
            #13. 额定电压
            ## 1、“电压等级”为交流220kV对应有242、230、220，以外为不合格数据；
            ## 2、“电压等级”为交流110kV对应有121、110，以外为不合格数据；
            ## 3、“电压等级”为交流35kV对应有38.5、35，以外为不合格数据；
            eddy_dict = {u'交流220kV':['242','230','220'],u'交流110kV':['121','110'],u'交流35kV':['38.5','35']}
            eddy = self._getcell(row,13)
            if (dydj  in  eddy_dict.keys() and eddy not in eddy_dict[dydj]):
                err_cells.append((row,13))
                this_line_is_error = True
            
            if this_line_is_error == True:
                total_error_lines += 1    
        #highligh error cells
        if total_error_lines > 0:
            for cell in err_cells:
                self.ws.Cells(cell[0],cell[1]).Interior.ColorIndex = ERROR_COLOR
            self.wb.Save()
        self.excel.Application.Quit()
        fd.close()
        return (total_error_lines,self.nrows)
    
class ExcelHandler():

    def __init__(self,geometry='800x600'):
        self.win = Tk()
        self.win.title('PMS data validation')
        self.win.geometry(geometry)
        self.info = StringVar()
        self.lb = Label(self.win,textvariable=self.info)
        self.lb.pack()
        self.info.set(' ')

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
            pmshandler = PmsValidate(fn)
            (total_error_lines,total_lines) = pmshandler.pms_validate()
            self.info.set("validate %d lines and error lines %d" %(total_lines,total_error_lines))     
        else:
            print("%s open failed" % fn)

    def about(self):
        win_about = Toplevel(self.win)
        lb_about = Label(win_about,text="this is a excel handler program")
        lb_about.pack()

    def help(self):
        print("usage:")


if __name__ == '__main__':
    handler= ExcelHandler('600x400')
    #test code
    #curdir=os.getcwd()
    #pmshandler = PmsValidate(os.path.join(curdir,u'主变压器.xls'))
    #pmshandler.pms_validate()
