#!/usr/bin/env python
# -*- coding:gbk -*-
__author__ = 'wgzhao<wgzhao@gmail.com>'

from Tkinter import *
import tkFileDialog
import xlrd 
import xlwt
from xlwt import easyxf
import win32com.client as win32
#湖南省主变PMS台帐参数说明
#key为参数名称，value为参数允许的值，大小写敏感，除此之外，还有以下条件
#1. 运行编号应与设备名称一致，如果设备名称为'#1主变'，则运行编号必须为'#1主变压器'
#2. 间隔单元必须与设备名称一致
#3. 额定电压和电压等级存在映射关系，条件如下：
##    1、“电压等级”为交流220kV对应有242、230、220，以外为不合格数据；
##    2、“电压等级”为交流110kV对应有121、110，以外为不合格数据；
##    3、“电压等级”为交流35kV对应有38.5、35，以外为不合格数据；"
#4. 设备型号和电压等级有映射关系，关系如下：
##         1、“电压等级”为交流220kV数据末尾对应有180000/220、120000/220、240000/220，以外为不合格数据；
##        2、“电压等级”为交流110kV数据末尾对应有20000/110、31500/110、50000/110，以外为不合格数据；
##        3、“电压等级”为交流35kV数据末尾对应有3150/35、4000/35、5000/35、6300/35、10000/35，以外为不合格数据


#########
#列顺序
# 设备名称
# 运行编号
# 所属市局
# 运行单位
# 变电站
# 资产性质
# 资产单位
# 设备类型
# 电压等级
# 间隔单元
# 相数
# 相别
# 额定电压(kV)
# 额定电流(A)
# 额定频率(Hz)
# 设备型号
# 生产厂家
# 出厂编号
# 产品代号
# 制造国家
# 出厂日期
# 投运日期
# 使用环境
# 绝缘耐热等级
# 资产编号
# 用途
# 绝缘介质
# 绕组型式
# 结构型式
# 冷却方式
# 调压方式
# 安装位置
# 额定容量(MVA)
# 自冷却容量(%)
# 电压比
# 额定电流(中压)(A)
# 额定电流(低压)(A)
# 短路阻抗高压－中压(%)
# 短路阻抗高压－低压(%)
# 短路阻抗中压－低压(%)
# 空载损耗(kV)
# 负载损耗(实测值)(满载)(kW)
# 自然冷却噪声(dB)
# 总重(T)
# 油号
# 油重
# 油产地
# SF6气体额定压力(Mpa)
# SF6气体报警压力(Mpa)
# 运行状态
# 最近投运日期
# 累计调档次数
# 备注
# 退运日期
# 审核状态
# 设备编码

pms={'设备名称':{'#1主变','#2主变','#3主变'},'运行编号':{'#1主变压器','#2主变压器','#3主变压器'},'资产性质':{'省（直辖市、自治区）公司'},
    '资产单位':{'国网湖南省电力公司','湖南省电力公司'},'电压等级':{'交流220kV','交流110kV','交流35kV'},'间隔单元':{'#1主变','#2主变','#3主变'},
    '相数':{'三相'},'相别':{'ABC相'},'自造国家':{'中国'},'使用环境':{'户外式'},'绝缘耐热等级':{'A','B'},'用途':{'降压变压器'},
    '结构形式':'芯式','冷却方式':{'强迫油循环导向风冷(ODAF)','自然油循环风冷(ONAF)','自然冷却(ONAN'},'调压方式':{'有载调压','无励磁调压'},
        }

def _pms_validate(filename):
    #Excel表格应该共有56列
    cell_error = easyxf('pattern: pattern solid, back_colour yellow')
    workbook =xlrd.open_workbook(filename)
    pms_sheet = workbook.sheet_by_index(0)
    ncols = pms_sheet.ncols
    nrows = pms_sheet.nrows
    i = 1 #跳过表头
    
    
    return (True,None)

def pms_validate(filename):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filename)
    # Alternately, specify the full path to the workbook 
    # wb = excel.Workbooks.Open(r'C:\myfiles\excel\workbook2.xlsx')
    excel.Visible = False
    
    wb.Cells(1,1).Interior.ColorIndex = 2
    execel.Application.Quit()
    
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
    pms_validate('E:\Codes\mygithub\excel_handler\主变压器.xls')

