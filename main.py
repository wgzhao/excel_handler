#!/usr/bin/env python
# -*- coding:utf-8 -*-
__author__ = 'wgzhao<wgzhao@gmail.com>'

from Tkinter import *
import tkFileDialog
import os
#import win32com.client as win32
import xlrd
from xlwt import Workbook,easyxf
from  xlutils.copy import copy
import sys
import time
import shutil
import collections
codec=sys.getfilesystemencoding()
reload(sys) 
sys.setdefaultencoding("utf-8")
__version__ = 1.0
ERROR_COLOR=27
DEBUG = True

ERROR_STYLE = easyxf('pattern:pattern solid,fore_colour yellow;')    

class ExcelHandler():

    def __init__(self,geometry='800x600'):
        self.win = Tk()
        self.win.title(u'电力设备数据检查程序')
        self.win.geometry(geometry)
        self.info = StringVar()
        self.lb = Label(self.win,textvariable=self.info)
        self.lb.pack()
        self.info.set(' ')

        self.menubar = Menu(self.win)

        # 菜单以一览表
        # 变电台账核查
        #   - 主变压器
        #   - 断路器
        #   - 隔离开关
        #   - 电流互感器
        #   - 电压互感器
        #   - 耦合电容器
        #   - 所用变
        #   - 楼地变
        #   - 母线
        #   - 电抗器
        #   - 组合电器
        #   - 阻波器
        #   - 放电线圈
        #   - 电力电缆
        #   - 接地网
        #   - 开关柜
        #   - 避雷器
        #   - 避雷针
        #   - 电力电容器
        #   - 消弧线圈
        # 输电台账核查
        #   - 线路
        #   - 杆塔
        #   - 绝缘子
        #   - 金具
        #   - 辅助设备
        #   - 导线
        #   - 电缆
        #   - 电缆头
        # 配电台账核查
        #   - 线路
        #   - 杆塔
        #   - 导线
        #   - 电缆段
        #   - 柱上变压器
        #   - 柱上负荷开关
        #   - 柱上隔离开关
        #   - 柱上断路器
        self.menulist = collections.OrderedDict()
        self.menulist[u'变电台账核查'] = [(u'主变压器','bdtz.zhubianyaqi'),(u'断路器','bdtz.duanluqi'),(u'隔离开关','bdtz.gelikaiguan'),
                                 (u'电流互感器','bdtz.dianliuhuganqi'),(u'电压互感器','bdtz.dianyahuganqi'),(u'耦合电容器','bdtz.ouhedianrongqi'),
                                 (u'所用变','bdtz.suoyongbian'),(u'楼地变','bdtz.loudibian'),(u'母线','bdtz.muxian'),(u'电抗器','bdtz.diankangqi'),
                                 (u'组合电器','bdtz.zuhedianqi'),(u'阻波器','bdtz.zuboqi'),(u'放电线圈','bdtz.fangdianxianquan'),(u'电力电缆','bdtz.dianlidianlan'),
                                 (u'接地网','bdtz.jiediwang'),(u'开关柜','bdtz.kaiguangui'),(u'避雷器','bdtz.bileizhen'),
                                 (u'电力电容器','bdtz.dianlidianrongqi'),(u'消弧线圈','bdtz.xiaohuxianquan')]
        self.menulist[u'输电台账核查']   = [(u'线路','sdtz.xianlu'),(u'杆塔','sdtz.ganta'),(u'绝缘子','sdtz.jueyuanzi'),(u'金具','sdtz.jinju'),
                                (u'辅助设备','sdtz.fuzhushebei'),(u'导线','sdtz.daoxian'),(u'电缆','sdtz.dianlan'),(u'电缆头','sdtz.dianlantou')]
        self.menulist[u'配电台账核查']   =  [(u'线路','pdtz.xianlu'),(u'杆塔','pdtz.ganta'),(u'电缆段','pdtz.dianlanduan'),
                                (u'柱上变压器','pdtz.zhushangbianyaqi'),(u'柱上负荷开关','pdtz.zhushangfuhekaiguan'),(u'柱上隔离开关','pdtz.zhushanggelikaiguan'),
                                (u'柱上断路器','pdtz.zhushangduanluqi')]
        
        for m in self.menulist.keys():
            tmp_menu = Menu(self.menubar,tearoff=0)
            #cascade menu
            for item in self.menulist[m]:
                tmp_menu.add_command(label=item[0],command=lambda i=item[1]: self.openexcel(i))
            self.menubar.add_cascade(label=m,menu=tmp_menu)
            
            del tmp_menu
        
        self.win.config(menu=self.menubar)
        self.win.mainloop()


        
        
    def openexcel(self,check_type):
        
        '''
            根据用户点击不同的类型，进行不同的校验，并给出校验结果
        '''
        # 将传递过来的当做Python语法写入临时文件，然后import这个临时文件，从而达到依据用户传递的参数来import指定的类的目的
        # 比如如果用户点击“主变压器”菜单，则将from bdtz.zhubianyaqi import Validate语句写入临时配置文件config.py
        # 而后import config.py文件即可。
        # 约定所有的类的名字均为Validate，真正的校验方法名为validate即可
        open('pmsconfig.py','w').write('from %s import Validate \n' % check_type)
        #import a excel file
        ftypes=[("97-2003 Excel files",'*.xls'),('2007 Excel files','*.xlsx')]
        fn = tkFileDialog.askopenfilename(parent=self.win,title='Choose a file',filetypes=ftypes)

        if fn != '':
            from pmsconfig import Validate
            pmshandler = Validate(fn)
            (total_error_lines,total_lines,err_cells) = pmshandler.validate()
            
            #将错误信息写入校验文件
            if total_error_lines > 0:
                self.wb = xlrd.open_workbook(fn,formatting_info=True)
                self.ws = self.wb.sheet_by_index(0)
                towb = copy(self.wb)
                tows = towb.get_sheet(0)
             
                for cell in err_cells:
                    #self.ws.Cells(cell[0],cell[1]).Interior.ColorIndex = ERROR_COLOR
                    tows.write(cell[0],cell[1],self.ws.cell(cell[0],cell[1]).value,style=ERROR_STYLE)
                
                filename,fileext = os.path.splitext(fn)
                newfilepath = filename + u'_valid' + fileext
                
                towb.save(newfilepath)
            
                self.info.set(u"共检查 %d 行\n其中包含错误数据的有 %d 行\n 详细情况请打开 %s 文件了解" %(total_lines,total_error_lines,newfilepath))
            else:
                self.info.set(u"恭喜，%s 文件检验全部痛过，没有错误！" % fn)
        else:    
            self.info.set(u'无法打开 %s' % fn)

    def about(self):
        win_about = Toplevel(self.win)
        lb_about = Label(win_about,text=u"电力设备数据检查程序 %f" % __version__)
        lb_about.pack()

    def help(self):
        print("usage:")

def test():
    curdir = os.getcwd()
    from bdtz.zhubianyaqi import Validate
    pmshandler = Validate(os.path.join(curdir,u'主变压器.xls'))
    (total_error_lines,total_lines,err_cells) = pmshandler.validate()
    print "total error lines = %d / total lines %d" % (total_error_lines,total_lines)
if __name__ == '__main__':
    handler= ExcelHandler('600x400')
    #test()
    #test code
    # begin_time = time.time()
    # curdir=os.getcwd()
    # pmshandler = PmsValidate(os.path.join(curdir,u'主变压器.xls'))
    # pmshandler.pms_validate()
    # estime = time.time() - begin_time
    # print "total use %.1f seconds" % estime
