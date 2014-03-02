#!/usr/bin/env python
# -*- coding:gbk -*-
__author__ = 'wgzhao<wgzhao@gmail.com>'

from Tkinter import *
import tkFileDialog
import xlrd 
import xlwt
from xlwt import easyxf
import win32com.client as win32
#����ʡ����PMS̨�ʲ���˵��
#keyΪ�������ƣ�valueΪ���������ֵ����Сд���У�����֮�⣬������������
#1. ���б��Ӧ���豸����һ�£�����豸����Ϊ'#1����'�������б�ű���Ϊ'#1����ѹ��'
#2. �����Ԫ�������豸����һ��
#3. ���ѹ�͵�ѹ�ȼ�����ӳ���ϵ���������£�
##    1������ѹ�ȼ���Ϊ����220kV��Ӧ��242��230��220������Ϊ���ϸ����ݣ�
##    2������ѹ�ȼ���Ϊ����110kV��Ӧ��121��110������Ϊ���ϸ����ݣ�
##    3������ѹ�ȼ���Ϊ����35kV��Ӧ��38.5��35������Ϊ���ϸ����ݣ�"
#4. �豸�ͺź͵�ѹ�ȼ���ӳ���ϵ����ϵ���£�
##         1������ѹ�ȼ���Ϊ����220kV����ĩβ��Ӧ��180000/220��120000/220��240000/220������Ϊ���ϸ����ݣ�
##        2������ѹ�ȼ���Ϊ����110kV����ĩβ��Ӧ��20000/110��31500/110��50000/110������Ϊ���ϸ����ݣ�
##        3������ѹ�ȼ���Ϊ����35kV����ĩβ��Ӧ��3150/35��4000/35��5000/35��6300/35��10000/35������Ϊ���ϸ�����


#########
#��˳��
# �豸����
# ���б��
# �����о�
# ���е�λ
# ���վ
# �ʲ�����
# �ʲ���λ
# �豸����
# ��ѹ�ȼ�
# �����Ԫ
# ����
# ���
# ���ѹ(kV)
# �����(A)
# �Ƶ��(Hz)
# �豸�ͺ�
# ��������
# �������
# ��Ʒ����
# �������
# ��������
# Ͷ������
# ʹ�û���
# ��Ե���ȵȼ�
# �ʲ����
# ��;
# ��Ե����
# ������ʽ
# �ṹ��ʽ
# ��ȴ��ʽ
# ��ѹ��ʽ
# ��װλ��
# �����(MVA)
# ����ȴ����(%)
# ��ѹ��
# �����(��ѹ)(A)
# �����(��ѹ)(A)
# ��·�迹��ѹ����ѹ(%)
# ��·�迹��ѹ����ѹ(%)
# ��·�迹��ѹ����ѹ(%)
# �������(kV)
# �������(ʵ��ֵ)(����)(kW)
# ��Ȼ��ȴ����(dB)
# ����(T)
# �ͺ�
# ����
# �Ͳ���
# SF6����ѹ��(Mpa)
# SF6���屨��ѹ��(Mpa)
# ����״̬
# ���Ͷ������
# �ۼƵ�������
# ��ע
# ��������
# ���״̬
# �豸����

ERROR_COLOR=27
pms={'�豸����':{'#1����','#2����','#3����'},'���б��':{'#1����ѹ��','#2����ѹ��','#3����ѹ��'},'�ʲ�����':{'ʡ��ֱϽ�С�����������˾'},
    '�ʲ���λ':{'��������ʡ������˾','����ʡ������˾'},'��ѹ�ȼ�':{'����220kV','����110kV','����35kV'},'�����Ԫ':{'#1����','#2����','#3����'},
    '����':{'����'},'���':{'ABC��'},'�������':{'�й�'},'ʹ�û���':{'����ʽ'},'��Ե���ȵȼ�':{'A','B'},'��;':{'��ѹ��ѹ��'},
    '�ṹ��ʽ':'оʽ','��ȴ��ʽ':{'ǿ����ѭ���������(ODAF)','��Ȼ��ѭ������(ONAF)','��Ȼ��ȴ(ONAN'},'��ѹ��ʽ':{'���ص�ѹ','�����ŵ�ѹ'},
        }

def _pms_validate(filename):
    #Excel���Ӧ�ù���56��
    cell_error = easyxf('pattern: pattern solid, back_colour yellow')
    workbook =xlrd.open_workbook(filename)
    pms_sheet = workbook.sheet_by_index(0)
    ncols = pms_sheet.ncols
    nrows = pms_sheet.nrows
    i = 1 #������ͷ
    
    
    return (True,None)

def pms_validate(filename):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filename)
    # Alternately, specify the full path to the workbook 
    # wb = excel.Workbooks.Open(r'C:\myfiles\excel\workbook2.xlsx')
    excel.Visible = False
    ws = wb.Worksheets(1)
    ws.Activate()
    nrows = ws.UsedRange.Rows.Count
    ncols = ws.UsedRange.Columns.Count
    print nrows,ncols
    #ws.Cells(1,1).Interior.ColorIndex = ERROR_COLOR
    #wb.Save()
    excel.Application.Quit()
    
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
    pms_validate('E:\Codes\mygithub\excel_handler\����ѹ��.xls')

