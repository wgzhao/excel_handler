# -*- coding:utf-8 -*-
import os
#import win32com.client as win32
import sys
import re
import time
import shutil
import collections
codec=sys.getfilesystemencoding()
reload(sys) 
sys.setdefaultencoding("utf-8")

class Validate():
    def validate():
        '''
        电压互感器数据校验
        '''
        #错误单元格列表，记录行列值，用于最后的涂色
        err_cells = []
        #错误行数
        total_error_lines = 0
        ## Notice: 下标从0开始，即第一个单元坐标为(0,0)
        self.pms={'5':u'省（直辖市、自治区）公司',
            '6':[u'国网湖南省电力公司',u'湖南省电力公司'],'8':[u'交流220kV',u'交流110kV',u'交流35kV',u'交流10kV'],
            '19':u'中国','22':u'户外式','23':[u'A',u'B'],'25':u'降压变压器','26':u'油浸','27':[u'双绕组',u'三绕组'],
            '28':u'芯式','29':[u'强迫油循环导向风冷(ODAF)',u'自然油循环风冷(ONAF)',u'自然冷却(ONAN)'],'30':[u'有载调压',u'无励磁调压'],
            '46':u'新疆克拉玛依石油'}
        
        #出产编号，用于校验是否有重复的出厂编号,为保证下标从1开始,第一个元素为虚构元素
        ccbh_list = ['this is sheet header,ignore it']
        
        #skip header
        fd = open('validate.log','w')
        for row in range(1,self.nrows):
            
            fd.write("validating %d/%d total error lines %d \n " % (row,self.nrows,total_error_lines))
            #print "validating %d/%d" % (row,self.nrows)
            the_line_is_error = False
            
            #1. 运行编号
            # 与设备名称(第一列)中数字不一致为不合格数据
            ## 分别取出设备名称和运行编号的数字（一般数字在开头)
            sbmc = self._getcell(row,0)
            sbmc_digital = re.findall(r'\d+',sbmc)[0]
            yxbm_digital = re.findall(r'\d+',self._getcell(row,1))[0]
            
            if sbmc_digital != yxbm_digital:
                err_cells.append((row,1))
                the_line_is_error = True
                
            #5. 资产性质
            if self._getcell(row,5) != self.pms['5']:
                err_cells.append((row,5))
                the_line_is_error = True
                
            #6. 资产单位
            if self._getcell(row,6) not in  self.pms['6']:
                err_cells.append((row,6))
                the_line_is_error = True
            
            
            #8. 电压等级
            dydj = self._getcell(row,8)
            if dydj not in  self.pms['8']:
                dydj_valid = False
                err_cells.append((row,8))
                the_line_is_error = True
            
            #9. 间隔单元
            # 与“设备名称”中数字不一致为不合格数据
            jgdy_digital = re.findall(r'\d+',self._getcell(row,9))[0]
            if jgdy_digital != sbmc_digital:
                err_cells.append((row,9))
                the_line_is_error = True
                
            
            #11. 相别
            ## 1.空白项为不合格数据       
            ## 2.“电压等级”为110kV、220kV 35kV的数据对应“设备名称”中“相”字前的字母A,B,C,O，否则为不合格数据；      
            ## 3.“电压等级”为10kV对应ABC，否则为不合格数据；
            xiangbie = self._getcell(row,11)
            result = True
            #规则1校验
            if xiangbie == '':
                result = False
            #规则3校验
            elif dydj == u'交流10kV' and xiangbie != u'ABC相':
                result = False
            elif dydj in self.pms['8'][:-1]:
                if sbmc.find(xiangbie) < 0:
                    result = False
            else:
                result = False
            
            if result == False:
                err_cells.append((row,11))
                the_line_is_error = True
                
            
            #10. 相数
            ## 1.空白项为不合格数据        
            ## 2.“相别”为A、B、C或者O对应单相，否则为不合格数据；“相别” 为ABC的对应三相，否则为不合格数据；         
            ## 3.“电压等级”为110kV、220kV 35kV对应单相，否则为不合格数据；         
            ## 4.“电压等级”为10kV对应三相，否则为不合格数据；                  "
            xiangshu = self._getcell(row,10)
            result = True
            if xiangshu == '':
                result = False
            elif dydj == u'交流10kV' and xiangshu != u'三相':
                result = False
            elif dydj in self.pms['8'][:-1] and xiangshu != u'单相':
                result = False
            else:
                xiangbie_symbol = xiangbie[:-1] #去掉相字
                #xiangshu_symbol = xiangshu[:-1] #去掉相字
                if xiangbie in ['A','B','C','O'] and xiangshu != u'单相':
                    result = False
                elif xiangbie == 'ABC' and xiangshu != u'三相':
                    result = False
                    
            if result == False:
                err_cells.append((row,10))
                the_line_is_error = True
                
                
            #14. 设备型号
            ## 不能为空白项
            sbxh = self._getcell(row,14)
            if sbxh == '':
                err_cells.append((row,14))
                the_line_is_error = True
                
            #12. 额定电压(kV)
            ## 1.空白项为不合格数据      
            ## 2. 对应“设备型号”中存在√3的数据额定电压为√3带上其前面的数字组成的数值，否则为不合格数据；
            ## 3.对应“设备型号”不存在√3的数据对应其“电压等级”的数据，如10kV为10、35kV为35，110kV为110、220kV为220，否则为不合格数据"
            
            eddy = self._getcell(row,12)
            result = True
            if eddy == '':
                result  = False
            elif sbxh.find('√3') > -1:
                if sbxh.find(eddy) < 0:
                    result = False
            else:
                eddyplus = eddy + 'kV'
                if dydj.find(eddyplus) < 0:
                    result = False
                    
            if result == False:
                err_cells.append((row,12))
                the_line_is_error = True
                
            #13. 额定频率
            ## 50以外不合格
            edpl = self._getcell(row,13)
            if edpl != '50':
                err_cells.append((row,13))
                the_line_is_error = True
                
            #15. 生产厂家
            ## 1、空白为不合格数据；
            ## 2、小于6个字的为不合格数据
            if len(self._getcell(row,15)) < 6:
                err_cells.append((row,15))
                the_line_is_error = True
                
            #16. 出厂编号
            ##1、空白为不合格数据；
            ##2、重复出现的为不合格数据 
            ccbh = self._getcell(row,16) 
            if  ccbh == '':
                err_cells.append((row,16))
                the_line_is_error = True
            else:
                try:
                    duprow = ccbh_list.index(ccbh)
                    #编号有重复把重复的单元都标记出来
                    err_cells.append((duprow,16))
                    err_cells.append((row,16))

                except ValueError:
                    ccbh_list.append(ccbh)

            #18. 制造国家
            ## 中国以外为不合格数据
            if self._getcell(row,18) != self.pms['20']:
                err_cells.append((row,18))
                the_line_is_error = True
            
            #19. 出厂日期
            ## 空白为不合格数据
            ccrq_ts = False
            ccrq = self._getcell(row,19) 
            if self._getcell(row,19) == '':
                err_cells.append((row,19))
                the_line_is_error = True
            else:
                try:
                    ccrq_ts = time.mktime(time.strptime(ccrq,'%Y-%m-%d %H:%M:%S'))
                except Exception,err:
                    err_cells.append((row,19))
                    the_line_is_error = True
                    ccrq_ts = False
                
            #20. 投运日期
            ## 1. 空白为不合格数据
            ## 2. 比出厂日期小10天的为不合格数据"
            tendays = 86400
            #把出厂日期和投运日期都转成时间戳，然后进行比较
            tyrq = self._getcell(row,20)
            if tyrq == '':
                err_cells.append((row,20))
                the_line_is_error = True
            else:
                try:
                    tyrq_ts = time.mktime(time.strptime(tyrq,'%Y-%m-%d %H:%M:%S'))
                except Exception,err:
                    err_cells.append((row,20))
                    the_line_is_error = True
                    tyrq_ts = False
                if tyrq_ts  and ccrq_ts  and (tyrq_ts - ccrq_ts) < tendays:
                     err_cells.append((row,20))
                     the_line_is_error = True
    
            #21. 使用环境
            ## 1.空白项为不合格数据
            if self._getcell(row,21) == '':
                err_cells.append(row,21)
                the_line_is_error = True
                
            
            #23. 组合设备类型
            ## 1.空白项为不合格数据；
            ## 2.“电压等级”为10kV为组合电器、10kV以外为开关柜的为不合格数据
            zhsblx = self._getcell(row,23)
            result = True
            if zhsblx == '':
                result = False
            elif (zhsblx == u'组合电器' and dydj == u'交流10kV') or (zhsblx == u'开关柜' and dydj != u'交流10kV'):
                result = False
                
            if result == False:    
                err_cells.append((row,23))
                the_line_is_error = True
                
            #24. 组合电器(开关柜)名称
            ## 1.“组合电器类型”为否对应此字段不为空白项的数据为不合格数据
            ## 2.“组合电器类型”为开关柜对应此字段内未包含“开关柜”为不合格数据；
            ## 3.“组合电器类型”为组合电器对应此字段内未包含“电器”为不合格数据；
            ## FIXME: 怀疑这里说的组合电器类型疑为组合设备类型，表格中找不到组合电器类型一列
            zhdqmc = self._getcell(row,24)
            result = True
            if zhsblx == u'否' and len(zhdqmc) > 0:
                result = False
            elif zhsblx == u'开关柜' and zhdqmc.find(zhsblx) < 0:
                result = False
            elif zhsblx == u'组合电器' and zhdqmc.find(u'电器') < 0:
                result = False
                
            if result == False:
                err_cells.append((row,24))
                the_line_is_error = True
                
                
                
            