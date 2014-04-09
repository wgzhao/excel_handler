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
  
from base import PmsBase

class Validate(PmsBase):
    def validate(self):
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
            '18':u'中国','22':u'户外式','23':[u'A',u'B'],'25':u'降压变压器','26':u'油浸','27':[u'双绕组',u'三绕组'],
            '28':u'芯式',
            '46':u'新疆克拉玛依石油'}
        
        #出产编号，用于校验是否有重复的出厂编号,为保证下标从1开始,第一个元素为虚构元素
        ccbh_list = ['this is sheet header,ignore it']
        
        #查找数字
        find_digits = re.compile(r'\d+')
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
            result = True
            try:
                sbmc_digital = find_digits.findall(sbmc)[0] #re.findall(r'\d+',sbmc)[0]
                yxbh_digital = find_digits.findall(self._getcell(row,1))[0] 
                if sbmc_digital != yxbh_digital:
                    result = False
            except IndexError,err:
                #print "sbmc = %s and yxbh = %s" %(sbmc,self._getcell(row,1))
                result = False
                
            if result == False:
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
            if self._getcell(row,18) != self.pms['18']:
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
                
            #27. 结构形式
            ## 1.空白项为不合格数据；；      
            ## 2.组合设备类型中“组合电器”对应“电磁式”，否则为不合格数据  
            ## 3.设备型号以“J”开头的TV，应为“电磁式”，型号以“T”开头的TV，应为“电容式”  ，否则为不合格数据
            ## 4.电压等级10kV对应“电磁式”，否则为不合格数据"
            
            jgxs = self._getcell(row,27)
            result = True
            if jgxs == '':
                result = False
            elif zhsblx == u'组合电器' and jgxs != u'电磁式':
                result = False
            elif ( sbxh[0].upper() == 'J' and jgxs != u'电磁式') or ( sbxh[0].upper() == 'T' and jgxs != u'电容式'):
                result = False
            elif dydj == u'交流10kV' and jgxs != u'电磁式':
                result = False
                    
            if result == False:
                err_cells.append((row,27))
                the_line_is_error = True
                
                
            #25. 绝缘介质
            ## 1.空白项为不合格数据；      
            ## 2.结构型式为“电容式”，绝缘介质类型应为“油浸”，否则为不合格数据；    
            ## 3.组合设备类型中“组合电器”对应“SF6”，否则为不合格数据"    
            
            jyjz = self._getcell(row,25)
            result = True
            
            if jyjz == '':
                result = False
            elif jgxs == u'电容式' and jyjz != u'油浸':
                result = False
            elif zhsblx == u'组合电器' and jyjz != u'SF6':
                result = False
                
            if result == False:
                err_cells.append((row,25))
                the_line_is_error = True
                
            #26. 外绝缘型式
            ## 1.空白项为不合格数据；       
            ## 2.组合设备类型中“组合电器”对应“环氧树脂”，否则为不合格数据   
            
            wjyxs = self._getcell(row,26)
            if wjyxs ==''  or (zhsblx == u'组合电器' and wjyxs != u'环氧树脂'):
                err_cells.append((row,26))
                the_line_is_error = True
                
            #28. 铁芯结构
            ## 1.空白项为不合格数据；     
            ## 2.相数“单相”对应“单柱式”，否则为不合格数据
            ## 3.电压等级110kV、220kV对应“单柱式”，否则为不合格数据
            
            txjg = self._getcell(row,28)
            if (txjg == '') or ((xiangshu == u'单相'  or dydj in [u'交流220kV',u'交流110kV']) and txjg != u'单柱式'):
                err_cells.append((row,28))
                the_line_is_error = True
            
            #29. 是否全绝缘
            ## 1.空白项为不合格数据；     
            ## 2.电压等级110kV、220kV对应“否”，否则为不合格数据
            
            sfqjy = self._getcell(row,29)
            if (sfqjy == '') or (dydj in [u'交流220kV',u'交流110kV'] and sfqjy != u'否'):
                err_cells.append((row,28))
                the_line_is_error = True
                
            #30. 额定电压比
            ## 1.空白项为不合格数据；      
            ## 2.其中含有“kV”为不合格数据；
            ## 3.第一个“/”前的数值与电压等级应相等，否则为不合格数据"
            eddyb = self._getcell(row,30)
            result = True
            if eddyb == '':
                result = False
            elif eddyb.upper().find('KV') > -1:
                result = False
            else:
                eddyb_v = u"交流%skV" % eddyb.split('/')[0]
                if eddyb_v != dydj:
                    result = False
                    
            if result == False:
                err_cells.append((row,30))
                the_line_is_error = True
                
            #31. 二次绕组总数量
            ## 1.空白项为不合格数据
            ## 2.二次绕组总数量与额定电压比中“0.1”的个数相等，否则为不合格数据
            ecrzzsl = self._getcell(row,31)
            result = True
            if ecrzzsl == '' or not self._isnumber(ecrzzsl):
                result = False
            else:
                try:
                    if len(re.findall('0.1',eddyb)) != int(ecrzzsl):
                        result = False
                except Exception:
                    pass
                    
            if result == False:
                err_cells.append((row,31))
                the_line_is_error = True
                
            #32. 爬电比距(mm/kV)
            ## 1.空白项为不合格数据      
            ## 2.组合设备类型中为“组合电器”的填写“0” ，否则为不合格数据  
            ## 3.组合设备类型中不为“组合电器”中电压等级10kV对应16、20，35kV对应20、25，110kV、220kV对应25、31，否则为不合格数据
            v = {u'交流10kV':['16','20'],u'交流35kV':['20','25'],u'交流110kV':['25','31'],u'交流220kV':['25','31']}
        
            pdbj = self._getcell(row,32)
            result = True
            try:
                if pdbj == '':
                    result = False
                elif zhsblx == u'组合电器' and pdbj != '0':
                    result = False
                elif zhsblx != u'组合电器' and pdbj not in v[dydj]:
                    result = False
            except Exception,err:
                result = False
                
            if result == False:
                err_cells.append((row,32))
                the_line_is_error = True
            
            #33. 总额定电容量(pF)
            ## 1.空白项为不合格数据；
            ## 2.结构型式为“电磁式”对应“0”，结构型式为“电容式”对应大于1000，否则为不合格数据
            zeddrl = self._getcell(row,33)
            result = True
            if not self._isnumber(zeddrl):
                result = False
            else:
                zeddrl_numeric = float(zeddrl)
                if (jgxs == u'电磁式' and zeddrl_numeric != 0) or (jgxs == u'电容式' and zeddrl_numeric <= 1000):
                    result = False
                    
            if result == False:
                err_cells.append((row,33))
                the_line_is_error = True
                
            #34. 电容器节数(节)
            ## 1.空白项为不合格数据      
            ## 2.结构型式为“电磁式”对应“0”，否则为不合格数据；
            ## 3.结构型式为“电容式”，电压等级220kV对应1或者2，其他等级为“1”，否则为不合格数据
            
            drqjs = self._getcell(row,34)
            result = True
            if not drqjs.isdigit():
                result = False
            elif jgxs == u'电磁式' and drqjs != '0':
                result = False
            elif jgxs == u'电容式' and ((dydj == u'交流220kV' and drqjs not in ['1','2']) or (dydj != u'交流220kV' and drqjs != '1')):
                result = False
                
            if result == False:
                err_cells.append((row,34))
                the_line_is_error = True
                
                
            #35. 上节电容量(pF)
            ## 1.空白项为不合格数据     
            ## 2.结构型式为“电磁式”对应“0”；           
            ## 3.结构型式为“电容式”，电容器节数为2节，数值应大于总额定容量，否则为不合格数据；
            ## 3.结构型式为“电容式”，电容器节数为1节，数值应为0，否则为不合格数据；
            sjdrl = self._getcell(row,35)
            result = True
            if not sjdrl.isdigit():
                result = False
            else:
                sjdrl_numeric = int(sjdrl)
                if jgxs == u'电磁式' and drqjs == '2' and sjdrl_numeric <= zeddrl_numeric:
                    result = False
                elif jgxs == u'电容式' and drqjs == '1' and sjdrl_numeric != 0:
                    result = False
                    
            if result == False:
                err_cells.append((row,35))
                the_line_is_error = True
                
            #36. 中节电容量(pF)
            ## 1.统一填“0”，否则为不合格数据；
            if self._getcell(row,36) != '0':
                err_cells.append((row,36))
                the_line_is_error = True
            
            #38. 下节C2电容量(pF)
            ## 1.不应有空白项；     
            ## 2.结构型式为“电磁式”对应“0”；           
            ## 3.结构型式为“电容式”，应大于总额定电容量，否则为不合格数据
            xjc2drl = self._getcell(row,38)
            result = True
            if not self._isnumber(xjc2drl):
                result = False
            elif jgxs == u'电磁式' and xjc2drl != '0':
                result = False
            elif jgxs == u'电容式' and float(xjc2drl) <= zeddrl_numeric:
                result = False
                
            if result == False:
                err_cells.append((row,38))
                the_line_is_error = True
                      
            #37. 下节C1电容量(pF)
            ## 1.空白项为不合格数据      
            ## 2.结构型式为“电磁式”对应“0”；           
            ## 3.结构型式为“电容式”，应大于总额定电容量，否则为不合格数据； 
            ## 4.下节C1电容量应小于下节C2电容量，否则为不合格数据
            
            xjc1drl = self._getcell(row,37)
            result = True
            if not self._isnumber(xjc1drl):
                result = False
            elif jgxs == u'电磁式' and xjc1drl != '0':
                result = False
            else:
                xjc1drl_numeric = float(xjc1drl)
                if jgxs == u'电容式' and xjc1drl_numeric <= zeddrl_numeric:
                    result = False
                elif xjc1drl_numeric > float(xjc2drl):
                    result = False
                    
            if result == False:
                err_cells.append((row,37))
                the_line_is_error = True
                
            #39. 油号
            ## 1.空白项为不合格数据
            ## 2、绝缘介质为油浸、油浸式此字段25以外为不合格数据；
            ## 3、绝缘介质非“油浸”、“油浸式”此字段“无”、“/”以外为不合格数据"
            youhao = self._getcell(row,39)
            result = True
            if youhao == '':
                result = False
            elif jyjz in [u'油浸',u'油浸式'] and youhao != '25':
                result = False
            elif jyjz not in [u'油浸',u'油浸式'] and youhao not in [u'无','/']:
                result = False
            
            if result == False:
                err_cells.append((row,39))
                the_line_is_error = True
            
            #41. SF6气体报警压力(Mpa)
            ## 1.空白项为不合格数据      
            ## 2.绝缘介质为“SF6”，数值在0.30至0.60之间，其余绝缘介质的填“0”，否则为不合格数据； ；  
            sf6qtbjyl = self._getcell(row,41)
            result = True
            if not self._isnumber(sf6qtbjyl):
                result = False
            else:
                sf6qtbjyl_numeric = float(sf6qtbjyl) 
                if (jyjz == u'SF6' and (sf6qtbjyl_numeric < 0.30 or sf6qtbjyl_numeric > 0.60)) or \
                       (jyjz != u'SF6' and sf6qtbjyl_numeric != 0):
                       result = False
            if result == False:
                err_cells.append((row,41))
                the_line_is_error = True
                
            #40. SF6气体额定压力(Mpa)
            ## 1.空白项为不合格数据  
            ## 2.绝缘介质为“SF6”，数值在0.35至0.65之间，其余绝缘介质的填“0”，否则为不合格数据； 
            ## 3.SF6气体额定压力应大于SF6气体报警压力，否则为不合格数据
            sf6qtedyl = self._getcell(row,40)
            result = True
            if not self._isnumber(sf6qtedyl):
                result = False
            else:
                sf6qtedyl_numeric = float(sf6qtedyl)
                if sf6qtedyl_numeric >0 and sf6qtedyl_numeric <= sf6qtbjyl_numeric:
                    result = False
                if (jyjz == u'SF6' and (sf6qtedyl_numeric < 0.30 or sf6qtedyl_numeric > 0.60)) or \
                    (jyjz != u'SF6' and sf6qtedyl_numeric != 0):
                    result = False
                    
            if result == False:
                err_cells.append((row,40))
                the_line_is_error = True
                
            #43. 最近投运日期
            ## 空白或早于投运日期的为不合格数据
            zjtyrq = self._getcell(row,43)
            result = True
            # 计算投运日期
            tyrq_ts = time.mktime(time.strptime(self._getcell(row,20),'%Y-%m-%d %H:%M:%S'))
            if zjtyrq == '':
                result = False    
            else:
                try:
                    zjtyrq_ts = time.mktime(time.strptime(zjtyrq,'%Y-%m-%d %H:%M:%S'))
                    if tyrq_ts and zjtyrq_ts < tyrq_ts:
                        result = False
                except ValueError:
                    result = False
            if result == False:
                err_cells.append((row,43))
                the_line_is_error = True 
            
            if the_line_is_error == True:
                total_error_lines += 1
                        
            # end for 一行校验结束
        
        fd.close()
        #错误行数，总行数，错误单元列表
        return (total_error_lines,self.nrows,err_cells)    
                