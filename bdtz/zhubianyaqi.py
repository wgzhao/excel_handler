# -*- coding:utf-8 -*-
import os
#import win32com.client as win32

import sys
import time
import shutil
import collections
codec=sys.getfilesystemencoding()
reload(sys) 
sys.setdefaultencoding("utf-8")
  
from base import PmsBase
class Validate(PmsBase):
    
            
    def validate(self):
        self.pms={'1':[u'#1主变',u'#2主变',u'#3主变'],'2':[u'#1主变压器',u'#2主变压器',u'#3主变压器'],'6':u'省（直辖市、自治区）公司',
            '7':[u'国网湖南省电力公司',u'湖南省电力公司'],'9':[u'交流220kV',u'交流110kV',u'交流35kV'],'11':u'三相','12':u'ABC相',
            '20':u'中国','23':u'户外式','24':[u'A',u'B'],'26':u'降压变压器','27':u'油浸','28':[u'双绕组',u'三绕组'],
            '29':u'芯式','30':[u'强迫油循环导向风冷(ODAF)',u'自然油循环风冷(ONAF)',u'自然冷却(ONAN)'],'31':[u'有载调压',u'无励磁调压'],
            '47':u'新疆克拉玛依石油'}
       
        #额定电压校验规则,column 13
        self.eddy_dict = {u'交流220kV':['242','230','220'],u'交流110kV':['121','110'],u'交流35kV':['38.5','35']}
        
        #设备型号校验规则,column 16
        self.sbxh_dict = {u'交流220kV':['180000/220','120000/220','240000/220'],
                         u'交流110kV':['20000/110','31500/110','50000/110'],
                         u'交流35kV':['3150/35','4000/35','5000/35','6300/35','10000/35']}
        
        #错误单元格列表，记录行列值，用于最后的涂色
        err_cells = []
        #错误行数
        total_error_lines = 0
        #出产编号，用于校验是否有重复的出厂编号,为保证下标从1开始,第一个元素为虚构元素
        ccbh_list = ['this is sheet header,ignore it']
        
        #skip header
        fd = open('validate.log','w')
        for row in range(1,self.nrows):
            
            fd.write("validating %d/%d total error lines %d \n " % (row,self.nrows,total_error_lines))
            #print "validating %d/%d" % (row,self.nrows)
            the_line_is_error = False
            #电压等级值是否正确，该值后面用的较多，需要进行比较，如果电压等级值本身不正确，那么后面的基于它的校验就没有意义，直接跳过
            dydj_valid = True
            #额定电流是否正确
            eddl_valid = True
            #额定电压是否正确
            eddy_valid = True
            #额定容量
            edrl_valid = True
            #设备型号是否正确
            sbxh_valid = True
            #绕阻型式是否正确
            rzxs_valid = True
            #1. 设备名称  "#1主变、#2主变、#3主变 以外为不合格数据
            cell11=self._getcell(row,0)
            if cell11 not in self.pms['1']:
                err_cells.append((row,0))
                the_line_is_error = True
            #2. 运行编号
            ##1、#1主变压器、#2主变压器、#3主变压器以外为不合格数据；
            ##2、与设备名称不一致为不合格数据"
            cell12 = self._getcell(row,1)
            if cell12 not in self.pms['2']:
                err_cells.append((row,1))
                the_line_is_error = True
                
            elif cell12 != cell11[:4] + u'压器':
                err_cells.append((row,1))
                the_line_is_error = True
                
            #6. 资产性质  "省（直辖市、自治区）公司" 以外为不合格数据；
            if self._getcell(row,5)!= self.pms['6']:
                err_cells.append((row,5))
                the_line_is_error = True
        
            #7. 资产单位  "国网湖南省电力公司、湖南省电力公司" 以外为不合格数据
            if self._getcell(row,6) not in  self.pms['7']:
                err_cells.append((row,6))
                the_line_is_error = True
       
            #9. 电压等级   "交流220kV、交流110kV、交流35kV" 以外为不合格数据
            dydj = self._getcell(row,8)
            if dydj not in  self.pms['9']:
                dydj_valid = False
                err_cells.append((row,8))
                the_line_is_error = True
                
            #11. 相数 三相以外为不合格数据
            if self._getcell(row,10) != self.pms['11']:
                err_cells.append((row,10))
                the_line_is_error = True
            #12. 相别  ABC相以外为不合格数据
            if self._getcell(row,11) != self.pms['12']:
                err_cells.append((row,11))
                the_line_is_error = True
            
            #13. 额定电压
            ## 数值类型
            ## 1、“电压等级”为交流220kV对应有242、230、220，以外为不合格数据；
            ## 2、“电压等级”为交流110kV对应有121、110，以外为不合格数据；
            ## 3、“电压等级”为交流35kV对应有38.5、35，以外为不合格数据；
            eddy = self._getcell(row,12)
            result = True
            if self._isnumber(eddy):
                if (dydj_valid and eddy not in self.eddy_dict[dydj]):
                    eddy_valid = False
                eddy = float(eddy)
            else:
                eddy_valid = False
                result = False
                
            if result == False:
                err_cells.append((row,12))
                the_line_is_error = True              
            
            #15. 额定频率
            ## 50以外为不合格数据
            if int(self._getcell(row,14)) != 50:
                err_cells.append((row,14))
                the_line_is_error = True
                
            #16. 设备型号               
            ################################################################################################
            ##1 “电压等级”为交流220kV数据末尾对应有180000/220、120000/220、240000/220，以外为不合格数据；
            ##2 “电压等级”为交流110kV数据末尾对应有20000/110、31500/110、50000/110，以外为不合格数据；
            ##3 “电压等级”为交流35kV数据末尾对应有3150/35、4000/35、5000/35、6300/35、10000/35，以外为不合格数据；
            ## 设备型号类似 SFSZ10-180000/220  或 SSZ10-Z60-20000/110
            #################################################################################################
            sbxh = self._getcell(row,15)
            try:
                sbxh_suffix = sbxh.split('-')[-1] #last item
                
                if dydj in self.pms['9'] and  sbxh_suffix not in self.sbxh_dict[dydj]:
                    sbxh_valid = False
                    err_cells.append((row,15))
                    the_line_is_error = True
                        
            except Exception,err:
                sbxh_valid = False
                err_cells.append((row,15))
                the_line_is_error = True
                
            #17. 生产厂家
            ##"1、空白为不合格数据；
            ##2、小于6个字的为不合格数据"    
            xccj = self._getcell(row,16)
            if xccj == '' or len(xccj) < 6:
                err_cells.append((row,16))
                the_line_is_error = True

                
            #18. 出厂编号
            ##1、空白为不合格数据；
            ##2、重复出现的为不合格数据 
            ccbh = self._getcell(row,17) 
            if  ccbh == '':
                err_cells.append((row,17))
                the_line_is_error = True
            else:
                try:
                    duprow = ccbh_list.index(ccbh)
                    #编号有重复把重复的单元都标记出来
                    err_cells.append((duprow,17))
                    err_cells.append((row,17))

                except ValueError:
                    ccbh_list.append(ccbh)
                    
                    
            #20. 制造国家
            ## 中国以外为不合格数据
            if self._getcell(row,19) != self.pms['20']:
                err_cells.append((row,19))
                the_line_is_error = True
            
            #21. 出厂日期
            ## 空白为不合格数据
            ccrq_ts = False
            ccrq = self._getcell(row,20) 
            if self._getcell(row,20) == '':
                err_cells.append((row,20))
                the_line_is_error = True
            else:
                try:
                    ccrq_ts = time.mktime(time.strptime(ccrq,'%Y-%m-%d %H:%M:%S'))
                except Exception,err:
                    err_cells.append((row,20))
                    the_line_is_error = True
                    ccrq_ts = False

            #22. 投运日期
            #"1、空白为不合格数据
            #2、比出厂日期小10天的为不合格数据"
            tendays = 86400
            #把出厂日期和投运日期都转成时间戳，然后进行比较
            tyrq = self._getcell(row,21)
            if tyrq == '':
                err_cells.append((row,21))
                the_line_is_error = True
            else:
                try:
                    tyrq_ts = time.mktime(time.strptime(tyrq,'%Y-%m-%d %H:%M:%S'))
                except Exception,err:
                    err_cells.append((row,21))
                    the_line_is_error = True
                    tyrq_ts = False
                if tyrq_ts  and ccrq_ts  and (tyrq_ts - ccrq_ts) < tendays:
                     err_cells.append((row,21))
                     the_line_is_error = True
    
            #23. 使用环境
            ## 户外式以外为不合格数据
            if self._getcell(row,22) != self.pms['23']:
                err_cells.append((row,22))
                the_line_is_error = True
            
            #24. 绝缘耐热等级
            ## A或B以外为不合格数据
            if self._getcell(row,23) not in  self.pms['24']:
                err_cells.append((row,23))
                the_line_is_error = True
                
            #26. 用途
            ## 降压变压器以外为不合格数据
            if self._getcell(row,25)  != self.pms['26']:
                err_cells.append((row,25))
                the_line_is_error = True
             
            #27. 绝缘介质
            ## 油浸以外为不合格数据
            if self._getcell(row,26) != self.pms['27']:
                err_cells.append((row,26))
                the_line_is_error = True
            
            #28. 绕组型式
            ## "1、双绕组,三绕组以外为不合格数据；
            ## 2、“电压等级”为交流220kV对应三绕组，以外为不合格数据"
            ## 3. 设备型号中包含2个"S"对应绕组型式为"双绕组"的数据为不合格数据
            import re
            rzxs = self._getcell(row,27)
            if rzxs not in self.pms['28']: 
                rzxs_valid = False
                err_cells.append((row,27))
                the_line_is_error = True
            elif (dydj_valid == True and dydj == u'交流220kV' and rzxs != u'三绕组'):
                err_cells.append((row,27))
                the_line_is_error = True
            elif (sbxh_valid == True and len(re.split('S',sbxh.upper())) == 3 and rzxs == u'双绕组' ):
                err_cells.append((row,27))
                the_line_is_error = True
                
            #29. 结构型式
            ## 芯式以外为不合格数据
            if self._getcell(row,28) != self.pms['29']:
                err_cells.append((row,28))
                the_line_is_error = True
                
            #30. 冷却方式
            ## "1、强迫油循环导向风冷(ODAF)、自然油循环风冷(ONAF)、自然冷却(ONAN)以外为不合格数据；
            ## 2、“设备型号”中有“F”对应自然冷却(ONAN)的数据为不合格数据"
            lqfs = self._getcell(row,29)
            if (lqfs not in self.pms['30']) or (sbxh.find('F') > -1 and lqfs == u'自然冷却(ONAN)' ):
                err_cells.append((row,29))
                the_line_is_error = True
                
            #31. 调压方式    
            ## "1、有载调压、无励磁调压以外为不合格数据；
            ## 2、“设备型号”中有“Z”对应无励磁调压的数据为不合格数据"
            
            tyfs = self._getcell(row,30)
            if tyfs not in self.pms['31'] or (tyfs.find('Z') > -1 and lyfs == u'无励磁调压'):
                err_cells.append((row,30))
                the_line_is_error = True
            
            #32. 安装位置
            ##空白为不合格数据
            if self._getcell(row,31) == '':
                err_cells.append((row,21))
                the_line_is_error = True
                
            #33. 额定容量(MVA)
            ## 整数类型
            ## 不等于“设备型号”中“/”之前的数值/1000的数据为不合格数据,设备型号类似 SFSZ10-180000/220为180
            edrl = self._getcell(row,32)
            result = True
            if self._isnumber(edrl):
                if edrl.find('.') > -1 :
                    result = False
                    edrl_valid = False
                else:
                    edrl = int(edrl)
                    if sbxh_valid and int(sbxh.split('-')[-1].split('/')[0]) / 1000 != edrl:
                        result = False
            else:
                result = False
                edrl_valid = False    
   
            
            if result == False:
                err_cells.append((row,32))
                the_line_is_error = True
             
             
            #14. 额定电流
            ## 为数值类型
            ## "│额定电流-额定容量*1000/额定电压/1.732│大于1的数据为不合格数据；不为数字的为不合格数据"
            ## 额定容量为33列
            eddl = self._getcell(row,13)
            result = True
            if not self._isnumber(eddl):
                eddl_valid = False
                result = False
            else:
                eddl = float(eddl)
                if eddy_valid  and edrl_valid and abs((eddl - edrl*1000  )/(eddy   * 1.732)) > 1:
                    # FIXME: should not time 1000, I guess
                    #print "eddl is invalid: abs((%.2f - %.2f)/ (%.2f * 1.732)) = %.2f" % (eddl,edrl,eddy, abs((eddl - edrl )/(eddy   * 1.732)))
                    result = False
                    

            if result == False:
               
                err_cells.append((row,13))
                the_line_is_error = True
                    
            #34. 自冷却容量(%)
            ## 数值类型
            ##"1、“冷却方式”为自然冷却(ONAN)对应100，以外的为不合格数据；
            ##2、“冷却方式”为自然油循环风冷(ONAF)对应60-75之间，以外的为不合格数据；
            ##3、“冷却方式”为强迫油循环导向风冷(ODAF)对应0,以外的为不合格数据"    
            result = True
            zlqrl = self._getcell(row,33)
            if self._isnumber(zlqrl):
                zlqrl = float(zlqrl)
                if lqfs == u'自然冷却(ONAN)' and zlqrl != 100:
                    result = False
                elif lqfs == u'自然油循环风冷(ONAF)' and (zlqrl < 60 or zlqrl > 75):
                    result = False
                elif lqfs == u'强迫油循环导向风冷(ODAF)' and zlqrl != 0:
                    result = False
            else:
                result = False
                    
            if result == False:
                err_cells.append((row,33))
                the_line_is_error = True
                
            #35. 电压比
            # 例子： 110±8×1.25%/35±2×2.5%/10.5 or (110±8×1.25%)/10.5
            ## "参照下列要求查询不合格数据
            ## 1、×号不能是大些或者小写的X或者*，（利用搜狗输入法，输入ch，选3就出现了×）
            ## 2、括弧必须用半角括弧
            ## 3、百分号、斜杠等符号必须用半角
            ## 4、±不能填成+-
            ## 5、不能带kV等单位
            ## 6、不能有空格"
            dyb = self._getcell(row,34)
            if re.match(u'[\*x（）+-]+',dyb):
                #matched means error 
                err_cells.append((row,34))
                the_line_is_error = True
             
            #36. 额定电流(中压)(A)
            ## 为数值型 or 0
            ## 1、“电压等级”为交流220kV对应额定电流/额定电流(中压)在0.52-0.56之间，以外为不合格数据；
            ## 2、“电压等级”为交流110kV对应额定电流/额定电流(中压)在0.34-0.36之间，以外为不合格数据；
            ## 3、“绕组型式”为双绕组对应0，以外的为不合格数据；；"  
            result = True
            eddlzy = self._getcell(row,35)
            if not self._isnumber(eddlzy):
                retsul = False
            else:
                eddlzy = float(eddlzy)
                #规则3
                if eddlzy == 0  and rzxs != u'双绕组':
                    result = False
                # 规则1，2前提
                elif dydj_valid == True and eddl_valid == True and eddlzy > 0 :
                    #规则1，2
                    p = eddl / eddlzy
                    if (dydj == u'交流220kV' and (p <0.52 or p > 0.56)) or (dydj == u'交流110kV' and (p < 0.34 or p > 0.36)):
                        result = False
                            
            if result == False:
                 err_cells.append((row,35))
                 the_line_is_error = True
                
            #37. 额定电流(低压)(A)
            ## 数值类型
            ## “电压等级”为交流110kV对应额定电流/额定电流(低压)在0.09-0.1之间，以外为不合格数据；
            eddldy = self._getcell(row,36)
            result = True
            if not self._isnumber(eddldy):
                result = False
            elif dydj_valid == True and eddl_valid == True:
                eddldy = float(eddldy)
                p = eddl / eddldy
                if dydj == u'交流110kV' and (p < 0.09 or p > 0.1):
                    result = False
            if result == False:
                err_cells.append((row,36))
                the_line_is_error = True 
                
            #38. 短路阻抗高压－中压(%)
            ## “绕组型式”为双绕组对应/，以外的为不合格数据；
            ## / 或数值型
            result = True
            dlkzgy_zy = self._getcell(row,37)
            if dlkzgy_zy == '/':
                if rzxs_valid == True and rzxs != u'双绕组':
                    result = False
            elif self._isnumber(dlkzgy_zy):
                dlkzgy_zy = float(dlkzgy_zy)
            else:
                result = False
            if result == False:
                err_cells.append((row,37))
                the_line_is_error = True
                
            #39. 短路阻抗高压－低压(%)
            ## 若“短路阻抗高压－中压(%)”不为/的数据应小于对应短路阻抗高压－低压(%)，以外为不合格数据
            # 数值类型
            dlkzgy_dy = self._getcell(row,38)
            result = True
            if not self._isnumber(dlkzgy_dy):
                result = False
            else:
                dlkzgy_dy = float(dlkzgy_dy)
                if dlkzgy_zy != '/' and dlkzgy_zy  >= dlkzgy_dy:
                    result = False
                    
            if result == False:
                err_cells.append((row,38))
                the_line_is_error = True 
                
                
            #40. 短路阻抗中压－低压(%)    
            ## 数值类型
            ## “绕组型式”为双绕组对应/，以外的为不合格数据；
            dlkzzy_dy = self._getcell(row,39)
            result = True
            if dlkzzy_dy == '/':
                if rzxs_valid == True and rzxs != u'双绕组':
                    result = False
            elif self._isnumber(dlkzzy_dy):
                dlkzzy_dy = float(dlkzzy_dy)
            else:
                result = False
                
            if result == False:
                err_cells.append((row,39))
                the_line_is_error = True
            
            #42. 负载损耗(实测值)(满载)(kW)
            ## 数值类型
            ## "1、“电压等级”为交流220kV对应大于350小于750之间，以外为不合格数据；
            ## 2、“电压等级”为交流110kV对应大于80小于等于300之间，以外为不合格数据；
            ## 3、“电压等级”为交流35kV对应大于19小于等于60之间，以外为不合格数据；"
            result = True
            fzsh_valid = True
            fzsh = self._getcell(row,41)
            if  self._isnumber(fzsh):
                fzsh = float(fzsh)
                if dydj_valid == True:
                    if (dydj == u'交流220kV' and (fzsh <= 350 or fzsh >=750)) or \
                       (dydj == u'交流110kV' and (fzsh <=80 or fzsh>=300)) or \
                       (dydj == u'交流35kV' and (fzsh <=19 or fzsh>=60)):
                           result = False
            else:
                result = False
                fzsh_valid = False
                
            if result == False:
                err_cells.append((row,41))
                the_line_is_error = True
                
            #41. 空载损耗(kV)
            ## 数值类型
            ## 大于2/5倍“负载损耗”的数据为不合格数据
            kzsh = self._getcell(row,40)
            result = True
            if self._isnumber(kzsh):
                kzsh = float(kzsh)
                if fzsh_valid == True and kzsh > 0.4 * fzsh:
                    result = False
            else:
                result = False
            
            if result == False:
                err_cells.append((row,40))
                the_line_is_error = True
             
            #43. 自然冷却噪声(dB)
            ## 数值类型
            ## "1、“电压等级”为交流220kV或者交流110kV数对应55-65之间，以外为不合格数据；
            ## 3、“电压等级”为交流35kV对应38-45之间，以外为不合格数据；" 
            zrlqzs = self._getcell(row,42)
            result = True
            if self._isnumber(zrlqzs):
                zrlqzs = float(zrlqzs)
                if dydj_valid == True:
                    if ((dydj == u'交流220kV' or dydj == u'交流110kV') and (zrlqzs <55 or zrlqzs >65)) or \
                        (dydj == u'交流35kV' and (zrlqzs < 38 or zrlqzs >45)):
                            result = False
            else:
                result = False
                
            if result == False:
                err_cells.append((row,42))
                the_line_is_error = True
                
            #46. 油重
            ## 数值类型
            ## "1、“电压等级”为交流220kV对应大于30小于95之间，以外为不合格数据；
            ## 2、“电压等级”为交流110kV对应大于10小于等于30之间，以外为不合格数据；
            ## 3、“电压等级”为交流35kV对应大于1小于等于10之间，以外为不合格数据；"
            yz = self._getcell(row,45)
            result = True
            if self._isnumber(yz):
                yz = float(yz)
                if dydj_valid == True:
                    if (dydj == u'交流220kV'  and (yz <=30 or yz >=95)) or \
                        ( dydj == u'交流110kV' and (yz <=10 or yz>30)) or \
                        (dydj == u'交流35kV' and (yz <=1 or yz >10)):
                            result = False
            else:
                result = False
            
            if result == False:
                err_cells.append((row,45))
                the_line_is_error = True
                    
            #44.总重(T) 
            ## 数值类型
            ## 小于2.5倍“油重”的为不合格数据
            zz = self._getcell(row,43)
            result = True
            if self._isnumber(zz):
                zz = float(zz)
                if zz < 2.5 * yz:
                    result = False
            else:
                result = False
                
            if result == False:
                err_cells.append((row,43))
                the_line_is_error = True
                
                
            #45. 油号
            ## 25以外为不合格数据
            yh = self._getcell(row,44)
            if (not self._isnumber(yh)) or int(yh) != 25:
                err_cells.append((row,44))
                the_line_is_error = True
                
            #47. 油产地
            ## 不包含“新疆克拉玛依石油”的为不合格数据
            if self._getcell(row,46).find(self.pms['47']) < 0:
                err_cells.append((row,46))
                the_line_is_error = True 
    
            #48. SF6气体额定压力(Mpa)
            ## 不为0的为不合格数据
            if self._getcell(row,47) != '0':
                err_cells.append((row,47))
                the_line_is_error = True 
                
            #49. SF6气体报警压力(Mpa)
            ## 不为0的为不合格数据
            if self._getcell(row,48) != '0':
                err_cells.append((row,48))
                the_line_is_error = True  
                
            #51. 最近投运日期
            ## 空白或早于投运日期的为不合格数据
            zjtyrq = self._getcell(row,50)
            result = True
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
                err_cells.append((row,50))
                the_line_is_error = True 
            
            if the_line_is_error == True:
                #print "the line is error line no= %d current total error lines = %d" % (row,total_error_lines)
                total_error_lines += 1
                        
            # end for 一行校验结束
        #highligh error cells
        
         #self.wb.Save()
        #self.excel.Application.Quit()
        fd.close()
        #错误行数，总行数，错误单元列表
        return (total_error_lines,self.nrows,err_cells)