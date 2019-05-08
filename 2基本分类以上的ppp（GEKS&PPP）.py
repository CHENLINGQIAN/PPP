# -*- coding: utf-8 -*-
"""
Created on Tue Oct 30 15:03:11 2018

@author: 陈玲倩
"""
#上接1-2基本分类ppp（CPD法）
#本篇代码目的是求出基本分类以上的ppp以及GEKS 
import pandas as pd
import numpy as np
from glob import glob

P = ('C','D','E','F','G')
for i in P:
    paths = glob(i+':\\**\\Rppp', recursive=True)
    if paths!= []:
        path = ''.join(paths)
        break
    else:
        pass
#path = 'C:\\Users\\Administrator\\Desktop\\Rppp\\'##文件夹存放的路径    
bbb1 = pd.read_excel(path+'\\CPD法\\小类ppp.xlsx')
####改行名列名#####
x = np.array([1,2,3,4,5,6,7,8,9,10,11,12])
bbb1['规格']= tuple(np.repeat(x, [29,4,5,5,10,6,12,2,10,1,2,4], axis=0))
bbb1.columns=['大类','上饶','九江','吉安','宜春','抚州','新余','景德镇','萍乡','赣州','鹰潭','南昌']

#读入权重表中的中类权重
temp_p = bbb1
temp_w = pd.read_excel(path + '\\各地区权重.xlsx', sheet_name = '中类权重')

#最终ppp输出结果为11*11的表，并且对角线为1，GEKS为1*11的表格
ppp = []
gg = []   
temp_p = np.array(temp_p)
temp_w = np.array(temp_w)
temp_pw = temp_p[:,1:] * temp_w[:,1:]
sum_pw = np.sum(temp_pw,axis = 0)   
#这两层循环是所求城市相对于剩下10个城市的费雪理想双边价格指数
for j in range(11):
    priceaa = []     
    for k in range(11):
        temp_fm1 = temp_p[:,k+1] * temp_w[:,j+1]
        temp_sum = np.sum(temp_fm1)
        temp_fz1 = temp_w[:,k+1] * temp_p[:,j+1]
        temp_sum2 = np.sum(temp_fz1)
        city_fz = sum_pw[j]
        city_fm2 = sum_pw[k]
        citya = pow(city_fz / temp_sum, 1/2)
        cityb = pow(temp_sum2 / city_fm2, 1/2)
        pricea = citya * cityb
        priceaa.append(pricea)
    #计算geks
    geks = pow(np.prod(priceaa),1/11)
    gg.append(geks)
    #将每次循环输出的ppp合并
    ppp.append(priceaa)
gge = np.array(gg)
gge = pd.DataFrame(gge)

pppm = np.array(ppp)
pppm = pd.DataFrame(pppm)
pppm.insert(0,'地区',(['上饶','九江','吉安','宜春','抚州','新余','景德镇','萍乡','赣州','鹰潭','南昌']))
pppm.columns=['地区','上饶','九江','吉安','宜春','抚州','新余','景德镇','萍乡','赣州','鹰潭','南昌']
#将结果ppp和geks存入表格并输出
pppm.to_excel(path+'最终ppp.xlsx',sheet_name='Sheet 1',index = False)
gge = gge.T
gge.columns=['上饶','九江','吉安','宜春','抚州','新余','景德镇','萍乡','赣州','鹰潭','南昌']
gge.insert(0,'地区',('GEKS'))
gge.to_excel(path+'GEKS.xlsx',sheet_name='Sheet 1',index = False)
    
