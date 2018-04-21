# -*- coding: utf-8 -*-
"""
Created on Tue Apr 17 12:41:49 2018

@author: Administrator
"""
# -*- coding: utf-8 -*-
import xlrd
import pylab
import matplotlib.pyplot as plt
import re,jieba

zh=re.compile(u'[\u4e00-\u9fa5]+')
plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号
data=xlrd.open_workbook("【2件装】班尼路短袖t恤男士 2018新款夏季衣服纯色t恤打底半袖_评论信息.xls")
table=data.sheets()[0]
col= table.col_values(3)
text=''
for i in col:
    text=text+i

def sort(text,n):
    word_num=[]
    word_fre=[]
    key_list=[]
    value_list=[]
    seg_list = jieba.cut_for_search(text)
    splited_string=("  ".join(seg_list)).split("  ")
    b=[]
    for i in splited_string:
        if i not in '，。！。. ，,;。 ～ &&~# 。 ！—— ;‘”’“？ ? hellip；、|=+-&' :
            b.append(i)
    string={}
    for aa in b:
        if aa in string:
            string[aa]+=1
        else:
            string[aa]=1
    for value,key in string.items():
        key_list.append(key)
        value_list.append(value)
    key=sorted(key_list,reverse=True)
    for n in range(n):
        try:
            for k,v in string.items():
                if (v==key[n]) and (k not in word_fre) and len(k)>1:
                    word_fre.append(k)
                    word_num.append(key[n])
                else:
                    pass
        except:
            pass
    return word_fre,word_num
def bar(word,word_fre,xlabel,ylabel,title,weizhi):
    pylab.figure(figsize=(9,9))
    #pylab.pie(single_word_fre,labels=single_word,labeldistance = 1.1,autopct = '%3.1f%%',shadow = False,startangle = 90,pctdistance = 0.6)
    pylab.bar(range(len(word)),word_fre,color='green',tick_label=word)
    pylab.xlabel(xlabel,fontproperties='SimHei',fontsize=20)
    pylab.ylabel(ylabel,fontproperties='SimHei',fontsize=20)
    pylab.title(title,fontproperties='SimHei',fontsize=20)
    plt.xticks(rotation=90)
    pylab.savefig(weizhi)

def pie(word,word_fre,title,weizhi):
    pylab.figure(figsize=(9,9))
    pylab.pie(word_fre,labels=word,labeldistance = 1.1,autopct = '%3.1f%%',shadow = False,startangle = 90,pctdistance = 0.6)
    pylab.title(title,fontproperties='SimHei',fontsize=20)
    pylab.savefig(weizhi)
    pylab.legend()   #是plot函数中的label标签生效，下同，不再赘述
word1,word_fre1=sort(text,n=60)
bar(word1,word_fre1,'','出现次数',"商务双层玻璃杯"+"商品评论中出现频率最高的词",'F:\\'+'bar2.png')
pie(word1,word_fre1,"商务双层玻璃杯"+'评论中出现频率最高的词','F:\\'+'pie.png',)
   
