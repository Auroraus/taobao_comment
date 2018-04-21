# -*- coding: utf-8 -*-
"""
Created on Mon Apr 16 17:43:12 2018

@author: 张凡
@mail 1379875051@qq.com
"""
import requests
import json
import xlrd
import time,xlwt
from bs4 import BeautifulSoup
import random
import threading
import traceback
def get_ip():
    url = 'https://www.kuaidaili.com/free/inha/'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36'
    }
    web_data = requests.get(url, headers=headers)
    soup = BeautifulSoup(web_data.text, 'lxml')
    ips = soup.find_all('tr')
    ip_list = []
    for i in range(1, len(ips)):
        ip_info = ips[i]
        tds = ip_info.find_all('td')
        ip_list.append(tds[0].text + ':' + tds[1].text)
    proxy_list = []
    for ip in ip_list:
        proxy_list.append('http://' + ip)
    return proxy_list
def spyder(p,t=0.5):
    data=xlrd.open_workbook(p)
    table1=data.sheets()[0]
    name= table1.col_values(0)
    ID= table1.col_values(10)
    comm=table1.col_values(6)
    for i in range(1,len(ID)):
        if int(comm[i])<1000:
            num=int(int(comm[i])/20)
        else:
            num=50
        I=ID[i]
        Name=name[i]
        Na=''
        for n in Name:
            if n not in '*#@！~￥%……&，。‘’；、/+=-？?><《》':
                Na=Na+n
        global proxies
        if i%50==1:
            prox= get_ip()
        proxy_ip = random.choice(prox)
        proxies = {'http': proxy_ip}
        try:
            T=threading.Thread(target=get,args=(num,I,Na,i,t,))
            T.start()
            time.sleep(random.randint(10,30))
        except:
            pass
def get(num,I,Na,number,t):
        print('> > >',number,"当前代理IP:",proxies)
        work=xlwt.Workbook()
        sheet =work.add_sheet("淘宝")
        style = xlwt.easyxf('font: bold 1, color red;')
        sheet.row(0).height = 256 *20
        sheet.row(1).height = 256 *20
        sheet.col(0).width = 256 *20
        sheet.col(1).width = 256 *20
        sheet.col(2).width = 256 *20
        sheet.col(3).width = 256 *100
        sheet.write(0,0,Na,style)                      
        sheet.write(0,1,I,style)

        sheet.write(1,0,"购买者昵称",style)                      
        sheet.write(1,1,"评论日期",style)
        sheet.write(1,2,"商品类型",style)
        sheet.write(1,3,"对商品的评论",style)
        n=1
        start=time.clock()
        for j in range(2,num):
            try:
                time.sleep(t)
                url = 'https://rate.taobao.com/feedRateList.htm?auctionNumId='+I+'&currentPageNum='+str(j)
                head={'user-agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36'}
                r=requests.get(url,headers=head,proxies=proxies)
                r.encoding=r.apparent_encoding
                jc = json.loads(r.text.strip().strip('()'))
                for m in range(20):
                    n=n+1
                    sheet.write(n,0,jc['comments'][m]['user']['nick'])
                    sheet.write(n,1,jc['comments'][m]['date'])
                    sheet.write(n,2,jc['comments'][m]['auction']['sku'])
                    sheet.write(n,3,jc['comments'][m]['content'])
            except:
                traceback.print_exc
                pass
        end=time.clock()
        path="F:\\Temp\\淘宝_"+Na+"_评论信息.xls"
        print(number,"下载完成，用时",end-start)
        work.save(path)


