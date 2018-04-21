# -*- coding: utf-8 -*-
"""
Created on Mon Apr 16 17:43:12 2018

@author: 张凡
@mail 1379875051@qq.com
"""

#from selenium import webdriver
import time,random
import tkinter
import xlwt,re
import traceback
#from tkinter import messagebox
import requests
import json
import xlrd
from bs4 import BeautifulSoup
import threading
from email.mime.multipart import MIMEMultipart  
from email.mime.application import MIMEApplication  
from email.mime.text import MIMEText  
import smtplib
import poplib,email
from email.header import decode_header


def accept_email():
    try:
        p = poplib.POP3('pop.126.com')  
        p.user('usb_disk_data@126.com')  
        p.pass_('bh2018')
        ret = p.stat()
    except:
        print('Login failed!')
    st = p.top(ret[0], 0)
    strlist = []
    for x in st[1]:
            try:
                strlist.append(x.decode())
            except:
                try:
                    strlist.append(x.decode('gbk'))
                except:
                    strlist.append((x.decode('big5')))
    mm = email.message_from_string('\n'.join(strlist))

    sub = decode_header(mm['subject'])
    if sub[0][1]:
        submsg = sub[0][0].decode(sub[0][1])
    else:
        submsg = sub[0][0]
    print(submsg.strip())
    p.quit()
    
class send():
    def __init__(self,subject,content):
       self.msg_from='usb_disk_data@126.com'                                 #发送方邮箱
       self.passwd='bh2018'                                   #填入发送方邮箱的授权码
       self.msg_to='1379875051@qq.com'                                  #收件人邮箱
                                
       self.subject=subject                                     #主
       
      
       self.content=content
       self.msg = MIMEText(self.content)
       self.msg['Subject'] = self.subject
       self.msg['From'] = self.msg_from
       self.msg['To'] = self.msg_to
       self.s=smtplib.SMTP_SSL("smtp.qq.com",465)
       self.s.login(self.msg_from, self.passwd)
       self.s.sendmail(self.msg_from, self.msg_to, self.msg.as_string())
       self.s.quit()
       
class com_spyder():
    def __init__(self,p,t=0.5):
        self.p=p
        self.t=t
        self.data=xlrd.open_workbook(self.p)
        self.table1=self.data.sheets()[0]
        self.name= self.table1.col_values(0)
        self.ID= self.table1.col_values(10)
        self.comm=self.table1.col_values(6)
        for i in range(1,len(self.ID)):
            if int(self.comm[i])<1000:
                num=int(int(self.comm[i])/20)
            else:
                num=50
            self.I=self.ID[i]
            self.Name=self.name[i]
            self.Na=''
            for n in self.Name:
                if n not in '*#@！~￥%……&，。‘’；、/+=-？?><《》':
                    self.Na=self.Na+n
            if i%50==1:
                prox= self.get_ip()
            proxy_ip = random.choice(prox)
            self.proxies = {'http': proxy_ip}
            try:
                T=threading.Thread(target=self.get,args=(num,self.I,self.Na,i,self.t,))
                T.start()
                time.sleep(random.randint(10,30))
            except:
                pass
    def get_ip(self):
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
    def get(self,num,I,Na,number,t):
        print('> > >',number,"当前代理IP:",self.proxies)
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
                self.url = 'https://rate.taobao.com/feedRateList.htm?auctionNumId='+I+'&currentPageNum='+str(j)
                head={'user-agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36'}
                r=requests.get(self.url,headers=head,proxies=self.proxies)
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
class spyder():
    def __init__(self,goods,pages,t):    
        self.goods=goods
        self.pages=pages
        self.t=t
        self.url='https://s.taobao.com/'
        self.work=xlwt.Workbook()
        self.sheet =self.work.add_sheet("淘宝")
        style = xlwt.easyxf('font: bold 1, color red;')
        self.sheet.row(0).height = 256 *20
        self.sheet.col(1).width = 256 *10
        self.sheet.col(2).width = 256 *10
        self.sheet.col(3).width = 256 *20
        self.sheet.col(4).width = 256 *10
        self.sheet.col(5).width = 256 *10
        self.sheet.col(6).width = 256 *10
        self.sheet.col(7).width = 256 *20
        self.sheet.col(8).width = 256 *20
        self.sheet.col(9).width = 256 *20
        self.sheet.col(10).width = 256 *20

        self.sheet.write(0,0,"商品名称",style)                      
        self.sheet.write(0,0+1,"价格",style)
        self.sheet.write(0,1+1,"付款人数",style)
        self.sheet.write(0,2+1,"店铺名称",style)
        self.sheet.write(0,3+1,"发货地",style)
        self.sheet.write(0,4+1,"运费",style)
        self.sheet.write(0,5+1,"评论人数",style)
        self.sheet.write(0,6+1,"性价比",style)
        self.sheet.write(0,7+1,"商品图片链接",style)
        self.sheet.write(0,8+1,"商品链接",style)
        self.sheet.write(0,9+1,"商品ID号",style)
           
        self.path="F:\\Temp\\淘宝_"+self.goods+"_商品信息.xls"
    def get_ip(self):
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
    def loop(self):
        try:
            data=self.r.text
            img_pat='"pic_url":"(//.*?)"'
            name_pat='"raw_title":"(.*?)"'
            nick_pat='"nick":"(.*?)"'
            price_pat='"view_price":"(.*?)"'
            fee_pat='"view_fee":"(.*?)"'
            sales_pat='"view_sales":"(.*?)"'
            comment_pat='"comment_count":"(.*?)"'
            city_pat='"item_loc":"(.*?)"'
            goods_link='"nid":"(.*?)"'
            #查找满足匹配规则的内容，并存在列表中
            imgL=re.compile(img_pat).findall(data)
            nameL=re.compile(name_pat).findall(data)
            nickL=re.compile(nick_pat).findall(data)
            priceL=re.compile(price_pat).findall(data)
            feeL=re.compile(fee_pat).findall(data)
            salesL=re.compile(sales_pat).findall(data)
            commentL=re.compile(comment_pat).findall(data)
            cityL=re.compile(city_pat).findall(data)
            goods_linK=re.compile(goods_link).findall(data)
            for j in range(len(salesL)):
                img="http:"+imgL[j]#商品图片链接
                name=nameL[j]#商品名称
                nick=nickL[j]#淘宝店铺名称
                price=priceL[j]#商品价格
                fee=feeL[j]#运费
                sales=salesL[j][:-3]#商品付款人数
                comment=commentL[j]#商品评论数，会存在为空值的情况
                goodsurl="https://detail.tmall.com/item.htm?id="+goods_linK[j]#货物链接
                if(comment==""):
                    comment=0
                city=cityL[j]#店铺所在城市
                xinjiabi=float(sales)/float(price)
                self.n+=1
                self.sheet.write(self.n,0,name)
                self.sheet.write(self.n,0+1,price)
                self.sheet.write(self.n,1+1,sales)
                self.sheet.write(self.n,2+1,nick)
                self.sheet.write(self.n,3+1,city)
                self.sheet.write(self.n,4+1,fee)
                self.sheet.write(self.n,5+1,comment)
                self.sheet.write(self.n,6+1,xinjiabi)
                self.sheet.write(self.n,7+1,img)
                self.sheet.write(self.n,8+1,goodsurl)
                self.sheet.write(self.n,9+1,goods_linK[j])
                #print(name,nick,price,fee,sales,img,city)
        except:
             traceback.print_exc()
             
    def run(self):
        try:
            if self.pages!='':
                self.p=int(self.pages)
            else:
                self.p=100
            self.n=0
            for i in range(self.p):
                if i%50==0:
                    prox= self.get_ip()
                proxy_ip = random.choice(prox)
                proxies = {'http': proxy_ip}
                head={'user-agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36'}
                self.r=requests.get('https://s.taobao.com/search?q='+self.goods+'&s='+str(44*i),headers=head,proxies=proxies)
                self.r.encoding=self.r.apparent_encoding
                self.loop()
                time.sleep(random.randint(1,5))
            self.work.save(self.path)
            #messagebox.showinfo('提示','商品信息爬取已经完成，正在下载评论，请不要关掉该窗口，保存地址在'+self.path)
            com_spyder(self.path,t=self.t)
        except:
            traceback.print_exc()
            self.work.save(self.path)
            print("出现错误或者页面已经爬取完成")
            com_spyder(self.path,t=self.t)


    
r=tkinter.Tk() #tkinter root初始化 
r.title('淘宝商品信息获取  制作者:张凡')#界面标题栏
r.geometry()#界面大小（自适应）

mu=tkinter.Menu(r)
fi=tkinter.Menu(mu,tearoff=False)
fi.add_command(label='开发者：张凡',command='callback')
fi.add_command(label='版本号：1.0.0 ',command='callback')
fi.add_command(label='软件介绍：用于获取淘宝商品信息！',command='callback')
fi.add_command(label='退出',command=r.destroy)
mu.add_cascade(label='关于',menu=fi)
r.config(menu=mu)

tkinter.Label(r,text='淘宝商品信息获取  制作者:张凡').pack()#第一个标签
tkinter.Label(r,text='        ').pack()#第一个标签
tkinter.Label(r,text='请输入你想要爬取的商品名称').pack()#第一个标签
input1=tkinter.StringVar()#捕获用户输入
xen=tkinter.Entry(r,textvariable=input1,width=25)#用户文本输入
input1.set('水杯')#输入框预设值
xen.pack()#使用户输入框生效

tkinter.Label(r,text='请输入爬取页面数（选填）').pack()#第一个标签
input2=tkinter.StringVar()#捕获用户输入
yen=tkinter.Entry(r,textvariable=input2,width=25)#用户文本输入
input2.set('1')#输入框预设值
yen.pack()#使用户输入框生效

tkinter.Label(r,text='请输入延时秒数（选填）').pack()#第一个标签
input3=tkinter.StringVar()#捕获用户输入
zen=tkinter.Entry(r,textvariable=input3,width=25)#用户文本输入
input3.set('')#输入框预设值
zen.pack()#使用户输入框生效
def start():
    goods=input1.get()
    pages=input2.get()
    t=input3.get()
    if t:
        t=float(t)
    else:
        t=0.5
    goods_spyder=spyder(goods=goods,pages=pages,t=t)
    goods_spyder.run()
    tkinter.Label(r,text='下载完成，信息已经保存在F盘中！').pack()#第一个标签

tkinter.Label(r,text='                                                       ').pack()#第一个标签
tkinter.Button(r,text=('开始'),command=start,width=10,height=1,bg='blue').pack()
tkinter.Label(r,text='                                                       ').pack()#第一个标签
r.mainloop()






