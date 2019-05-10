import json
import os
import random
import re
import sys
import time
import urllib
from urllib.parse import urljoin

import chardet
import pandas as pd
import requests
from lxml import etree
from pydispatch import dispatcher
from scrapy import signals
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


class GetCompanyName():
    def __init__(self):
        self.root_dir=os.path.dirname(os.path.abspath(__file__))
        self.compane_excel=os.path.join(self.root_dir,'config','数据库的事务所名单.xlsx')
        self.company_list = os.path.join(self.root_dir, 'out', 'company_list.json')

    def read_excel(self):
        df= pd.read_excel(self.compane_excel)
        # print(df[['[境内审计事务所]','现用名']]) ['[境内审计事务所]','现用名']
        list1=df['[境内审计事务所]']
        list2=df['现用名']
        list3=df['链接']
        meger_dict={}
        tmp_list=[]
        tmp_list2 = []
        type1={}
        type2={}
        final_list=[]
        for i in range(len(list2)):
            value1=list1[i]
            value2 = list2[i]
            if not pd.isnull(value2):
                print(value2)
                value1=value2
            tmp_list.append(value1)
        for i in range(len(tmp_list)):
            key=tmp_list[i]
            value=list3[i]
            if not pd.isnull(value):
                tmp_list2.append((key,value))
            else:
                tmp_list2.append((key,''))
        for item in tmp_list2:
            name=item[0]
            url=item[1]
            if name not in meger_dict.keys():
                if '特殊普通合伙' in item:
                    type1[name]=url
                else:
                    type2[name]=url
        print('len(type1)={}'.format(len(type1)))
        meger_dict=dict(type1,**type2)
        print('len(meger_dict)={}'.format(len(meger_dict)))
        str_dict=json.dumps(meger_dict,ensure_ascii=False,indent=4)
        with open(self.company_list,'w',encoding='utf-8') as f:
            f.write(str_dict)

class TransCookie:
    def __init__(self, cookie):
        #default_cookie="QCCSESSID=bfnnhdutc69qhrstdfmmupaht4; UM_distinctid=169e0fe574e151-05dfcafcc67344-5f1d3a17-1fa400-169e0fe574f3b6; hasShow=1; _uab_collina=155425796442358186443062; acw_tc=42e7ef4515542579647657088ea2bdba5a9b2910813157eefcf1522efd; zg_did=%7B%22did%22%3A%20%22169e0fe98ff781-00a3d015086695-5f1d3a17-1fa400-169e0fe9900651%22%7D; Hm_lvt_3456bee468c83cc63fb5147f119f1075=1554257978; CNZZDATA1254842228=1036141894-1554255455-%7C1554277139; Hm_lpvt_3456bee468c83cc63fb5147f119f1075=1554277502; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201554277452264%2C%22updated%22%3A%201554277561922%2C%22info%22%3A%201554257975581%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.qichacha.com%22%2C%22cuid%22%3A%20%22018e7692b98477813ce17a33dfff9dd4%22%7D"
        default_cookie='QCCSESSID=roi3n42hkhhm7jmca9ask6b5e0; zg_did=%7B%22did%22%3A%20%2216a49f2f11c1df-0fc5f0cb94f1fe-39395704-1fa400-16a49f2f11d350%22%7D; UM_distinctid=16a49f2f23a1f-0e42364c9f1087-39395704-1fa400-16a49f2f23b23c; CNZZDATA1254842228=1527141422-1556014374-https%253A%252F%252Fwww.baidu.com%252F%7C1556014374; hasShow=1; Hm_lvt_3456bee468c83cc63fb5147f119f1075=1555760520,1555816297,1555932693,1556018820; acw_tc=6f48649615560186847108698e6b75759f0c13fb26d1ab74c600380353; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201556018819361%2C%22updated%22%3A%201556019109740%2C%22info%22%3A%201556018819364%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.baidu.com%22%2C%22cuid%22%3A%20%22018e7692b98477813ce17a33dfff9dd4%22%7D; Hm_lpvt_3456bee468c83cc63fb5147f119f1075=1556019110'
        if not default_cookie:
            self.cookie = cookie
        else:
            self.cookie=default_cookie

    def stringToDict(self):
        '''
        将从浏览器上Copy来的cookie字符串转化为Scrapy能使用的Dict
        :return:
        '''
        itemDict = {}
        items = self.cookie.split(';')
        for item in items:
            key = item.split('=')[0].replace(' ', '')
            value = item.split('=')[1]
            itemDict[key] = value
        return itemDict

class GetCompanyUrl(GetCompanyName):
    def __init__(self):
        super(GetCompanyUrl,self).__init__()
        self.url='https://www.qichacha.com/'
        self.headers = {
            'Cookie': 'QCCSESSID=5b3n6ob3pgbmh8aohj32be8ab4; zg_did=%7B%22did%22%3A%20%2216a4f2bafc1372-0a944ab8a707d9-39395704-1fa400-16a4f2bafc22be%22%7D; UM_distinctid=16a4f2bb18f1a8-0b089e5e5788d5-39395704-1fa400-16a4f2bb1902b1; CNZZDATA1254842228=2119061671-1556105807-https%253A%252F%252Fwww.baidu.com%252F%7C1556105807; Hm_lvt_3456bee468c83cc63fb5147f119f1075=1556018820,1556023586,1556023927,1556106425; hasShow=1; acw_tc=da5bdda815561062900722812e74788704ee6135b1fc6db964ab3f9006; zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f=%7B%22sid%22%3A%201556106424272%2C%22updated%22%3A%201556108810524%2C%22info%22%3A%201556106424284%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.baidu.com%22%2C%22cuid%22%3A%20%22018e7692b98477813ce17a33dfff9dd4%22%7D; Hm_lpvt_3456bee468c83cc63fb5147f119f1075=1556108811',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language':'zh-CN,zh;q=0.9',
            'Accept':'text/html, */*; q=0.01',
            'Referer': 'https://www.qichacha.com/',
            'X-Requested-With': 'XMLHttpRequest',
            'Connection': 'keep-alive',
            'Content-Type': 'application/x-www-form-urlencoded',
            }
        self.company_url = os.path.join(self.root_dir, 'out', 'company_url.json')
        self.compane_dict_new={}
        with open(os.path.join(self.root_dir,'config','user_agent.json'),'r') as f:
            self.agent_list=json.load(f)
        # tc=TransCookie('None')
        # self.cookies_dict=tc.stringToDict()
        # print('cookies_dict={}'.format(self.cookies_dict))
        with open(self.compane_list,'r',encoding='utf-8') as f:
            self.compane_dict=json.load(f)
        print('compane_dict={}'.format(self.compane_dict))

    @staticmethod
    def get(url,headers={},timeout=60,proxies=None,params={}):
        time.sleep(2)
        # s=requests.session()
        print(params)
        try:
            r=requests.get(url=url,headers=headers,timeout=timeout,proxies=proxies,params=params)
            print(chardet.detect(r.content))
            print(type(r.content))
            print(r.headers)
            encoding=chardet.detect(r.content)['encoding']
            r.encoding=encoding

            if r.status_code==200:
                if "text/html" in r.headers['Content-Type']:
                    print('"text/html" in r.headers,return r.text')
                    return r.text
                elif "application/json;" in r.headers['Content-Type']:
                    print('"application/json;" in r.headers,return r.text')
                    return r.json()
            else:
                print('get reponse false!')
                print(r.status_code)
                print(r.reason)
                return None

        except Exception as e:
            print(e)
            return None
    def process_parse(self,response,key):
        # 这部分是股东信息
        # 股东的入姿
        res=response
        company=key
        print('In process_parse')
        print('will get url of {}'.format(company))
        final_url=''
        try:
            tbody = etree.HTML(res).xpath('//*[@id="search-result"]')[0]
            print('tbody={}'.format(tbody))
            if tbody is not None:
                print("Find tbody")
                url=tbody.xpath('./tr[1]/td[3]/a/@href')[0]
                print('get url = {}'.format(url))
                final_url=urljoin(self.url,url)
                print('final_url = {}'.format(final_url))
        except Exception as  e:
            print(e)
            with open(os.path.join(self.root_dir,'error',key+".html"),'w',encoding='utf-8') as f:
                f.write(response)
            print('get url of {} wrong,will wxit!'.format(company))
            sys.exit()
        finally:
            self.compane_dict_new[company]=final_url
            return  final_url


    def process_request(self):
        print("Will process_request")
        base_url='https://www.qichacha.com/search?key='
        for key, value in self.compane_dict.items():
            url=''
            if not value:
                # url=base_url+urllib.parse.quote(key)
                url=base_url+key
                print("Will get url {}".format(url))
                response=''
                try:
                    self.headers['User-Agent']=random.choice(self.agent_list)
                    print("headers['User-Agent']={}".format(self.headers['User-Agent']))
                    response=self.get(url,headers=self.headers)
                except Exception as e:
                    print("Exception = {}".format(e))
                finally:
                    res_url=self.process_parse(response,key)
                    with open(os.path.join(self.root_dir,'out','activate_compane_url.txt'), 'a') as f:
                        f.write(key+"="+res_url+'\n')

    def write_company_url(self):
        url_str=json.dumps(self.compane_dict_new,ensure_ascii=False,indent=4)
        with open(self.company_url,'a',encoding='utf-8') as f:
            f.write(url_str)

if __name__ == "__main__":
    gc=GetCompanyName()
    #gc.read_excel()
    gcu=GetCompanyUrl()
    start=time.time()
    gcu.process_request()
    gcu.write_company_url()
    end=time.time()
    use_time=(end-start)/60
    print("Totle use {} minutes".format(use_time))

