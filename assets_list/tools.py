import time
from time import sleep

from lxml import etree
import json
import os
import shutil
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as  EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options


class Tools():
    def __init__(self):
        self.root_dir=os.path.dirname(os.path.abspath(__file__))
        self.out=os.path.join(self.root_dir,'out')
        self.config = os.path.join(self.root_dir, 'config')
        self.error = os.path.join(self.root_dir, 'error')
        self.company_excel = os.path.join(self.config, '事务所名单0506.xlsx')
        self.pass_list=['1']
        self.page_len=0
        self.current_page=0

    def make_folder(self):
        with open(os.path.join(self.root_dir,'out','have_write.json'),'r') as f:
            assets_list=json.load(f)
        for item in assets_list:
            dirx=os.path.join(self.root_dir,'assets',item)
            os.makedirs(dirx)
            os.popen('cd {}'.format(dirx))
            uni=os.path.join(dirx,'union.txt')
            print("Will create {}".format(uni))
            with open(uni,'w') as f:
                f.write('')

    def get_have_write(self):
        root_dir = os.path.dirname(os.path.abspath(__file__))
        have_write_file = os.path.join(root_dir, 'out', 'have_write.json')
        have_write_excle= os.path.join(root_dir, 'out', '股东变更.xlsx')

        df=pd.read_excel(have_write_excle)
        aslist=list(set(df.iloc[:,0]))
        print(aslist)
        print(len(aslist))
        str_list = json.dumps(aslist, ensure_ascii=False, indent=4)
        with open(have_write_file, 'w') as f:
            f.write(str_list)


    def get_current_url(self):
        final_name=[]
        company_url={}
        current_url={}
        with open(os.path.join(self.config,'company_url.json'),encoding='utf-8') as f:
            company_url=json.load(f)
        print(company_url)
        df=pd.read_excel(self.company_excel)
        a=df.iloc[0:52,[0,2]]
        old_name=a['[境内审计事务所]']
        new_name=a['现用名']
        indexs=old_name.index
        for i in indexs:
            name1=old_name[i]
            name2=new_name[i]
            name= name2 if not pd.isnull(name2) else name1
            final_name.append(name)
        print(final_name)

        for name in final_name:
            if name == "天健会计师事务所有限公司2":
                continue
            url=company_url[name]
            current_url[name]=url

        obj=json.dumps(current_url,indent=4,ensure_ascii=False)
        with open(os.path.join(self.out,'current_url.json'),'w',encoding='gbk') as f:
            f.write(obj)

    def get_tr(self):
        html=b'<ul class="pagination"> <li><a id="ajaxpage" href="111">&lt;</a></li> <li><a id="ajaxpage" href="222">1</a></li><li class="active"><a htef="#">2</a></li><li><a id="ajaxpage" href="333">3</a></li> <li><a id="ajaxpage" href="444">&gt;</a></li> </ul>'
        prety=etree.fromstring(html)
        name = prety.xpath('//ul[@class="pagination"]/li/a/text()')[0:] #./td[2]/table/tr/td[2]/a/h3/text()
        print(name)

    def myselenium(self):
        browser=webdriver.Chrome()
        try:
            dic1={
                'domain': '.qichacha.com',
                'name':'CNZZDATA1254842228',
                'value':'944474535-1556153151-https%253A%252F%252Fwww.baidu.com%252F%7C1557277335',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic1)
            dic2={
                'domain': '.qichacha.com',
                'name':'Hm_lpvt_3456bee468c83cc63fb5147f119f1075',
                'value':'1557282450',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic2)
            dic3={
                'domain': '.qichacha.com',
                'name':'Hm_lvt_3456bee468c83cc63fb5147f119f1075',
                'value':'1556418827,1556592894,1557111516,1557211763',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic3)
            dic4={
                'domain': '.qichacha.com',
                'name':'QCCSESSID',
                'value':'k3cv99grhs7arn2koq8pi45q63',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic4)
            dic5={
                'domain': '.qichacha.com',
                'name':'UM_distinctid',
                'value':'16a520daba55-0d6121a68dfda5-e323069-1fa400-16a520daba71e4',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic5)
            dic6={
                'domain': '.qichacha.com',
                'name':'_uab_collina',
                'value':'155615479301622472480909',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic6)
            dic7={
                'domain': '.qichacha.com',
                'name':'acw_tc',
                'value':'42c6b24515561547932928506ee7f9feddcedff91d78e9659c9128002e',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic7)
            dic8={
                'domain': '.qichacha.com',
                'name':'hasShow',
                'value':'1',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic8)
            dic9={
                'domain': '.qichacha.com',
                'name':'zg_de1d1a35bfa24ce29bbf2c7eb17e6c4f',
                'value':'%7B%22sid%22%3A%201557282373033%2C%22updated%22%3A%201557282448826%2C%22info%22%3A%201557111515503%2C%22superProperty%22%3A%20%22%7B%7D%22%2C%22platform%22%3A%20%22%7B%7D%22%2C%22utm%22%3A%20%22%7B%7D%22%2C%22referrerDomain%22%3A%20%22www.qichacha.com%22%2C%22cuid%22%3A%20%224f7112a3d7c45eb79a3965e430e5943b%22%7D',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic9)
            dic10={
                'domain': '.qichacha.com',
                'name':'zg_did',
                'value':'%7B%22did%22%3A%20%2216a520dc94d2d8-056260113d68e8-e323069-1fa400-16a520dc94e6e6%22%7D',
                'expires': None,
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'secure': False,
            }
            browser.add_cookie(dic10)
            browser.get('https://www.qichacha.com/firm_84bca9d6d559d89882a0721b5f91e6d9.html')
            browser.implicitly_wait(20)
            size=browser.get_window_size()
            print(size)
            # browser.get_screenshot_as_file('selenium.png')
            page=browser.find_element(by=By.ID,value='ajaxpage')
            print(page)
            print(type(page))
            # print(input)
            # input.send_keys('Python')
            # input.send_keys(Keys.ENTER)
            # wait=WebDriverWait(browser,10)
            # wait.until(EC.presence_of_element_located((By.ID,'content_left')))
            print(browser.current_url)
            print(browser.get_cookies())
        except Exception as e:
            print(e)

    def driver_celcle(self,driver):
        driver=driver
        driver.get('https://www.qichacha.com/firm_84bca9d6d559d89882a0721b5f91e6d9.html')
        driver.implicitly_wait(20)
        size = driver.get_window_size()
        print(size)
        page=[]
        pa = driver.find_elements_by_xpath('//*[@id="ajaxpage"]')
        for p in pa:
            if p.text.isdigit():
                page.append(p)
        page_lan=len(page)+1
        pass_len=len(self.pass_list)
        print('pass_len={}'.format(pass_len))
        print("page_lan={}".format(page_lan))
        if pass_len < page_lan:
            print('In cycle!')
            for p in page:
                print(p.id)
                page_num=p.text
                if page_num not in self.pass_list:
                    print(p.id)
                    print(p.tag_name)
                    print(p.text)
                    p.click()
                    # html=driver.page_source
                    # print(html)
                    time.sleep(1)
                    self.pass_list.append(page_num)
                    self.driver_celcle(driver)
        else:
            print("Page done !")
            return True

    def myselenium2(self):
        chrome_option=Options()
        chrome_option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        executable_path="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        os.environ["webdriver.chrome.driver"]=r'C:\Users\Administrator\AppData\Local\Google\Chrome\Application'
        driver=webdriver.Chrome(chrome_options=chrome_option)
        print(driver.title)
        # self.driver_celcle(driver)
        driver.get('https://www.qichacha.com/firm_84bca9d6d559d89882a0721b5f91e6d9.html')
        driver.implicitly_wait(20)
        size = driver.get_window_size()
        print(size)
        page=[]
        pa = driver.find_elements_by_xpath('//*[@id="ajaxpage"]')
        for p in pa:
            if p.text.isdigit():
                page.append(p)
        self.page_len=len(page)
        print("page length {}".format(self.page_len))

        pa=page[self.current_page]
        tx=pa.text
        pa.click()
        sleep(2)
        new_pa = driver.find_elements_by_xpath('//*[@id="partnerslist"]/div[2]/nav/ul//li/a')
        for new_p in new_pa:
            if new_p.text == tx:
                print(new_p.id)
                print(new_p.tag_name)
                print(new_p.text)
                print(type(driver.page_source))
        self.current_page=self.current_page+1
        if self.current_page > self.page_len-1:
            return True
        else:
            self.myselenium2()


# def test():
#     root_dir = os.path.dirname(os.path.abspath(__file__))
#     flag = os.path.join(root_dir, "run")
#     with open(flag, "w") as f:
#         f.write("xx")

if __name__=="__main__":
    # make_folder()
    # get_have_write()
    tool=Tools()
    # tool.get_current_url()
    # tool.get_have_write()
    # tool.get_tr()
    # tool.myselenium2()
    tool.make_folder()