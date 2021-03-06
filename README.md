一。概况：
此项目主要爬取企查查网站公司的 股东信息和变更信息。然后根据爬取的数据统计出一个公司的股东变更情况写入Excel文件。程序会记录爬取过的公司，爬取出错的公司，解析出错的公司，没有股东信息的公司，有变更信息的公司，公司的成立时间。将这些信息保存到json文件，以便分析。其中有变更信息的公司和公司成立时间两项数据会用在变更信息的处理。

二。技术要点
 1.反爬策略：
 
  1）程序使用User-Agent库，每次请求随机使用一个User-Agent。
  
  2）程序使用浏览器登陆后的cookies。需要首先手动登陆企查查网站，然后把cookies复制下来
  
  3）请求时要使用随机间隔的时间。
  
 2.处理变更信息第一条变更记录时，如果没有明确入股时间，则入股时间等于公司成立时间，而此次变更时间等于持股时间。处理如下：
 
    if '/'.join(time.split('-')) == self.page_change_time[column_name][0]:
      union_list.append((name, self.build_dict[column_name], amount, '入股', column_name))
      union_list.append((name, time, amount, '持股', column_name))
  请求头如下所示:
            self.headers = {
            'Cookie': 'xxx',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103                                    Safari/537.36',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language':'zh-CN,zh;q=0.9',
            'Accept':'text/html, */*; q=0.01',
            'Referer': 'https://www.qichacha.com/',
            'X-Requested-With': 'XMLHttpRequest',
            'Connection': 'keep-alive',
            'Content-Type': 'application/x-www-form-urlencoded'
            }
      2.请求时要使用随机间隔的时间
      
三。主要数据结构：

1）最终结果是如下的Excel条目
  所在事务所	                       成立时间	    姓名	出资日期/缴付期限	出资额	        撤资日期	   变更日期
  中瑞岳华会计师事务所(特殊普通合伙)	2011/02/18	丁勇	2011/02/18	    货币50万人民币	2013/04/27	2011/02/18;2012/02/09;2012/03/23

2）使用shareholders_change_info和shareholders_info数据结构以时间为线索记录一个人在公司里面轨迹。其中shareholders_info是股东信息里面的数据，
   shareholders_change_info是变更信息里面的数据
   
{
    "('刘琦祺', '中瑞岳华会计师事务所(特殊普通合伙)')": [
        [
            "2011/02/18",
            "50",
            "入股"
        ],
        [
            "2012-02-09",
            "50",
            "持股"
        ],
        [
            "2012-03-23",
            "50",
            "持股"
        ],
        [
            "2012-12-21",
            "50",
            "持股"
        ],
        [
            "2013-01-29",
            "50",
            "持股"
        ],
        [
            "2013-04-27",
            "50",
            "持股"
        ]
    ],
}

3）使用company_change_time数据结构记录股东变更和网页记录形成的变更时间，数据记录在 【变更日期】一列：

  {
        "中瑞岳华会计师事务所(特殊普通合伙)": [
        "2011/02/18",
        "2012/02/09",
        "2012/03/23",
        "2012/12/21",
        "2013/01/29",
        "2013/04/27"
    ],
  }
  
4）使用page_change_time记录网页上记录面的变更时间。此数据结构记录的是网页记录的时间。people_change_time跟此数据比较会得出更准确的入股和撤资时间。

{
      "中瑞岳华会计师事务所(特殊普通合伙)": [
        "2012/02/09",
        "2012/03/23",
        "2012/12/21",
        "2013/01/29",
    ],
}

5）使用people_change_time数据结构记录一个人在一个公司的变更时间，跟page_change_time比较推断出一个人的入股时间和撤资时间。

 {
      "王需如": [
        "2012/02/09",
        "2012/03/23",
        "2012/12/21",
        "2013/01/29",
        "2013/04/27"
    ],
 }
 
6）build_dict 记录的是公司成立的时间，如下所示：

  {
    "立信会计师事务所有限公司": "2000/06/26",
  }
  
三。程序流程：

1.此步骤的类在 GetCompanyUrl.GetCompanyUrl。根据公司名字构建 requests.get()请求的url:https://www.qichacha.com/search?        key=%E4%BF%A1%E6%B0%B8%E4%B8%AD%E5%92%8C%E4%BC%9A%E8%AE%A1%E5%B8%88%E4%BA%8B%E5%8A%A1%E6%89%80%E6%9C%89%E9%99%90%E8%B4%A3%E4%BB%BB%E5%85%AC%E5%8F%B8.

根据返回的页面解析出公司的详细URL，如："立信会计师事务所有限公司": "https://www.qichacha.com/firm_ee05fd1d1523b8bfefe5522997d7e86f.html"

2.此步骤的类在shareholders_info.MyRequests.此类处理current_url.json文件里面的公司链接。current_url.json在Tool.get_current_url()函数生成。

  使用命令：chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenum\AutomationProfile"打开chrome浏览器。此浏览器会在9222端口监听   selenium的http请求。在此浏览器上面登陆。这样就可以访问需要登陆才能看到的信息。
  
  程序会首先检查 股东信息 有多少页。如果是单页就使用步骤一得到的url抓取详细信息。如果是多页，第一页处理方法同上，后面的页使用selenium翻页处理。
  此步骤会生成shareholders_info.son的文件供后面使用。
  
3.此步骤的类在changes_info.MakeTransfer.由于公司的变更信息不规则，所以需要把变更信息手动复制下来。手动处理一下，然后再用此程序处理。此步骤会生成

  shareholders_change_info文件，此文件跟上一步的shareholders_info.son merge后生成合并的信息。程序在根据此信息生成最终数据。
  
  1）Tools.make_folder根据have_write.json生成需要手动处理的目录和union.txt。
  
  2）复制变更信息到union.txt，手动整理一下数据。
  
  3）运行程序生成最终结果。
  


