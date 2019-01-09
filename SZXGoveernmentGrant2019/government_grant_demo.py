#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019/1/3
# @Author  : nanqiang
import re
import time

import requests
import pandas as pd
import warnings
from urllib.parse import urlencode
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
"""
需求：收集深圳市政府资助过的企业列表
思路：下载网页，解析字段，检索附件，下载保存附件（word和excel两种格式）
"""

class GovernmentGrant():

    def __init__(self):
        self.ua = UserAgent(verify_ssl=False)
        warnings.filterwarnings('ignore')

    def getPageReuslt(self,searchword,page,type):
        '''
        爬取一页的公示信息
        :param searchword: 搜索关键字
        :param page: 第几页
        :return: html
        '''
        host='http://61.144.227.212/was5/web/search?'
        params=urlencode({
            "andsen":"" ,
            "channelid":"203937",
            "classsql":"CTYPE=('%s')"%type,
            "exclude":"",
            "keyword":searchword,
            "orderby":"-DOCRELTIME",
            "orsen":"" ,
            "outlinepage":5,
            "page":page,
            "perpage":10,
            "searchscope":"",
            "searchword":searchword,
            "timescope":"",
            "timescopecolumn":"" ,
            "total":"" ,
        })
        url=host+params
        headers = {
            'User-Agent': self.ua.chrome,
        }
        tf=True
        while tf:
            try:
                response=requests.get(url,headers)
            except requests.exceptions.ChunkedEncodingError:
                print('requests.exceptions.ChunkedEncodingError')
                time.sleep(2)
            except requests.exceptions.ConnectionError:
                print('requests.exceptions.ConnectionError')
                time.sleep(2)
            else:
                tf=False
        status=response.status_code
        if status==200:
            print(url)
            result=response.text
            return result

    def clear_html_re(self,src_html):
        '''
        正则清除HTML标签
        :param src_html:原文本
        :return: 清除后的文本
        '''
        content = re.sub(r"</?(.+?)>", "", src_html)  # 去除标签
        # content = re.sub(r"&nbsp;", "", content)
        content = re.sub(r"\s+", "", content) # 去除空白字符
        return content

    def downloadfile(self,filelink,filename):
        '''
        下载文件
        :param filename: 文件保存名字
        :param fileLink: 文件链接
        '''
        r = requests.get(filelink)
        if 'doc'in filename :
            filepath="E:\\深圳政府资助公示文件\\word\\"
        elif 'xls'in filename:
            filepath="E:\\深圳政府资助公示文件\\excel\\"
        elif 'zip'in filename or 'rar'in filename:
            filepath="E:\\深圳政府资助公示文件\\zip\\"
        elif 'pdf'in filename:
            filepath="E:\\深圳政府资助公示文件\\pdf\\"
        else:
            filepath="E:\\深圳政府资助公示文件\\else"
        with open(filepath+filename, "wb+") as code:
            code.write(r.content)

    def addfilename(self,name,filelink):
        '''
        补全文件名称，加上后缀名
        :param name:
        :param filelink:
        :return:
        '''
        if '.docx' not in name and '.xlsx' not in name and '.doc' not in name and '.xls' not in name and '.pdf' not in name:
            if '.docx' in filelink:
                name = name + '.docx'
            elif '.doc' in filelink:
                name = name + '.doc'
            elif '.xlsx' in filelink:
                name = name + '.xlsx'
            elif '.xls' in filelink:
                name = name + '.xls'
            elif '.pdf' in filelink:
                name = name + '.pdf'
            elif '.rar' in filelink:
                name = name + '.rar'
            elif '.zip' in filelink:
                name = name + '.zip'
        return name

    #解析字段并保存入excel
    def getAnalyseInfo(self,excelfilename,searchword,page,type):
        page_result = self.getPageReuslt(searchword,page,type)
        page_soup=BeautifulSoup(page_result,'lxml')
        dllist=page_soup.select('.sb-left dl')
        if dllist:
            code=0
            oldtitle=""
            for dl in dllist:
                title=dl.select_one('dt a').get_text()#标题
                title_link=dl.select_one('.trt_js_tit4 a').get_text()#原链接
                title_date=dl.select_one('.trt_js_tit4 span').get_text().replace(" ","")#发布日期
                if title_link and oldtitle!=title:#and '企业'in title and '资助'in title and '公示'in title
                    oldtitle=title
                    titleCode="%s%s"%(page,code)
                    print(title,titleCode,title_link)
                    filehost = re.search("(.*)\\/", title_link)[0]
                    headers={'User-Agent':self.ua.chrome}

                    tf=True
                    while tf:
                        try:
                            response=requests.get(title_link,headers)
                        except requests.exceptions.ConnectionError:
                            print('requests.exceptions.ConnectionError')
                            time.sleep(1)
                        else:
                            tf=False

                    response.encoding=response.apparent_encoding
                    result=response.text
                    soup=BeautifulSoup(result,'lxml')
                    #不完全统计正文有三种格式
                    plist=soup.select('.TRS_Editor p')
                    if plist is None or len(plist) < 1:
                        plist=soup.select('.contentWrap p')
                        if plist is None or len(plist)<1:
                            plist=soup.select('div.updatembcss p')

                    ptext=self.clear_html_re(str(plist))
                    government=plist[-2].get_text().replace(" ","").replace("　","") if len(plist)>2 and '深圳'in plist[-2].get_text() else ''
                    release_date = plist[-1].get_text().replace(" ","").replace("　","") if len(plist)>2 else ''

                    #文件下载链接，可直接解析提取，不完全统计，有7种格式
                    filelist=soup.select('.contentWrap ul li a')
                    if filelist is None or len(filelist)<1:
                        filelist=soup.select('.contentWrap div table tr td table tr td a')
                        if filelist is None or len(filelist)<1:
                            filelist=soup.select('.nr li a')
                            if filelist is None or len(filelist)<1:
                                filelist=soup.select('.nr-xgfj li a')
                                if filelist is None or len(filelist) < 1:
                                    filelist = soup.select('.list a')#http://www.szft.gov.cn/bmxx/qrlzyj/rl_zwdt/zwdt_tzgg/201811/t20181123_14678038.htm
                                    if filelist is None or len(filelist)<1:
                                        filelist=soup.select('#appendix a')
                                        if filelist is None or len(filelist) < 1:
                                            filelist = soup.select('.fjdown p a')
                    title_year = re.search('20(.*?)年', title)[0] if re.search('20(.*?)年', title) is not None else 'null'
                    if filelist:
                        code = code + 1
                        for file in filelist:
                            filelink=filehost+file.get('href')[2:]
                            name=file.get_text()
                            name=self.addfilename(name,filelink)
                            if name!="":
                                filename="p"+str(page)+"_"+str(titleCode)+"_"+name
                                # self.downloadfile(filelink,filename)
                                #存入excel
                                df1 = pd.DataFrame(pd.read_excel(excelfilename, sheet_name='Sheet1'))
                                df2=pd.DataFrame({'code':[titleCode],'标题': [title],'年份': [title_year],'发布日期': [title_date],'公示日期': [release_date],
                                                  '附件名称': [filename],'公示部门': [government],'公示内容': [ptext], '公示链接': [title_link]})
                                writer = pd.ExcelWriter(excelfilename)
                                pd.concat([df1,df2]).to_excel(writer, sheet_name='Sheet1',index=False)
                                writer.save()
                                print(filename)
                    # 文件下载链接，script加载，不完全统计，有3种格式
                    else:
                        code = code + 1
                        namekey,linkkey="",""
                        if 'var linkdesc='in result and 'var filedesc'not in result:
                            namekey='linkdesc'
                            linkkey='linkurl'
                        elif 'var isAPPENDIX' in result:
                            namekey = 'name'
                            linkkey = 'isAPPENDIX'
                        elif 'var filedesc' in result:
                            namekey = 'filedesc'
                            linkkey = 'fileurl'
                        if namekey!="" and linkkey!="":
                            filelinkdesclist=re.search('var %s(.*?)";'%namekey,result)[0].replace('var %s'%namekey,'').replace('=','').replace('";','').split(';')
                            filelinklist = re.search('var %s(.*?)";'%linkkey, result)[0].replace('var %s'%linkkey,'').replace('=','').replace('";','').split(';')
                            if filelinkdesclist and filelinklist:
                                for linkdesc,filelink in zip(filelinkdesclist,filelinklist):
                                    name = linkdesc.replace('"', '').replace(' ', '')
                                    filelink = filehost+filelink.replace('"', '').replace(' ', '')
                                    if name != "":
                                        name = self.addfilename(name, filelink)
                                        filename = "p" + str(page) + "_"+str(titleCode)+"_"+ name
                                        # self.downloadfile(filelink, filename)
                                        # 存入excel
                                        df1 = pd.DataFrame(pd.read_excel(excelfilename, sheet_name='Sheet1'))
                                        df2 = pd.DataFrame(
                                            {'标题': [title],'code':[titleCode], '年份': [title_year], '发布日期': [title_date],
                                             '公示日期': [release_date], '附件名称': [filename], '公示部门': [government],
                                             '公示内容': [ptext], '公示链接': [title_link]})
                                        writer = pd.ExcelWriter(excelfilename)
                                        pd.concat([df1,df2]).to_excel(writer, sheet_name='Sheet1', index=False)
                                        writer.save()
                                        print(filename)

        # #下一页
        # next=soup.select_one('.next-page')
        # if next:
        #     nextlink="http://61.144.227.212/was5/web/"+next.get('href')

    # 运行函数
    def run(self):
        excelfilename='深圳市政府资助项目汇总.xls'
        searchword = '企业 资助 公示'
        type='通知公告'
        for page in range(1,210):
            print(page)
            self.getAnalyseInfo(excelfilename,searchword,page,type)
            time.sleep(2)


if __name__ == '__main__':
    governmentgrant=GovernmentGrant()
    governmentgrant.run()
