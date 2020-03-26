#!/usr/bin/env python
# coding: utf-8

# In[5]:


import requests
import json
import pandas as pd
import numpy as np
from pandas.io.json import json_normalize
from urllib.request import urlopen
import os
import webbrowser
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from bs4 import BeautifulSoup
from xbrl import XBRLParser, GAAP, GAAPSerializer
from html_table_parser import parser_functions as parser
from xml.etree.ElementTree import parse
import tarfile
import time
import zipfile
from pathlib import Path
path_to_download_folder = str(os.path.join(Path.home(), "Downloads"))


# In[6]:


# API 호출
apikey='************************************'
crpcode = "00184667"
type(apikey), len(apikey)
company_list = ['A기업','B기업', 'C기업']
code_list=['00184667','00117337', '00293886']


# In[7]:


def assignment(crpcode):
    global df, table
    url_company = "https://opendart.fss.or.kr/api/list.json?crtfc_key={0}&corp_code={1}&bgn_de=20160101&end_de=20191231&pblntf_ty=A&pblntf_detail_ty=A002&page_no=1&page_count=10"
    url = url_company.format(apikey, crpcode)
    
    
    response = requests.get(url)
    output = json.loads(response.content)
    output_df=json_normalize(output['list'])
    company_code = output_df[output_df['report_nm']=='사업보고서 (2018.12)']['rcept_no'].iloc[0]
    
    url_parser = "https://opendart.fss.or.kr/api/document.xml?crtfc_key={0}&rcept_no="+company_code
    url = url_parser.format(apikey)
    
    webbrowser.open(url)
    time.sleep(3) #다운로드 시간 고려
    os.rename(path_to_download_folder+'/document.xml', path_to_download_folder+'/'+company_code+'.zip')
    
    os.chdir(path_to_download_folder)
    ex_zip=zipfile.ZipFile(company_code+'.zip')
    ex_zip.extractall()
    ex_zip.close()
    
    soup = BeautifulSoup(open((path_to_download_folder+'/'+company_code+'.xml'), 'rb'), 'html.parser')
    body=soup.find("body")
    table=body.find_all('table')
    
    for i in range(len(table)):
        a=pd.DataFrame(parser.make2d(table[i]))
    
        if a.iloc[0,0]=='재무상태표':
            df=pd.DataFrame(parser.make2d(table[i+1]))
            break


    df.columns = df.iloc[0]
    df=df.set_index(df.iloc[:,0])
    df=df.drop(df.index[0])
    df=df.drop(df.columns[0], axis=1)


# In[4]:


try:
    os.remove('dart_output.xlsx')
except FileNotFoundError:
    pass

wb = Workbook()
wb.active
wb.save('dart_output.xlsx')

writer = pd.ExcelWriter('dart_output.xlsx',engine="openpyxl", mode="w")

for i in range(len(code_list)):
    try:
        assignment(code_list[i])
        df.to_excel(writer, sheet_name=company_list[i])
    except AttributeError:
        print(company_list[i] + '의 데이터가 존재하지 않습니다.')
        
writer.save()        

wb=openpyxl.load_workbook('dart_output.xlsx')
wb.save

