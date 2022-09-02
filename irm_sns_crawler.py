import requests
from bs4 import BeautifulSoup as bs
import pandas as pd
import os
import time


try:
    df_irm = pd.read_excel('./irm_data.xlsx')
    
except:
    df_irm = pd.DataFrame()
    column_lst = ['authorName','industry','stockCode','coname','question','reply']
    for item in column_lst:
        df_irm[item] = []
    
page_number = round(len(df_irm)/10+1)

url = 'http://irm.cninfo.com.cn/ircs/index/search'
while True:
    data = {
    'pageNo': page_number,
    'pageSize': 10,
    'searchTypes': '1,11,'}
    try:
        r = requests.post(url,data)
    except:
        time.sleep(10)
        try:
            r = requests.post(url,data)
        except:
            break
    results = r.json()['results']
    if len(results) == 0:
        break
    row_index = len(df_irm)
    for i in range(len(results)):
        df_irm.loc[row_index,'industry'] = results[i]['trade']
        df_irm.loc[row_index,'stockCode'] = results[i]['stockCode']
        df_irm.loc[row_index,'coname'] = results[i]['companyShortName']
        df_irm.loc[row_index,'question'] = results[i]['mainContent']
        df_irm.loc[row_index,'authorName'] = results[i]['authorName']
        try:
            df_irm.loc[row_index,'reply'] = results[i]['attachedContent']
        except:
            df_irm.loc[row_index,'reply'] = 'None'
        row_index += 1
        
    
    if page_number%100 == 0:
        df_irm = df_irm[['authorName','stockCode','coname','question','reply']]    
        df_irm.to_excel('./irm_data.xlsx',encoding = 'utf-8')
    print(f'Page {page_number} has been processed.')
    page_number+=1

df_irm = df_irm[['authorName','stockCode','coname','question','reply']]    
df_irm.to_excel('./irm_data.xlsx',encoding = 'utf-8')

try:
    df_sns = pd.read_excel('./sns_data.xlsx')
    
except:
    df_sns = pd.DataFrame()
    column_lst = ['authorName','stockCode','coname','question','reply']
    for item in column_lst:
        df_sns[item] = []
    
page_number = round(len(df_sns)/10+1)


while True:
    url = f'http://sns.sseinfo.com/ajax/feeds.do?type=11&pageSize=10&lastid=-1&show=1&page={page_number}'
    try:
        r = requests.get(url)
    except:
        time.sleep(10)
        try:
            r = requests.get(url)
        except:
            break
    resultBoxes = bs(r.content,"lxml").findAll(attrs = {'class':'m_feed_item'})
    if len(resultBoxes) == 0:
        break
    
    row_index = len(df_sns)
    for i in range(len(resultBoxes)):
        question_box = resultBoxes[i].find_all(attrs = {'class':'m_feed_detail m_qa_detail'})
        reply_box = resultBoxes[i].find_all(attrs = {'class':'m_feed_detail m_qa'})

        user_id = question_box[0].find_all('img')[0]['title']
        company = question_box[0].find_all(attrs = {'class':'m_feed_txt'})[0].text.split(")")[0].split("(")[0].split(":")[1]
        stock_code = question_box[0].find_all(attrs = {'class':'m_feed_txt'})[0].text.split(")")[0].split("(")[1]
        question = question_box[0].find_all(attrs = {'class':'m_feed_txt'})[0].text.split(")")[1].replace('\t',"").replace("\n","")
        reply = reply_box[0].find_all(attrs = {'class':"m_feed_txt"})[0].text.replace("\t","").replace('\n',"")
        
        df_sns.loc[row_index,'authorName'] = user_id
        df_sns.loc[row_index,'stockCode'] = stock_code
        df_sns.loc[row_index,'coname'] = company
        df_sns.loc[row_index,'question'] = question
        try:
            df_sns.loc[row_index,'reply'] = reply
        except:
            df_sns.loc[row_index,'reply'] = 'None'
        
        
        row_index+=1
    if page_number % 100 == 0:

        df_sns = df_sns[['authorName','stockCode','coname','question','reply']]    
        df_sns.to_excel('./sns_data.xlsx',encoding = 'utf-8')

    print(f'Page_sns {page_number} has been processed.')
    page_number+=1
    
df_sns = df_sns[['authorName','stockCode','coname','question','reply']]    
df_sns.to_excel('./sns_data.xlsx',encoding = 'utf-8')
