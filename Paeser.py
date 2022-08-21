import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup as bs
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os,time
import random
import datetime
import selenium
from selenium.webdriver.chrome.options import Options
import subprocess
from multiprocessing.dummy import Process, Queue

# find the date of the review
def __find_the_date(soup):
    date_soup = soup.findAll(attrs = {"class":"authorInfo"})
    date = date_soup[0].text[0:12]
    date = date.replace("-","")
    return date

# find the overall score of the review
def __find_overall(soup):
    return soup.span.text

# transfer the star type to score
def __transfer_type_to_score(star_type):
    if star_type == 'css-xd4dom':
        score = 1
    elif star_type == 'css-qt3l8j':
        score = 1.5
    elif star_type == 'css-18v8tui':
        score = 2
    elif star_type == 'css-z2gtf':
        score = 2.5
    elif star_type == 'css-vl2edp':
        score = 3
    elif star_type == 'css-11rm5hs':
        score = 3.5
    elif star_type == 'css-1nuumx7':
        score = 4
    elif star_type == 'css-1areqgb':
        score = 4.5
    elif star_type == 'css-s88v13':
        score = 5
    return score

# find the specific scores and their classification
def __find_scores(soup):
    star_types = []
    classifications = []
    aside = soup.aside
    if aside == None:
        return None
    ratings = aside.find_all("li") 
    for i in range(len(ratings)):
        class_soup = ratings[i].find_all("div")
        classification = class_soup[0].text
        star_type = class_soup[1].attrs["class"][0]
        star_types.append(star_type)
        classifications.append(classification)
        scores = [__transfer_type_to_score(x) for x in star_types]
        
    return classifications,scores

# edit the dataframe and save as excel
def edit_save(company_soups,company,coname):
    for i in range(len(company_soups)):
        soup = company_soups.loc[i,'reviews']
        date = __find_the_date(soup)
        overall = __find_overall(soup)
        company_soups.loc[i,'Date'] = date
        company_soups.loc[i,'Overall'] = overall
        if __find_scores(soup) != None:
            classifications = __find_scores(soup)[0]
            scores = __find_scores(soup)[1]
            for j in range(len(scores)):
                company_soups.loc[i,classifications[j]] = scores[j]
    outname = f'{coname.replace(".pickle","")}.xlsx'
    outdir = f"./excel_folder2/{company}"
#     if not os.path.exists(outdir):
#         os.mkdir(outdir)
    fullname = os.path.join(outdir, outname)
    company_soups.iloc[:,0:9].to_excel(fullname)            
    return None

def find_process_folder():
    pickle_folder_lst = os.listdir('./pickle_process_folder')
    excel_folder_lst = os.listdir('./excel_folder2')
    processing_folder_lst = []
    for folder in pickle_folder_lst:
        if not folder in excel_folder_lst:
            processing_folder_lst.append(folder)
            
    return processing_folder_lst


folder_lst = find_process_folder()
for company in folder_lst:
    os.mkdir(f"./excel_folder2/{company}")
    file_lst = os.listdir(f"./pickle_process_folder/{company}")
    for file in file_lst:
        soup = pd.read_pickle(f"./pickle_process_folder/{company}/{file}")
        edit_save(soup,company,file)
        
    print(f"{company} saved.")