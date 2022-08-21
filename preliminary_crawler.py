import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os,time
import pandas as pd
import random
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import subprocess
import os

def _start_webdriver(browser, driverpath):
    '''
    Function to start a webdriver

    Parameters
    ----------
    browser : str
        Type of the browser you want to open
    driverpath : str
        Path of the driver.

    Returns
    -------
    selenium.webdriver
        Webdriver object for further usage
    '''
    
    if browser.lower() == 'edge':
        return webdriver.Edge(executable_path=driverpath)
    elif browser.lower() == 'chrome':
        return webdriver.Chrome(executable_path=driverpath)
    else:
        raise NotImplementedError(f'Code for {browser} is not implemented')
        
def _open_browser_cmd(port, cache_dir):
    '''
    Open chrome in debugging mode
    '''
    chrome_cmd = f'chrome.exe --remote-debugging-port={port} --user-data-dir="{cache_dir}"'
    subprocess.Popen(chrome_cmd)
    
def _connect_selenium(driverpath, port, cache_dir):
    '''
    connect your browser to python

    Returns
    -------
    driver: Selenium.webdriver object that is connected to your browser

    '''
    
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    
    driver = webdriver.Chrome(driverpath, options=chrome_options)
    
    return driver


def extract_rating(driver,conm):
    #  close the pop up window
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
#     driver.find_element_by_tag_name('body').send_keys(Keys.ESCAPE)
    #  clear the searching braket
    driver.find_element_by_xpath('//*[@id="sc.keyword"]').send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    #  input keys to the search braket
    driver.find_element_by_xpath('//*[@id="sc.keyword"]').send_keys(conm)
    #  click the search buttom
    driver.find_element_by_xpath('//*[@id="scBar"]/div/button').click()
    count = 0
    find = False
    while True:
        try: 
            element1 = driver.find_element_by_xpath('//*[@id="SearchSuggestions"]/p[1]')
            find = True
            break
        except:
            try:
                element2 = driver.find_element_by_xpath('//*[@id="EIOverviewContainer"]/div/div[3]/div[1]/div[2]')
                find = True
                break
            except:
                try:
                    element3 = driver.find_element_by_xpath('//*[@id="MainCol"]/div/div[1]/div/div[1]/div/div[2]/h2/a')
                    find = True
                    break
                except:
                    try:
                        element4 = driver.find_element_by_xpath('//*[@id="EIOverviewContainer"]/div/div[5]/div/div[1]/div[2]/div')
                        find = True
                        break
                    except:
                        time.sleep(1)
                        count+=1
                        if count==10:
                            break
            
    time.sleep(random.random()+3)       
    rating = ''
    coaddress = ''
    outcome_conm = ''
    try: 
        element = element1
        rating = element.text

    except: 
        pass
    try: 
        element = element2
        #  unfold the rating
        #  acquire the company address and website
        outcome_conm = driver.find_element_by_xpath('//*[@id="DivisionsDropdownComponent"]').text
        buttom = driver.find_element_by_xpath('//*[@id="EIOverviewContainer"]/div/div[3]/div[1]/div[2]')
        webdriver.ActionChains(driver).move_to_element(buttom).click(buttom).perform()
        time.sleep(random.random()+4)
        #  find the deversity and inclusion rating
        rating = driver.find_element_by_xpath('//*[@id="reviewDetailsModal"]/div[2]/div[2]/div/div/div/div[1]/div[1]/div/div[3]/div/div[3]').text

    except:
        pass
    try: 
        element = element3
        coaddress = driver.find_element_by_xpath('//*[@id="MainCol"]/div/div[1]/div/div[1]/div/div[2]/div').text
        #  click the first outcome
        first_outcome = driver.find_element_by_xpath('//*[@id="MainCol"]/div/div[1]/div/div[1]/div/div[2]/h2/a')
        first_outcome.click()
        while True:
            try:
                with_rating = driver.find_element_by_xpath('//*[@id="EIOverviewContainer"]/div/div[3]/div[1]/div[2]')
                break
            except:
                try:
                    without_rating = driver.find_element_by_xpath('//*[@id="EIOverviewContainer"]/div/div[5]/div/div[1]/div[2]/div')
                    break
                except:
                    time.sleep(0.5)
                
        #  acquire the company address and website
        outcome_conm = driver.find_element_by_xpath('//*[@id="DivisionsDropdownComponent"]').text
       
        try:
            buttom = driver.find_element_by_xpath('//*[@id="EIOverviewContainer"]/div/div[3]/div[1]/div[2]')
            webdriver.ActionChains(driver).move_to_element(buttom).click(buttom).perform()
            time.sleep(random.random()+4)
             #  find the deversity and inclusion rating
            rating = driver.find_element_by_xpath(
                    '//*[@id="reviewDetailsModal"]/div[2]/div[2]/div/div/div/div[1]/div[1]/div/div[3]/div/div[3]').text
        except:
            try:
                driver.find_element_by_xpath('//*[@id="EIOverviewContainer"]/div/div[5]/div/div[1]/div[2]/div')
                rating = 'This company has no rating.'
               
            except:
                pass
    except:
        pass
    try:
        element = element4
        rating = 'This company has no rating.'
        outcome_conm = driver.find_element_by_xpath('//*[@id="DivisionsDropdownComponent"]').text
    except:
        pass
    if find == False:
        rating = "Special Situation!"
        
    return rating,outcome_conm,coaddress

def renew(original,present):
    firm_sub = df[original:present]
    firm_sub['rating'] = ratings
    firm_sub['out_name'] = out_name
    firm_sub['address'] = address
    print(firm_sub.head())
    print(firm_sub.tail())
    pd.concat([pd.read_excel('./recrawl_{}.xlsx'.format(original)),firm_sub]).to_excel('recrawl_{}.xlsx'.format(present))
    return None


_open_browser_cmd(8080,'C:/Users/Wendy/Desktop/glassdoor/driver/')
driver = _connect_selenium('./chromedriver.exe',8080,'C:/Users/Wendy/Desktop/glassdoor/driver/')

df = pd.read_excel('./recrawl.xlsx')
driver.get("https://www.glassdoor.com.hk/member/home/companies.htm?countryRedirect=true")

start_index = ......
firms = df[Start_index:]
ratings = []
address = []
out_name = []
count = 0
for i,row in firms.iterrows():
    conm = row['pname']
    info = extract_rating(driver,conm)
    ratings.append(list(info)[0])
    out_name.append(list(info)[1])
    address.append(list(info)[2])
    print(info)
    time.sleep(random.random()+3)
 

 renew(Start_index,Start_index+len(ratings))
