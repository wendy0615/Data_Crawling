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

def sign_in():
    #  extract the useremail and keyword
    message = pd.read_excel('./sign.xlsx')
    user_name = message['user'][0]
    key = message['key'][0]
    user_input = driver.find_element_by_xpath('//*[@id="inlineUserEmail"]')
    key_input = driver.find_element_by_xpath('//*[@id="inlineUserPassword"]')

    #  send username and key, press the sign-in button
    user_input.send_keys(user_name)
    key_input.send_keys(int(key))
    sign_in_button = driver.find_element_by_xpath('//*[@id="InlineLoginModule"]/div/div[1]/div/form/div[3]/button')
    sign_in_button.click()
    
    return None

def input_conm_click(conm):
    #  locate the input box
    while True:
        input_box = driver.find_elements_by_css_selector('#sc\.keyword')
        if len(input_box) != 0:
            break
    #  clear the input box
    input_box[0].send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
    #  send the company name to the input box
    in_lst = conm.split(" ")
    if len(in_lst) > 1:
        in_lst = in_lst[:-1]
    conm = ''
    for i in in_lst:
        conm += i + ' '
    input_box[0].send_keys(conm)
    #  click the search button
    search = driver.find_element_by_xpath('//*[@id="scBar"]/div/button')
    webdriver.ActionChains(driver).move_to_element(search).click(search).perform()
    return None

def judge_result():
    while True:
        review_button = driver.find_elements_by_css_selector('#EIProductHeaders > div > a.eiCell.cell.reviews')
        companies = driver.find_elements(By.CLASS_NAME,'col-lg-7')
        search_suggestion = driver.find_elements_by_css_selector('#SearchSuggestions')
        if review_button != [] or companies != [] or search_suggestion !=0:
            break            
    return review_button,companies,search_suggestion

def matching(companies,coname):
    for company in companies:
        soup = bs(company.get_attribute('innerHTML'),'html.parser')
        coinfo = company.text
        # locate the row
        company_row = firms[firms['coname']==coname]
        # select the city
        company_city = company_row['city'].iloc[0]
        if company_city in coinfo:
            url = 'https://www.glassdoor.com.hk/'+soup.find_all(attrs={"class":"sqLogoLink"})[0]['href']
            return url
    
    first_company_soup = bs(companies[0].get_attribute('innerHTML'),'html.parser')
    url = 'https://www.glassdoor.com.hk/'+first_company_soup.find_all(attrs={"class":"sqLogoLink"})[0]['href']
    return url
        
# def open_specific_page():
#     url = matching(companies,coname)
#     if url == None:
#         first_company_soup = bs(companies[0].get_attribute('innerHTML'),'html.parser')
#         url = first_company_soup.find_all(attrs={"class":"sqLogoLink"})[0]['href']
#     driver.get(url)
#     return None   

def open_specific_page(url):
    if url == None:
        first_company_soup = bs(companies[0].get_attribute('innerHTML'),'html.parser')
        url = first_company_soup.find_all(attrs={"class":"sqLogoLink"})[0]['href']
    driver.get(url)
    return None

# def get_city():
#     city = driver.find_element_by_css_selector('#EmpBasicInfo > div:nth-child(1) > div > div:nth-child(2)').text
#     return city

def press_review_button():
    while True:
        try:
            review_button = driver.find_elements_by_css_selector('#EIProductHeaders > div > a.eiCell.cell.reviews')
            
            button = review_button[0]
            webdriver.ActionChains(driver).move_to_element(button).click(button).perform()
            time.sleep(random.random()+1.3)
            break
        except:
            continue
    return None

def __judge_if_reviews():
    while True:
        web_name = driver.find_elements_by_css_selector('#DivisionsDropdownComponent')
        review_amount = driver.find_elements_by_css_selector('#MainContent > div > div:nth-child(1) > div.eiReviews__EIReviewsPageStyles__EIReviewsPage.pt-0.pr-std.pb-std.pl-std.p-md-std.els8ktd0.gd-ui-module.css-8hewl0.ec4dwm00 > div.common__EIReviewSortBarStyles__sortsHeader.row.justify-content-between.mt-md-xl.mt-sm > h2 > span > strong:nth-child(1)')
        no_review = driver.find_elements_by_css_selector('#MainContent > div > div:nth-child(1) > div.eiReviews__EIReviewsPageStyles__EIReviewsPage.pt-0.pr-std.pb-std.pl-std.p-md-std.els8ktd0.gd-ui-module.css-8hewl0.ec4dwm00 > div.eiReviews__EIReviewsPageStyles__filterDropdown.mt-xl > div > div:nth-child(3) > p')
        no_reviews = driver.find_elements_by_css_selector('#MainContent > div > div:nth-child(1) > div.eiReviews__EIReviewsPageStyles__EIReviewsPage.pt-0.pr-std.pb-std.pl-std.p-md-std.els8ktd0.gd-ui-module.css-8hewl0.ec4dwm00 > div.eiReviews__EIReviewsPageStyles__filterDropdown.mt-xl > div > div:nth-child(2)')
        first_page_url = driver.current_url
        if len(review_amount) != 0 or len(no_review) != 0 or len(no_reviews) != 0:
            break
    return review_amount,no_review,web_name,first_page_url,no_reviews

def __parser_page_amount(review_amount):
    text = review_amount.text
    amount = int(text.replace(",",""))
    if amount%10 == 0:
        page_amount = amount//10
    else:
        page_amount = amount//10+1
    print(f'Totally {page_amount} pages')
        
    return page_amount

def __get_one_page():
    try_count = 0
    while True:
        divs = driver.find_elements_by_class_name('gdReview')
        node = driver.find_elements_by_css_selector('#NodeReplace > main > div > div')
        if len(node)!=0 and "Something's wrong." in node[0].text:
            driver.refresh()
#         web_name = driver.find_elements_by_css_selector('#DivisionsDropdownComponent')
        if len(divs)>0:
            try:
                divs_soup = [bs(div.get_attribute('innerHTML'),'html.parser') for div in divs]
#                 web_name = web_name[0].text
#                 review_boxes += divs_soup
                try_count+=1
                print(f'This is the {try_count} time to try.Succeed')
                break
            except:
                time.sleep(random.random()+1.2)
                try_count+=1
                continue


            if try_count >= 10:
                print('Have benn tried for more than 10 times but failed')
                break

    #             time.sleep(random.random()+1.3)
    return divs_soup

def __turn_next_page(page_amount):
    while page_amount > 1:
        try:
            next_page = driver.find_element_by_css_selector('#MainContent > div > div:nth-child(1) > div.d-flex.flex-column.align-items-top > div > div.pageContainer > button.nextButton.css-1hq9k8.e13qs2071')
            webdriver.ActionChains(driver).move_to_element(next_page).click(next_page).perform()
            time.sleep(random.random()+1.1)
            break
        except:
            continue
    return None

def __extract_cocode(first_page_url):
    cocode = ''
    for letter in first_page_url:
        if letter.isdigit():
            cocode+=letter
    return cocode

# def determine_start_index():
#     if os.path.exists(f'./pickle_process_folder/{coname}.pickle'):
#         past_data = pd.read_pickle(f'./pickle_process_folder/{coname}.pickle')
#         review_boxes += list(past_data['reviews'])
#         start_page = int((len(review_boxes)/10)+1)
#     else:
#         start_page = 1
#     return start_page


def __create_dir_save(coname,start_page,review_boxes):
    
    df = pd.DataFrame()
    df['coname'] = []
    df['Date'] = []
    df['Overall'] = []
    df['Work/Life Balance'] = []
    df['Culture & Values'] = []
    df['Diversity & Inclusion'] = []
    df['Career Opportunities'] = []
    df['Compensation and Benefits'] = []
    df['Senior Management'] = []
    df['reviews'] = review_boxes
    
    outdir = f'./pickle_process_folder/{coname.replace("/","_")}'
    if not os.path.exists(outdir):
        os.mkdir(outdir)
    outname = f'{coname.replace("/","_")}_{start_page/50}.pickle'
    fullname = os.path.join(outdir, outname)
    coname_lst = []
    for i,row in df.iterrows():
        coname_lst.append(coname)
    df['coname'] = coname_lst
    df.to_pickle(fullname)
    return None


def get_review_boxes():
#     soup_lst = []
    
    start_page = 1
    review_boxes = []
    judge_of_reviews = __judge_if_reviews()
    review_amount = judge_of_reviews[0]
    no_review = judge_of_reviews[1]
    first_page_url = judge_of_reviews[3]
    no_reviews = judge_of_reviews[4]
    web_name = judge_of_reviews[2][0].text
    web_name = web_name.replace(' ','-')
    cocode = __extract_cocode(first_page_url)
    if len(no_review) != 0 or len(no_reviews) != 0:
        review_boxes = []
        return review_boxes
    review_amount = review_amount[0]
    page_amount = __parser_page_amount(review_amount)
    if os.path.exists(f'./pickle_process_folder/{coname.replace("/","_")}'):
        start_page = len(os.listdir(f'./pickle_process_folder/{coname.replace("/","_")}'))*50
        driver.get(f'https://www.glassdoor.com.hk/Reviews/{web_name}-Reviews-E{cocode}_P{start_page}.htm?filter.iso3Language=eng')
    print(f'Now processing the page {start_page}/{page_amount}.')
    divs_soup = __get_one_page()
    review_boxes += divs_soup
    if page_amount == 1:
        return review_boxes
    while True:
        __turn_next_page(page_amount)
        print(f'Now processing the page {start_page+1}/{page_amount}.')
        current_url = driver.current_url
        if not f'P{start_page+1}' in current_url:
            page = start_page+1
            driver.get(f'https://www.glassdoor.com.hk/Reviews/{web_name}-Reviews-E{cocode}_P{page}.htm?filter.iso3Language=eng')
    
        divs_soup = __get_one_page()
        review_boxes += divs_soup
        start_page+=1
        if start_page%50 == 0:
            __create_dir_save(coname,start_page,review_boxes)
            print(f'{start_page} have been saved!')
            review_boxes = []
        if start_page >= page_amount:
            break
        
    return review_boxes






def save_to_pickle(coname,review_boxes):
    df = pd.DataFrame()
    df['coname'] = []
    df['Date'] = []
    df['Overall'] = []
    df['Work/Life Balance'] = []
    df['Culture & Values'] = []
    df['Diversity & Inclusion'] = []
    df['Career Opportunities'] = []
    df['Compensation and Benefits'] = []
    df['Senior Management'] = []
#     df['City'] = []

    outdir = f'./pickle_process_folder/{coname.replace("/","_")}'
    if not os.path.exists(outdir):
        os.mkdir(outdir)
    outname = f'{coname.replace("/","_")}_final.pickle'
    fullname = os.path.join(outdir, outname)

    df['reviews'] = review_boxes
    coname_lst = []
#     city_lst = []
    for i,row in df.iterrows():
        coname_lst.append(coname)
#         city_lst.append(city)
    df['coname'] = coname_lst
#     df['City'] = city_lst
    df.to_pickle(fullname)
    print('{} was saved.'.format(coname))
    
    return None

def download(coname):
    input_conm_click(coname)
    search_result = judge_result()
    companies = search_result[1]
    search_suggestion = search_result[2]
    if len(search_suggestion) != 0:
        return None
    if len(companies) != 0:
        url = matching(companies,coname)
        driver.get(url)    
    time.sleep(random.random()+3)
#     city = get_city()
    press_review_button()
    review_boxes = get_review_boxes()
    save_to_pickle(coname,review_boxes)
    company_button = driver.find_element_by_css_selector('#SiteNav > nav.d-none.d-lg-block.HeaderStyles__bottomShadow.HeaderStyles__navigationBackground.HeaderStyles__primaryNavigation > div > div.px-std.px-md-lg.col.HeaderStyles__navigationScroll > div.d-inline-flex.align-items-center.mr-xl.HeaderStyles__navigationItem.HeaderStyles__activeNavigationItem > div > a')
    webdriver.ActionChains(driver).move_to_element(company_button).click(company_button).perform()
    time.sleep(random.random()+3)
    
    return None

_open_browser_cmd(8080,'C:/Users/Wendy/Desktop/specific rating/driver/')
driver = _connect_selenium('./chromedriver.exe',8080,'C:/Users/Wendy/Desktop/specific rating/driver/')
# driver.maximize_window()
driver.get('https://www.glassdoor.com.hk/member/home/companies.htm')
time.sleep(random.random()+3)
firms = pd.read_excel('sp500.xlsx')
coname_lst = firms['coname']
index_start = 2810
for coname in coname_lst[index_start:]:
    download(coname)
    index_start+=1
    print(f'The start index is {index_start}')