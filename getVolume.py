#!/usr/bin/env python
# coding: utf-8
#
#pkgs to pip install: xlsxwriter; beautifulsoup; bs4; lxml; pandas; senenium
#import requests
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import pandas as df

driver = webdriver.Chrome('D:/GitHub/SeleniumDriver/Chrome/chromedriver.exe')
driver.maximize_window()
driver.get("https://www.tdcc.com.tw/QStatWAR/inputframe.htm")
time.sleep(2)

#soup = BeautifulSoup(driver.page_source,'lxml')
#print(soup.prettify())

parentHandle = driver.window_handles[0]
parentHandle_title = driver.title
#print(parentHandle_title)

#<select name="Report" onchange="displayField(this)">
select = Select(driver.find_element_by_name('Report'))
#<option value="indw003">上櫃混藏保管有價證券週餘額表</option>
select.select_by_value('indw003')

#<input type="submit" value="查詢" name="queryButton">
confirm_button = driver.find_element_by_name('queryButton')

confirm_button.click()
time.sleep(5)
childHandle=driver.window_handles[1]
driver.switch_to.window(childHandle)

#code to extract data to excel
soup = BeautifulSoup(driver.page_source, 'lxml')
#weeklyVol = soup.find("table", attrs={"class":"mt"})
#print(soup.prettify)
time.sleep(3)

#quick beautifulsoup tutorial
#https://www.dataquest.io/blog/web-scraping-tutorial-python/
#list(soup.children)
#[type(item) for item in list(soup.children)]

#檔案名稱
data_title = soup.find(class_='head').text
#print(data_title)

#資料日期
date_Update = soup.find('td', class_='bwl9').text
#print(len(date_Update))
#print(date_Update)

#名稱欄
title_row = []
for item in soup.find_all('td',class_='wuc9'):
    #print(item.text)
    title_row.append(item.text)
#print(title_row)

# 證券代號 list
code_coln= []
for element_code in soup.find_all('td',class_='wul9',align='left'):
    code_coln.append(element_code.text)
#print(len(code_coln))
#print first 10 items of the list
#print(code_coln[:10])

#證券名稱
name_coln = []
results = soup.find_all('td',{"class":"wul9"})
for result in results:
    if(len(result.attrs)==1):
        #print(result.text)
        name_coln.append(result.text)
#print(len(name_coln))
#print first 10 items of the list
#print(name_coln[:10])

# 本周股額 / 上週股額 / 增減數額
volume = []
for element_volume in soup.find_all('td',class_='wur9', align='right'):
    volume.append(element_volume.text)
this_week_volume=volume[0::3]
last_week_volume=volume[1::3]
change_volume=volume[2::3]
#print(change_volume)

#change percentage & issue percentage
percentage_colns = []
for element_percentage in soup.find_all('td',{"class":"wur9"}):
    if(len(element_percentage.attrs)==1):
        percentage_colns.append(element_percentage.text)
#print(percentage_colns)
change_percentage = percentage_colns[0::2]
issue_percentage = percentage_colns[1::2]
#print(issue_percentage[:10])

weekly_volume = df.DataFrame(
    {title_row[0]: code_coln,
     title_row[1]: name_coln,
     title_row[2]: this_week_volume,
     title_row[3]: last_week_volume,
     title_row[4]: change_volume,
     title_row[5]: change_percentage,
     title_row[6]: issue_percentage
    })
#print(weekly_volume.iloc[:10])
time.sleep(1)

#export csv
#string manipulation: slicing strings
export_path_csv = './files/' + data_title + '_'+date_Update[-7:] + '.csv'
export_path_excel = './files/' + data_title + '_'+date_Update[-7:] + '.xlsx'
#print(export_path_csv)
#print(export_path_excel)
weekly_volume.to_csv(export_path_csv, index='false',encoding='utf-8')
weekly_volume.to_excel(export_path_excel, engine='xlsxwriter', index=True,encoding='utf-8')
time.sleep(1)

driver.close()
time.sleep(1)
driver.switch_to.window(parentHandle)
time.sleep(1)
driver.close()


# In[ ]:




