#!/usr/bin/env python
# coding: utf-8

# In[1]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains

from time import sleep
from dateutil import relativedelta
import os, sys, re, csv, time, io, filecmp
import datetime as dt
import pandas as pd

from module.func import *
from GUI.救指調假_app import *
from PyQt5.QtWidgets import QApplication, QMainWindow


# In[2]:


# Open GUI to select/read filenames
app = QtWidgets.QApplication(sys.argv)
MainWindow = QtWidgets.QMainWindow()
ui = MainWindowUIClass()
ui.setupUi(MainWindow)
MainWindow.show()
app.exec_()

(year, month, account, password, start_day, end_day) = ui.getParam()
if not (year and month and account and password and start_day and 
        end_day):
    sys.exit()


# In[5]:


dir_path = r'C:\\Users\\TFD\\救指調假\\救指調假資訊\\'

errleavenum = 0
cont_leave_name = []


# In[6]:


option = webdriver.ChromeOptions()
# option.add_argument('--headless')
# option.add_argument('--window-size=1280,800"')
if getattr(sys, 'frozen', False) :
    # running in a bundle
    chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
    browser = webdriver.Chrome(chromedriver_path, options=option)
else:
    # executed as a simple script, the driver should be in "PATH"
    browser = webdriver.Chrome(options=option)
#開啟google首頁
browser.get('https://webitr.gov.taipei/WebITR/')


# In[7]:


# Login
element_user = browser.find_element_by_id("userName")
element_user.send_keys(account)
element_pw = browser.find_element_by_id("login_key")
element_pw.send_keys(password)
element_pw.send_keys(Keys.RETURN)


# In[8]:


# Click "差勤管理/請假管理/請假資料維護"
switch_to_topmost(browser)
wait = WebDriverWait(browser, 10)
## 差勤管理
li1 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > a")
ActionChains(browser).move_to_element_with_offset(li1, 10, -10).move_to_element_with_offset(li1, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > a")))
## 請假管理
li2 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > a")
ActionChains(browser).move_to_element_with_offset(li2, 10, -10).move_to_element_with_offset(li2, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > ul > li:nth-child(1) > a")))
## 請假資料維護
li3 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > ul > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li3, 10, -10).move_to_element_with_offset(li3, 10, 10).perform()
li3.click()
## 移出左方列表使浮出列表消失
li4 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li4, 10, -10).perform()
element = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > a")))

switch_to_iframe(browser)
browser.find_element_by_link_text("[請假資料查詢、編輯]").click()


# In[9]:


# Select 救災救護指揮中心
select = Select(browser.find_element_by_id("unit"))
select.select_by_visible_text("救災救護指揮中心")
# Select all members
selAllPerson = browser.find_element_by_id('selAll').click()


# In[10]:


# Search start day
datestrbtn = browser.find_element_by_id("strDate_dc_bt")
browser.execute_script("arguments[0].scrollIntoView(false);", datestrbtn)
datestrbtn.click()
chooseDay(start_day, browser)

# Search end day
dateendbtn = browser.find_element_by_id("endDate_dc_bt")
browser.execute_script("arguments[0].scrollIntoView(false);", dateendbtn)
dateendbtn.click()
chooseDay(end_day, browser)

# Press query button
querybtn = browser.find_element_by_id('queryBtn')
browser.execute_script("arguments[0].scrollIntoView(false);", querybtn)
querybtn.click()


# In[11]:


leave_table = browser.find_element_by_css_selector(
    '#unitsResult > div > div > table > tbody')
leaves = leave_table.find_elements_by_class_name('stripeMe')
for leave in leaves:
    # Select wrong data
    if '合計日時數：0.0' in leave.text:
        errleavenum += 1
        main_window = browser.window_handles[0]
        modify_btn = leave.find_element_by_id('modifyLink')
        browser.execute_script("arguments[0].scrollIntoView(false);", modify_btn)
        modify_btn.send_keys(Keys.CONTROL + Keys.RETURN)
        
        # Edit in new tab
        pop_window = browser.window_handles[1]
        browser.switch_to.window(pop_window)
        element = wait.until(EC.visibility_of_element_located((By.NAME, "strDate_xx")))
        # Read start date and time
        strDate_str = browser.find_element_by_id('strDate_xx').get_attribute('value')
        strDate_list = strDate_str.split('-')
        strDate_year = str(int(strDate_list[0]) + 1911)
        strDate_mon = strDate_list[1]
        strDate_day = strDate_list[2]
        strTime_str = browser.find_element_by_id('strTime').get_attribute('value')
        strTime = strTime_str[0:2] + '-' + strTime_str[2:4]
        strDT_str = strDate_year + '-' + strDate_mon + '-' + strDate_day + '-' + strTime
        strDate = dt.datetime.strptime(strDT_str, '%Y-%m-%d-%H-%M')
        # Read end date and time
        endDate_str = browser.find_element_by_id('endDate_xx').get_attribute('value')
        endDate_list = endDate_str.split('-')
        endDate_year = str(int(endDate_list[0]) + 1911)
        endDate_mon = endDate_list[1]
        endDate_day = endDate_list[2]
        endTime_str = browser.find_element_by_id('endTime').get_attribute('value')
        endTime = endTime_str[0:2] + '-' + endTime_str[2:4]
        endDT_str = endDate_year + '-' + endDate_mon + '-' + endDate_day + '-' + endTime
        endDate = dt.datetime.strptime(endDT_str, '%Y-%m-%d-%H-%M')
        # Calculate time delta
        timedelta = relativedelta.relativedelta(endDate, strDate)
        if timedelta.hours == 0 and timedelta.minutes == 0:
            povdays = browser.find_element_by_id('povdays')
            browser.execute_script("arguments[0].scrollIntoView(false);", povdays)
            povdays.clear()
            povdays.send_keys(str(float(timedelta.days)))
        # Click edit button
        btns = browser.find_elements_by_tag_name('input')
        for btn in btns:
            if btn.get_attribute('value') == '修改':
                edit_btn = btn
                break
        browser.execute_script("arguments[0].scrollIntoView(false);", edit_btn)
        edit_btn.click()
        # Close the edit tab
        browser.close()
        browser.switch_to.window(main_window)
        switch_to_iframe(browser)


# In[12]:


# Select 救災救護指揮中心
select = Select(browser.find_element_by_id("unit"))
select.select_by_visible_text("救災救護指揮中心")
members_span = browser.find_element_by_id('location_persons')
mem_num = len(members_span.find_elements_by_name('persons'))

for i in range(mem_num):
    # Select 救災救護指揮中心
    select = Select(browser.find_element_by_id("unit"))
    select.select_by_visible_text("救災救護指揮中心")
    # Deselect all members
    selAllPerson = browser.find_element_by_id('selAll')
    selAllPerson.click()
    selAllPerson.click()
    # Select member
    members_span = browser.find_element_by_id('location_persons')
    mem = members_span.find_elements_by_name('persons')[i]
    browser.execute_script("arguments[0].scrollIntoView(false);", mem)
    mem.click()
    # Search start day
    datestrbtn = browser.find_element_by_id("strDate_dc_bt")
    browser.execute_script("arguments[0].scrollIntoView(false);", datestrbtn)
    datestrbtn.click()
    chooseDay(start_day, browser)
    # Search end day
    dateendbtn = browser.find_element_by_id("endDate_dc_bt")
    browser.execute_script("arguments[0].scrollIntoView(false);", dateendbtn)
    dateendbtn.click()
    chooseDay(end_day, browser)
    # Press query button
    querybtn = browser.find_element_by_id('queryBtn')
    browser.execute_script("arguments[0].scrollIntoView(false);", querybtn)
    querybtn.click()

    # Check for continuous leaves
    cont_start = dt.datetime(1900, 1, 1)
    cont_end = dt.datetime(1900, 1, 1)
    leavelist = []
    elseleavelist = []
    leavesum = []
    elseleavesum = []
    leave_table = browser.find_element_by_css_selector(
        '#unitsResult > div > div > table > tbody')
    leaves = leave_table.find_elements_by_class_name('stripeMe')
    for leave in leaves:
        if not isValidContinuousLeave(leave.text, cont_end):
            updateContinueLeave(leavelist, elseleavelist, leavesum,
                                elseleavesum, cont_start, cont_end, 
                                cont_leave_name, browser)
            cont_start = str2Dates(leave.text)[0]
        if re.search('(休假、請假期間)', leave.text):
        # 休假
            leavelist.append(leave)
            leavesum.append(leavePeriodDays(leave.text)) 
        else:
        # 其他假
            elseleavelist.append(leave)
            elseleavesum.append(leavePeriodDays(leave.text)) 
        cont_end = str2Dates(leave.text)[1]
    updateContinueLeave(leavelist, elseleavelist, leavesum, 
                        elseleavesum, cont_start, cont_end, 
                        cont_leave_name, browser)       
        
browser.quit()


# In[15]:


# Save information
filename = str(year - 1911) + '年' + str(month) + '月救指調假資訊.txt'
with io.open(dir_path + filename, 'w', encoding='utf8') as outfile:
    outfile.write('已修改 ' + str(errleavenum) + ' 筆錯誤假單\n')
    outfile.write('調整 ' + str(len(cont_leave_name)) + 
                  ' 筆連續休假，分別為:\n')
    for cleave_info in cont_leave_name:
        outfile.write(cleave_info + '\n')


# In[16]:


# Summation
print('已修改 ' + str(errleavenum) + ' 筆錯誤假單')
print('調整 ' + str(len(cont_leave_name)) + ' 筆連續休假，分別為:')
for cleave_info in cont_leave_name:
    print(cleave_info)
    
print('調假資訊已新增於: ' + dir_path)
input("\n按任意鍵結束")

