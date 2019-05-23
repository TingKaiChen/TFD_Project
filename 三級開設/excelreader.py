#!/usr/bin/env python
# coding: utf-8

# In[1]:


import xlwings as xw
import datetime as dt
import pandas as pd
import numpy as np

from dateutil import relativedelta
from bs4 import BeautifulSoup 


# In[2]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains

from time import sleep
import os, sys, re, csv, time, io, filecmp

from module.func import *
from GUI.三級加班_app import *
from PyQt5.QtWidgets import QApplication, QMainWindow


# In[3]:


# Open GUI to select/read the filename, year, month and account info
app = QApplication(sys.argv)
MainWindow = QMainWindow()
ui = MainWindowUIClass()
ui.setupUi(MainWindow)
MainWindow.show()
app.exec_()
(filename, filesht, year, month, account, password) = ui.getParam()

if not (filename and filesht and year and month and account and 
        password):
    ui.clear()
    sys.exit()

# Determine the file name and date
Project_Name = '%d年%d月本市災害應變中心三級開設本局進駐'%(
    year - 1911, month)
Year = str(year)
Month = str(month)
print("專案名稱: " + Project_Name)


# In[4]:


Green = (146, 208, 80)
Red = (255, 0, 0)
Green_t = 1
Red_t = 2

# Unit list
OfficeList = {"資通作業科": "資通作業科", 
              "資通科": "資通作業科",
              "減災規劃科": "減災規劃科", 
              "減災科": "減災規劃科",
              "整備應變科": "整備應變科", 
              "應變科": "整備應變科",
              "災害搶救科": "災害搶救科", 
              "搶救科": "災害搶救科",
              "火災調查科": "火災調查科", 
              "火調科": "火災調查科", 
              "綜合企劃科": "綜合企劃科", 
              "綜企科": "綜合企劃科",
              "緊急救護科": "緊急救護科", 
              "救護科": "緊急救護科",
              "督察室": "督察室", 
              "主任秘書室": "主任秘書室", 
              "秘書室": "秘書室", 
              "人事室": "人事室", 
              "政風室": "政風室", 
              "會計室": "會計室", 
              "訓練中心": "訓練中心", 
              "救指中心": "救災救護指揮中心", 
              "火災預防科": "火災預防科", 
              "火預科": "火災預防科",
              "局長室": "局長室", 
              "第一副局長室": "第一副局長室", 
              "第二副局長室": "第二副局長室",
              "第三副局長室": "第三副局長室"} 
# Name list
namelist = {}
for key in list(OfficeList.values()):
    namelist[key] = set()


# In[5]:


# xlwings workbook
print("正在檢查加班時數表格式...", end = '')
workbook = xw.Book(filename)
sheet = workbook.sheets[filesht]
rng = sheet.range('A1').expand('table')
if r'錯誤訊息' not in rng.rows[0].value:
    nCol = sheet.api.UsedRange.Columns.count
    sheet.range(chr(65 + nCol) + '1').value = r'錯誤訊息'
    rng = sheet.range('A1').expand('table')
# pandas dataframe
df = pd.read_excel(filename)


# In[6]:


# Select days in that month
df['startdt'] = None
df['enddt'] = None
df['月份'] = None
for idx in df.index[df.index < len(rng.rows)-1]:
    # Set "月份" column
    try:
        df.loc[idx, '月份'] = df.loc[idx, '日期'].month
    except:
        if type(df.loc[idx, '日期']) == str:
            df.loc[idx, '月份'] = pd.to_datetime(
                df.loc[idx, '日期'], format='%m月%d日', errors='coerce').month
    # Set "Color" column
    #todo: different between rng and df index
    if rng.rows[idx + 1].color == Green:
        df.loc[idx, 'Color'] = Green_t
    elif rng.rows[idx].color == Red:
        df.loc[idx, 'Color'] = Red_t
    else:
        df.loc[idx, 'Color'] = None
        
df['startdt'] = None
df['enddt'] = None
dfidx = df[(df['月份'] == int(Month)) & (df['Color'] != Green_t)].index


# In[7]:


# Overwork period validation
for idx in dfidx:
    # Process overwork period string
    (startdt, enddt) = worktime(df.loc[idx, '值勤時段'], Year)
    df.loc[idx, 'startdt'] = startdt
    df.loc[idx, 'enddt'] = enddt
    
    # Clean data
    df.loc[idx, '單位'] = df.loc[idx, '單位'].strip()
    df.loc[idx, '姓名'] = df.loc[idx, '姓名'].strip()
    
    # Label with red color when the overwork time is wrong
    if df.loc[idx, 'startdt'] == None:
        rng.rows[idx+1].color = (255, 0, 0)
        rng.rows[idx+1][-1].value = '時間區段或格式錯誤'
    else:
        time_diff = (df.loc[idx, 'enddt'] - df.loc[idx, 'startdt']).total_seconds()
        if time_diff <= 0:
            rng.rows[idx+1].color = (255, 0, 0)
            rng.rows[idx+1][-1].value = '結束時間早於開始時間'
        
    # Update name list
    namelist[OfficeList[df.loc[idx, '單位'].strip()]].add(df.loc[idx, '姓名'])

print("完成")


# In[8]:


# Selenium process
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


# In[9]:


# Login
element_user = browser.find_element_by_id("userName")
element_user.send_keys(account)
element_pw = browser.find_element_by_id("login_key")
element_pw.send_keys(password)
element_pw.send_keys(Keys.RETURN)


# In[10]:


# 專案加班建置
# Click "差勤管理/加班管理/專案加班維護"
switch_to_topmost(browser)
wait = WebDriverWait(browser, 10)
## 差勤管理
li1 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > a")
ActionChains(browser).move_to_element_with_offset(li1, 10, -10).move_to_element_with_offset(li1, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(5) > a")))
## 加班管理
li2 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(5) > a")
ActionChains(browser).move_to_element_with_offset(li2, 10, -10).move_to_element_with_offset(li2, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(5) > ul > li:nth-child(1) > a")))
## 專案加班維護
li3 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(5) > ul > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li3, 10, -10).move_to_element_with_offset(li3, 10, 10).perform()
li3.click()
## 移出左方列表使浮出列表消失
li4 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li4, 10, -10).perform()
element = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(1) > a")))


# In[11]:


print('搜尋現存專案')
switch_to_iframe(browser)
browser.find_element_by_link_text('[專案加班立案查詢]').click()


# In[12]:


# Search start day
datestrbtn = browser.find_element_by_id("begintime_dc_bt")
browser.execute_script("arguments[0].scrollIntoView(false);", datestrbtn)
datestrbtn.click()
start_day = dt.datetime.strptime(Year + '-' + Month, '%Y-%m')
chooseDay(start_day, browser)

# Search end day
dateendbtn = browser.find_element_by_id("endtime_dc_bt")
browser.execute_script("arguments[0].scrollIntoView(false);", dateendbtn)
dateendbtn.click()
end_day = start_day + relativedelta.relativedelta(months = 1, days = -1)
chooseDay(end_day, browser)

# Press query button
querybtn = browser.find_element_by_id('b1')
browser.execute_script("arguments[0].scrollIntoView(false);", querybtn)
querybtn.click()


# In[13]:


# Search for the project
table = browser.find_element_by_css_selector("#form1 > table:nth-child(7)")
projects = table.find_elements_by_tag_name("tr")
project = None
projectidx = None
for pidx in range(len(projects)):
    # Pass first two rows
    if pidx in [0, 1]:
        continue
    project_name = projects[pidx].find_elements_by_tag_name('td')[1].text
    if project_name == Project_Name:
        project = projects[pidx]
        projectidx = pidx - 1
if project is None:
    print('無現存專案，新增專案')
    # Create new project
    switch_to_iframe(browser)
    pagebtn = browser.find_element_by_link_text('[新增專案加班立案]')
    browser.execute_script("arguments[0].scrollIntoView(false);", pagebtn)
    pagebtn.click()
    # Enter project name
    pjname_box = browser.find_element_by_id('projectName')
    browser.execute_script("arguments[0].scrollIntoView(false);", pjname_box)
    pjname_box.clear()
    pjname_box.send_keys(Project_Name)
    # Select project type
    Select(browser.find_element_by_id('projectType')).select_by_visible_text('只可領錢')
    # Choose project date period
    beginbtn = browser.find_element_by_id('begintime_dc_bt')
    browser.execute_script("arguments[0].scrollIntoView(false);", beginbtn)
    beginbtn.click()
    chooseDay(start_day, browser)
    endbtn = browser.find_element_by_id('endtime_dc_bt')
    browser.execute_script("arguments[0].scrollIntoView(false);", endbtn)
    endbtn.click()
    chooseDay(end_day, browser)
    # Choose project time period
    Select(browser.find_element_by_id('begTimehh')).select_by_visible_text('00')
    Select(browser.find_element_by_id('begTimemm')).select_by_visible_text('00')
    Select(browser.find_element_by_id('endTimehh')).select_by_visible_text('23')
    Select(browser.find_element_by_id('endTimemm')).select_by_visible_text('59')
    # Enter time limit
    timelimitbtn = browser.find_element_by_id('monthLimit')
    browser.execute_script("arguments[0].scrollIntoView(false);", timelimitbtn)
    timelimitbtn.clear()
    timelimitbtn.send_keys(70)
    # Submit the new project
    submitbtn1 = browser.find_element_by_css_selector('#form1 > table > tbody > tr:nth-child(9) > td > input[type="button"]')
    browser.execute_script("arguments[0].scrollIntoView(false);", submitbtn1)
    submitbtn1.click()
    submitbtn2 = browser.find_element_by_css_selector('#form1 > table:nth-child(5) > tbody > tr:nth-child(4) > td > input[type="button"]:nth-child(1)')
    browser.execute_script("arguments[0].scrollIntoView(false);", submitbtn2)
    submitbtn2.click()
else:
    # Edit existed project
    btnid = 'edit_' + str(projectidx)
    editbtn = project.find_element_by_name(btnid)
    browser.execute_script("arguments[0].scrollIntoView(false);", editbtn)
    editbtn.click()
    # 編輯專案人員
    editmemberbtn = browser.find_element_by_css_selector(
        '#form1 > table:nth-child(11) > tbody > tr:nth-child(9) > td > input[type="button"]:nth-child(1)')
    browser.execute_script("arguments[0].scrollIntoView(false);", editmemberbtn)
    editmemberbtn.click()


# In[14]:


# Adding member
print('新增加班人員')
switch_to_iframe(browser)
departlist = list(namelist.keys())
# Select department
for depart in departlist:
    # Skip empty departments
    if namelist[depart] == set():
        continue
    element = wait.until(EC.presence_of_element_located((By.NAME, "depart")))
    depart_sel = Select(browser.find_element_by_name("depart"))
    depart_sel.select_by_visible_text(depart)
    sleep(1)
    depart_sel = Select(browser.find_element_by_name("depart"))
    depart_sel.select_by_visible_text(depart)

    # Wait for name list updated
    mem_num = 0
    while(mem_num == 0):
        sleep(1)
        mem_sel = Select(browser.find_element_by_name("members"))
        mem_num = len(mem_sel.options)

    # Select workover member in the department
    for mem in namelist[depart]:
        mem_sel.select_by_visible_text(mem)
    add_btn = browser.find_element_by_id("add")
    browser.execute_script("arguments[0].scrollIntoView(false);", add_btn)
    add_btn.click()
    addmembtn = browser.find_element_by_css_selector("#form1 > table:nth-child(7) > tbody > tr:nth-child(11) > td > input[type=\"button\"]")
    browser.execute_script("arguments[0].scrollIntoView(false);", addmembtn)
    addmembtn.click()


# In[63]:


# Construct a name-to-personalID dictionary
print('搜尋加班人員身分證號')
main_window = browser.window_handles[0]
switch_to_accountframe(browser)
admin_login_btn = browser.find_element_by_id('adminLoginByName')
browser.execute_script("arguments[0].scrollIntoView(false);", admin_login_btn)
admin_login_btn.click()
pop_window = browser.window_handles[1]
browser.switch_to.window(pop_window)
NameIdList = {}
for unit in namelist:
    for name in namelist[unit]:
        element = wait.until(EC.visibility_of_element_located((By.NAME, "searchName")))
        name_input = browser.find_element_by_name('searchName')
        browser.execute_script("arguments[0].scrollIntoView(false);", name_input)
        name_input.send_keys(name)
        name_input.send_keys(Keys.RETURN)
        tb = browser.find_elements_by_css_selector('body > div > form > table > tbody')[0]
        trs = tb.find_elements_by_tag_name('tr')
        for i in range(1, len(trs)):
            tr = trs[i]
            if unit in tr.text:
                NameIdList[name] = tr.find_element_by_tag_name('input')                                     .get_attribute('value')
browser.close()
browser.switch_to.window(main_window)


# In[16]:


# 新增加班人員資料
# Click "差勤管理/加班管理/加班資料維護"
switch_to_topmost(browser)
wait = WebDriverWait(browser, 10)
## 差勤管理
li1 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > a")
ActionChains(browser).move_to_element_with_offset(li1, 10, -10).move_to_element_with_offset(li1, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(5) > a")))
## 加班管理
li2 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(5) > a")
ActionChains(browser).move_to_element_with_offset(li2, 10, -10).move_to_element_with_offset(li2, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(5) > ul > li:nth-child(2) > a")))
## 專案加班維護
li3 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(5) > ul > li:nth-child(2) > a")
ActionChains(browser).move_to_element_with_offset(li3, 10, -10).move_to_element_with_offset(li3, 10, 10).perform()
li3.click()
## 移出左方列表使浮出列表消失
li4 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li4, 10, -10).perform()
element = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(1) > a")))


# In[17]:


switch_to_iframe(browser)
browser.find_element_by_link_text('[新增加班紀錄]').click()


# In[18]:


# Add workover data
departlist = list(namelist.keys())
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#unit")))

# 執行兩次以防網路延遲的輸入失敗
for i in range(2):
    print('執行第{}次輸入，共2次'.format(i + 1))
    for idx in dfidx:
        print('\r處理進度: {}%    '.format(int(idx/dfidx[-1]*100)), 
              end = '')
        try:
            if (rng.rows[idx+1].color == Green) or                (rng.rows[idx+1][-1].value != None):
                continue
            # Clear selected item
            deselect(browser)
            # Select unit
            depart_sel = Select(browser.find_element_by_id("unit"))
            depart_sel.select_by_visible_text(OfficeList[df.loc[idx, '單位']])
            # Select member
            mem_name = NameIdList[df.loc[idx, '姓名']]
            name_xpath = "//input[@value='" + mem_name + "']"
            mem_box = browser.find_element_by_xpath(name_xpath)
            browser.execute_script("arguments[0].scrollIntoView(false);", mem_box)
            mem_box.click()
            # Enter start time, end time, and work-over hours
            enter_wo_period(df.loc[idx, 'startdt'], df.loc[idx, 'enddt'], 
                            int(df.loc[idx, '總計時數(1+2)']), browser)
            # Add project reason
            proj_box = browser.find_element_by_id('projectInput')
            browser.execute_script("arguments[0].scrollIntoView(false);", proj_box)
            if proj_box.is_selected():
                browser.find_element_by_id('projectInput').click()
                browser.find_element_by_id('projectInput').click()
            else:
                browser.find_element_by_id('projectInput').click()        
            element = wait.until(EC.visibility_of_element_located((By.ID, "project_ptid")))
            proj_ptid = Select(browser.find_element_by_id('project_ptid'))
            proj_ptid.select_by_visible_text(Project_Name)
            browser.execute_script("arguments[0].scrollIntoView(false);", proj_box)

            proj_reason = browser.find_element_by_id('prreason')
            browser.execute_script("arguments[0].scrollIntoView(false);", proj_reason)
            proj_reason.clear()
            proj_reason.send_keys(Project_Name)
            # Click submit button
            submit_btn = browser.find_element_by_id('btn_submit')
            browser.execute_script("arguments[0].scrollIntoView(false);", submit_btn)
            submit_btn.click()
            # Record the result
            result = browser.find_element_by_tag_name('p')
            (success, fail) = re.findall(
                r'成功:(\d+)筆、失敗:(\d+)筆。', result.text)[0]
            if bool(int(success)) and not bool(int(fail)):
                rng.rows[idx+1].color = Green
            else:
                rng.rows[idx+1].color = Red
                err_msg = re.findall('新增失敗:(\w*)', result.text)
                # Set error message
                if err_msg:
                    rng.rows[idx+1][-1].value = err_msg
        except KeyboardInterrupt:
            break
        except:
            if rng.rows[idx+1].color not in [Green, Red]:
                rng.rows[idx+1].color = Red
    print()
        


# In[19]:


browser.quit()
workbook.save()


# In[20]:


# Check the number of error items
err_num = 0
for idx in dfidx:
    if rng.rows[idx+1].color == Red:
        err_num += 1

if err_num:
    print("共有 {} 筆錯誤項目，請於加班時數統計表手動更正".format(err_num))
else:
    print("無錯誤項目，三級加班專案開設已完成")
    workbook.close()
    
input("\n按任意鍵結束")

