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

from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

from time import sleep
from dateutil import relativedelta
import os, sys, re, csv, time, io, filecmp, math
import datetime as dt
import pandas as pd
import configparser as cfgparser


# In[2]:

cfg = cfgparser.ConfigParser()
cfg.read('C:\\Users\\TFD\\主管差假搜尋\\config.ini')
account = cfg.get('DEFAULT', 'account', fallback = '')
password = cfg.get('DEFAULT', 'password', fallback = '')


# In[3]:


# Frame switch function
def switch_to_iframe(browser):
    browser.switch_to.default_content()
    browser.switch_to.frame("topmost")
    browser.switch_to.frame("iframe")
def switch_to_topmost(browser):
    browser.switch_to.default_content()
    browser.switch_to.frame("topmost")
def switch_to_accountframe(browser):
    browser.switch_to.default_content()
    browser.switch_to.frame("topmost")
    browser.switch_to.frame(browser.find_element_by_css_selector("body > center > table:nth-child(1) > tbody > tr:nth-child(1) > td > table > tbody > tr > td:nth-child(2) > div > iframe"))


# In[4]:


class MyDocument():
    def __init__(self, doc_name):
        self.doc = Document(doc_name)
        
        regex = '(\w+)\t(\w+) (\w+-\w+-\w+)\(\w+\) (\w+:\w+) ~ (\w+-\w+-\w+)\(\w+\) (\w+:\w+)\n\t事由: (\w+)'
        self.re_comp = re.compile(regex)
        
        # The time limit of the leave period
        td = dt.date.today()
        self.start_lmt = dt.datetime(td.year, td.month, td.day, 8, 30)
        self.mid1_lmt = dt.datetime(td.year, td.month, td.day, 12, 30)
        self.mid2_lmt = dt.datetime(td.year, td.month, td.day, 13, 30)
        self.end_lmt = dt.datetime(td.year, td.month, td.day, 17, 30)
        
        # Paragraph style: 'Normal'
        styles = self.doc.styles
        style = styles['Normal']
        style.font.name = u'標楷體'
        style._element.rPr.rFonts.set(qn('w:eastAsia'), u'標楷體')
        style.font.size = Pt(28)
        style.font.bold = True
        # New paragraph style: 'New'
        if 'NewParagraph' not in styles:
            newpg = styles.add_style('NewParagraph', WD_STYLE_TYPE.PARAGRAPH)
            newpg.font.name = '標楷體'
            newpg._element.rPr.rFonts.set(qn('w:eastAsia'), u'標楷體')
            newpg.font.size = Pt(12)
            newpg.font.bold = True
        # Big font style: 'BigFont'
        if 'BigFont' not in styles:
            bigfont = styles.add_style('BigFont', WD_STYLE_TYPE.CHARACTER)
            bigfont.font.name = '標楷體'
            bigfont._element.rPr.rFonts.set(qn('w:eastAsia'), u'標楷體')
            bigfont.font.size = Pt(28)
        # Small font style: 'SmallFont'
        if 'SmallFont' not in styles:
            smallfont = styles.add_style('SmallFont', WD_STYLE_TYPE.CHARACTER)
            smallfont.font.name = '標楷體'
            smallfont._element.rPr.rFonts.set(qn('w:eastAsia'), u'標楷體')
            smallfont.font.size = Pt(12)
        
    def setTitleDate(self, date):
        run = self.doc.tables[0].cell(0, 0).paragraphs[0].runs[0]
        run.text = '{}年{}月{}日 主管差假'.format(date.year - 1911, 
                                                 date.month, 
                                                 date.day)
        
    def addParagraph(self):
        newpg = self.doc.tables[0].cell(1, 0).add_paragraph(style = 'NewParagraph')
        newpg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    def writeBigFont(self, string):
        pg = self.doc.tables[0].cell(1, 0).paragraphs[-1]
        pg.add_run(string, style = 'BigFont')
        self.addParagraph()
        
    def writeSmallFont(self, string):
        pg = self.doc.tables[0].cell(1, 0).paragraphs[-1]
        pg.add_run(string, style = 'SmallFont')
        self.addParagraph()
        
    def addLeave(self, raw_string):
        leave, reason = self.leaveStringProcess(raw_string)
        self.writeBigFont(leave)
        self.writeSmallFont('(' + reason + ')')
        
    def leaveStringProcess(self, string):
        '''
        string: ex. 游專委家懿\t休假 108-05-03(二) 09:30 ~ 108-05-04(二) 12:30\n\t事由: 國旅卡休假
        return ('游專委家懿9-12時休假', '國旅卡休假')
        '''
        (officername, leavetype, start_date, start_time, 
           end_date, end_time, reason) = self.re_comp.findall(string)[0]
        # Leave period 
        start_date = self.changeYear(start_date)
        end_date = self.changeYear(end_date)
        start_dt = start_date + ' ' + start_time
        end_dt = end_date + ' ' + end_time
        onduty = dt.datetime.strptime(start_dt, '%Y-%m-%d %H:%M')
        offduty = dt.datetime.strptime(end_dt,  '%Y-%m-%d %H:%M')
        
        # Using '上午', '下午', or hour numbers
        StartFromMorning = False
        StartFromNoon = False
        EndInNoon = False
        EndInAfternoon = False
        
        if (onduty - self.start_lmt).total_seconds() <= 0:
            StartFromMorning = True
        elif onduty.hour in [12, 13]:
            StartFromNoon = True
        if(offduty - self.end_lmt).total_seconds() >= 0:
            EndInAfternoon = True
        elif offduty.hour in [12, 13]:
            EndInNoon = True
            
        if StartFromMorning and EndInAfternoon:
            date_str = ''
        elif StartFromMorning and EndInNoon:
            date_str = '上午'
        elif StartFromNoon and EndInAfternoon:
            date_str = '下午'
        else:
            date_str = ''
            if StartFromMorning:
                date_str += '8'
            else:
                date_str += str(onduty.hour)

            date_str += '-'

            if EndInAfternoon:
                date_str += '17'
            else:
                date_str += str(offduty.hour)

            date_str += '時'
            
        # Change '主任秘書' into '主秘'
        if '主任秘書' in officername:
            officername = officername.replace('主任秘書', '主秘')
        # Change leavetype
        if ('其他假' in leavetype) or ('補休' in leavetype):
            leavetype = '補休'
            
        leave_str = officername + date_str + leavetype
        return leave_str, reason
    
    def changeYear(self, date_str):
        '''
        date_str: ex. 108-10-05
        return: ex. 2019-10-05
        '''
        dt_list = date_str.split('-')
        dt_list[0] = str(int(dt_list[0]) + 1911)
        return ('-'.join(dt_list))
        
    def save(self, save_name):
        self.doc.save(save_name)


# In[5]:


# 請假清單
leavelist_1 = []   #已批核請假
leavelist_2 = []   #未批核


# In[6]:


# Unit list
OfficeList = ["資通作業科", "減災規劃科", "整備應變科", "局長室", "第一副局長室", "災害搶救科", 
              "火災調查科", "綜合企劃科", "緊急救護科", "督察室", "第三副局長室", "主任秘書室", 
              "秘書室", "人事室", "政風室", "會計室", "訓練中心", "救災救護指揮中心", "火災預防科", 
              "第二副局長室"]


# In[7]:


# File output directory
file_dir = r'Z:\\02獎懲待遇股\\湘苹\\主管差假\\'
detail_dir = r'Z:\\02獎懲待遇股\\湘苹\\主管差假\\主管請假清單\\'
download_dir = 'Z:\\02獎懲待遇股\\湘苹\\主管差假\\主管請假清單\\downloads\\'


# In[8]:


# Month list
monthlist = {'一': '1', '二': '2', '三': '3', '四': '4', '五': '5', 
             '六': '6', '七': '7', '八': '8', '九': '9', '十': '10', 
             '十一': '11', '十二': '12'}


# In[9]:


# Todo List
# windowless
# multiple leaves
# password


# In[10]:


print("正在執行主管差假搜尋，請勿關閉此視窗")


# In[11]:


option = webdriver.ChromeOptions()
# option.add_argument('--headless')
# option.add_argument('--window-size=1280,800"')
prefs = {'download.default_directory' : download_dir}
option.add_experimental_option('prefs', prefs)
if getattr(sys, 'frozen', False) :
    # running in a bundle
    chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
    browser = webdriver.Chrome(chromedriver_path, options=option)
else:
    # executed as a simple script, the driver should be in "PATH"
    browser = webdriver.Chrome(options=option)
#開啟google首頁
browser.get('https://webitr.gov.taipei/WebITR/')


# In[12]:


# Login
element_user = browser.find_element_by_id("userName")
element_user.send_keys(account)
element_pw = browser.find_element_by_id("login_key")
element_pw.send_keys(password)
element_pw.send_keys(Keys.RETURN)


# In[13]:


# 主管名單更新
titlename_dict = {}
# Click "差勤管理/制度管理/基本資料維護"
switch_to_topmost(browser)
wait = WebDriverWait(browser, 10)
## 差勤管理
li1 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > a")
ActionChains(browser).move_to_element_with_offset(li1, 10, -10).move_to_element_with_offset(li1, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(1) > a")))
## 制度管理
li2 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li2, 10, -10).move_to_element_with_offset(li2, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(1) > ul > li:nth-child(3) > a")))
## 基本資料維護
li3 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(1) > ul > li:nth-child(3) > a")
ActionChains(browser).move_to_element_with_offset(li3, 10, -10).move_to_element_with_offset(li3, 10, 10).perform()
li3.click()
## 移出左方列表使浮出列表消失
li4 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li4, 10, -10).perform()
element = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(1) > a")))

switch_to_iframe(browser)
unit_procnum = 0
for office_name in OfficeList:
    unit = Select(browser.find_element_by_id("unit"))
    unit.select_by_visible_text(office_name)
    namelist = browser.find_elements_by_name("persons")
    if len(namelist) == 0:
    ## 無人處室
        continue
    ## 局長室全選
    elif office_name == "局長室":
        browser.find_element_by_id("personsAll").click()
    ## 其餘科室
    else:
        namelist[0].click()
    exebtn = browser.find_element_by_name("exeBtn")
    browser.execute_script("arguments[0].scrollIntoView(false);", exebtn)
    exebtn.click()

    rows = browser.find_elements_by_class_name("p4_list_mouse")
    for row in rows:
        name = row.find_elements_by_tag_name("td")[2].text
        title = row.find_elements_by_tag_name("td")[1].text
        if title == "局長室":
            title = "局長"
        elif title == "專門委員":
            title = "專委"
        elif title == "簡任技正":
            title = "簡技"
        title_name = name[0]+title+name[1:]
        titlename_dict[name] = title_name
    
    unit_procnum += 1
    print("\r(1/5) 主管名單更新: %3d%%"%(unit_procnum/20*100), end = '')
print("\r(1/5) 主管名單更新:   完成")


# In[14]:


# 請假資料報表查詢
# Click "差勤管理/請假管理/請假資料報表"
switch_to_topmost(browser)
wait = WebDriverWait(browser, 10)
## 差勤管理
li1 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > a")
ActionChains(browser).move_to_element_with_offset(li1, 10, -10).move_to_element_with_offset(li1, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > a")))
## 請假管理
li2 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > a")
ActionChains(browser).move_to_element_with_offset(li2, 10, -10).move_to_element_with_offset(li2, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > ul > li:nth-child(3) > a")))
## 請假資料報表
li3 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > ul > li:nth-child(3) > a")
ActionChains(browser).move_to_element_with_offset(li3, 10, -10).move_to_element_with_offset(li3, 10, 10).perform()
li3.click()
## 移出左方列表使浮出列表消失
li4 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li4, 10, -10).perform()
element = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(5) > ul > li:nth-child(3) > a")))


# In[15]:


def chooseDay(chosen_date, browser):
    # Back to that month
    cal_str = browser.find_elements_by_class_name('title')
    cal_yr_mon = re.findall(r'民國\ (\d+)\ 年\ (.*)月', cal_str[-1].text)[0]
    cal_yr = str(int(cal_yr_mon[0])+1911)
    cal_yr_mon = cal_yr + '-' + monthlist[cal_yr_mon[1]]
    mon1 = dt.datetime.strptime(cal_yr_mon, '%Y-%m')
    mon2 = dt.datetime.strptime(
        chosen_date.strftime("%Y-%m"), '%Y-%m')
    pressnum = relativedelta.relativedelta(mon1, mon2).months

    bar = browser.find_elements_by_class_name('headrow')
    if pressnum > 0:
        monthbtn = bar[-1].find_elements_by_class_name('button')[1]
    else:
        monthbtn = bar[-1].find_elements_by_class_name('button')[3]
    browser.execute_script("arguments[0].scrollIntoView(false);", monthbtn)
    for i in range(abs(pressnum)):
        monthbtn.click()
        
    # Click the chosen day of that month
    calendar = browser.find_elements_by_class_name('calendar')
    days = calendar[-1].find_elements_by_class_name('day')
    days_in_month = []
    for day in days:
        if 'name' not in day.get_attribute('class') and            'othermonth' not in day.get_attribute('class'):
            days_in_month.append(day)
    sel_day = days_in_month[chosen_date.day - 1]
    browser.execute_script("arguments[0].scrollIntoView(false);", sel_day)
    sel_day.click()
 


# In[16]:


switch_to_iframe(browser)
# Search start day
datestrbtn = browser.find_element_by_id("begintime_bt")
browser.execute_script("arguments[0].scrollIntoView(false);", datestrbtn)
datestrbtn.click()
start_day = dt.datetime.now()
chooseDay(start_day, browser)

# Search end day
dateendbtn = browser.find_element_by_id("endtime_bt")
browser.execute_script("arguments[0].scrollIntoView(false);", dateendbtn)
dateendbtn.click()
end_day = dt.datetime.now()
chooseDay(end_day, browser)

select = Select(browser.find_element_by_id("qryType"))
select.select_by_visible_text('查詢各單位主管 ')

# Select all officers and leave types
allmembox = browser.find_element_by_id('personsAll')
browser.execute_script("arguments[0].scrollIntoView(false);", allmembox)
if not allmembox.is_selected():
    allmembox.click()
allleavebox = browser.find_element_by_id('allLeavetype')
browser.execute_script("arguments[0].scrollIntoView(false);", allleavebox)
if not allleavebox.is_selected():
    allleavebox.click()
# Export excel
excelbtn = browser.find_element_by_id('excelBtn')
browser.execute_script("arguments[0].scrollIntoView(false);", excelbtn)
excelbtn.click()

sleep(3)

# Select 簡技 & 專委
select = Select(browser.find_element_by_id("qryType"))
select.select_by_visible_text('依單位查詢')
select = Select(browser.find_element_by_id("unit"))
select.select_by_visible_text('局長室')

# Select all officers and leave types
allmembox = browser.find_element_by_id('personsAll')
browser.execute_script("arguments[0].scrollIntoView(false);", allmembox)
if not allmembox.is_selected():
    allmembox.click()
allleavebox = browser.find_element_by_id('allLeavetype')
browser.execute_script("arguments[0].scrollIntoView(false);", allleavebox)
if not allleavebox.is_selected():
    allleavebox.click()
    
# Deselect the box of director
directorbox = browser.find_elements_by_name('persons')[0]
browser.execute_script("arguments[0].scrollIntoView(false);", directorbox)
if directorbox.is_selected():
    directorbox.click()

# Export excel
excelbtn = browser.find_element_by_id('excelBtn')
browser.execute_script("arguments[0].scrollIntoView(false);", excelbtn)
excelbtn.click()

# Wait until download complete
while(1):
    sleep(1)
    if len(os.listdir(download_dir)) > 1:
        break
        
# Loop through 2 download files
for fn in os.listdir(download_dir):
    if '~' not in fn:
        df = pd.read_excel(download_dir + fn)
        for idx, row in df.iterrows():
            if (type(row['Unnamed: 1']) == str) & (row['Unnamed: 1'] != '姓名'):
                name = titlename_dict[row['Unnamed: 1']]  
                lvtype = row['Unnamed: 3']   
                str_date = row['Unnamed: 4'] 
                str_time = row['Unnamed: 5'] 
                end_date = row['Unnamed: 6'] 
                end_time = row['Unnamed: 7'] 
                reason = row['Unnamed: 10']  
                lv_str = (name + '\t' + lvtype + ' ' + str_date + ' ' + 
                          str_time + ' ~ ' + end_date + ' ' + end_time +
                          '\n\t事由: ' + reason)
                leavelist_1.append(lv_str)
        os.remove(download_dir + fn)
        
print("\r(2/5) 已批核查詢:   完成")


# In[17]:


# 待批核假查詢
# Click "系統維護/表單進度查詢"
switch_to_topmost(browser)
wait = WebDriverWait(browser, 10)
## 系統維護
li1 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(9) > a")
ActionChains(browser).move_to_element_with_offset(li1, 10, -10).move_to_element_with_offset(li1, 10, 10).perform()
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(9) > ul > li:nth-child(1) > a")))
## 表單進度查詢
li2 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(9) > ul > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li2, 10, -10).move_to_element_with_offset(li2, 10, 10).perform()
li2.click()
## 移出左方列表使浮出列表消失
li3 = browser.find_element_by_css_selector("#MenuBar1 > li:nth-child(1) > a")
ActionChains(browser).move_to_element_with_offset(li3, 10, -10).perform()
element = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "#MenuBar1 > li:nth-child(9) > ul > li:nth-child(1) > a")))

switch_to_iframe(browser)
print("\r(4/5) 待批核查詢:   ", end = '')
## 選取全部部門
depart = Select(browser.find_element_by_id("depart"))
depart.select_by_visible_text("全部部門")
## 選取請假開始日期
datestrbtn = browser.find_element_by_id("begintime_dc_bt")
browser.execute_script("arguments[0].scrollIntoView(false);", datestrbtn)
datestrbtn.click()
today = browser.find_elements_by_class_name("today")[-1]
browser.execute_script("arguments[0].scrollIntoView(false);", today)
today.click()
##選取請假結束日期
dateendbtn = browser.find_element_by_id("endtime_dc_bt")
browser.execute_script("arguments[0].scrollIntoView(false);", dateendbtn)
dateendbtn.click()
today = browser.find_elements_by_class_name("today")
today = browser.find_elements_by_class_name("today")[-1]
browser.execute_script("arguments[0].scrollIntoView(false);", today)
today.click()
## 點選查詢
querybtn = browser.find_element_by_name("selecta")
browser.execute_script("arguments[0].scrollIntoView(false);", querybtn)
querybtn.click()

## 從搜尋結果中擷取長官差假
table = browser.find_element_by_css_selector("#div2 > table")
rows = table.find_elements_by_class_name("stripeMe")
depart = Select(browser.find_element_by_id("depart"))
for row in rows:
    is_officer = False
    unit = row.find_elements_by_tag_name("td")[0].text
    name = row.find_elements_by_tag_name("td")[1].text
    leavetype = row.find_elements_by_tag_name("td")[2].text
    leavedate = row.find_elements_by_tag_name("td")[3].text
    if name in titlename_dict.keys():
        # Find the reason of the leave
        main_window = browser.window_handles[0]
        leave_link = row.find_element_by_tag_name('a')
        browser.execute_script("arguments[0].scrollIntoView(false);", leave_link)
        leave_link.send_keys(Keys.CONTROL + Keys.RETURN)
        # Edit in new tab
        pop_window = browser.window_handles[1]
        browser.switch_to.window(pop_window)
        element = wait.until(EC.presence_of_element_located((By.TAG_NAME, "tr")))
        for row in browser.find_elements_by_tag_name('tr'):
            # Check whether the person is an officer
            if '表單申請人' in row.text:
                mem_str = row.find_elements_by_tag_name('td')[1].text
                title = re.findall(':\((\w+)\)\((\w+)\)', mem_str)[0][0]
                level = re.findall(':\((\w+)\)\((\w+)\)', mem_str)[0][1]
                if (title in ['局長', '副局長', '專門委員', '主任秘書', '主任', '科長']) or                    ((title == '技正') and ('簡任' in level)):
                        is_officer = True
                else:
                    break
            # Leave period
            if '差假(或加班)時間' in row.text:
                lv_dt = row.find_elements_by_tag_name('td')[1].text
                lv_dt = re.findall(
                    '(\d+-\d+-\d+\(\w+\) \d+:\d+ ~ \d+-\d+-\d+\(\w+\) \d+:\d+)', lv_dt)[0]
            elif '事由' in row.text:
            # Leave reason
               reason = row.find_elements_by_tag_name('td')[1].text
            else:
                continue
                
        # Close the edit tab
        browser.close()
        browser.switch_to.window(main_window)
        switch_to_iframe(browser)
        if is_officer:
            leave_str = (titlename_dict[name]+"\t"+leavetype+" "+ lv_dt +
                         '\n\t事由: ' + reason)
            leavelist_2.append(leave_str)
        
print("\r(4/5) 待批核查詢:   完成")


# In[18]:


# Leave summary
print("\r(5/5) 輸出主管差假清單:   ", end = '')
year = str(int(time.strftime("%Y", time.localtime()))-1911)
date = time.strftime("%m%d", time.localtime())
filename = detail_dir + year + date + "主管差假清單"
if os.path.isfile(filename+"_上午.txt"):
    filetag = "_下午.txt"
else:
    filetag = "_上午.txt"
with io.open(filename+filetag, 'w', encoding='utf8') as outfile:
    outfile.write("===已批核===\n")
    outfile.write('\n'.join(leavelist_1))
    outfile.write("\n=======未批核========\n")
    outfile.write('\n'.join(leavelist_2))

noupdate = False
if filetag == "_下午.txt" and    filecmp.cmp(filename+"_上午.txt", filename+"_下午.txt"):
        with io.open(filename+filetag, 'w', encoding='utf8') as outfile:
            outfile.write("無更改")
        noupdate = True
    
print("\r(5/5) 輸出主管差假清單:   完成")
print("\n\n===主管差假查詢完成===")
if noupdate:
    print("下午無主管差假更新")
else:
    print("主管差假已更新於\""+filename+filetag+"\"")
browser.quit()
input("\n按任意鍵結束")


# In[19]:


today_str = dt.datetime.strftime(dt.date.today(), '%Y-%m-%d')
today_list = today_str.split('-')
today_list[0] = str(int(today_list[0]) - 1911)
today_str = ''.join(today_list)
if dt.datetime.today().hour < 13:
    word_tag = '_上午'
else:
    word_tag = '_下午'
wordfilename = file_dir + today_str + '主管差假' + word_tag + '.docx'
if (not os.path.isfile(wordfilename)) or    (os.path.isfile(wordfilename) and not noupdate):
    # Generate Word file
    doc = MyDocument('template.docx')
    doc.setTitleDate(dt.date.today())
    for raw_string in leavelist_1:
        doc.addLeave(raw_string)
    for raw_string in leavelist_2:
        doc.addLeave(raw_string)
    doc.save(wordfilename)


# In[20]:


# print("===已批核===")
# for leave in leavelist_1:
#     print(leave)

# print("=======未批核========")
# for leave in leavelist_2:
#     print(leave)

