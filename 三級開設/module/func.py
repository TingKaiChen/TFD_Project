import datetime as dt
import re
from dateutil import relativedelta
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By

# Selenium frame switch function
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

def worktime(wt_str, Year):
    '''
    argument
        wt_str: work period string in format "MMDD/hhmm-MMDD/hhmm"
        Year: string of the year
    return
        datetime pair (start, end) for a valid period
        or
        None for an invalid period
    '''
    try:
        [strdt, enddt] = wt_str.split('-')
        [strdate, strtime] = list(strdt.split('/'))
        [enddate, endtime] = list(enddt.split('/'))
    except ValueError:
        return (None, None)
    
    # Check basic format (string length)
    if len(strdate + strtime + enddate + endtime) != 16:
        return (None, None)
    
    strdt = Year + strdt
    enddt = Year + enddt
    # date & time validation
    try:
        datetime_str = dt.datetime.strptime(strdt, '%Y%m%d/%H%M')
        datetime_end = dt.datetime.strptime(enddt, '%Y%m%d/%H%M')
        return (datetime_str, datetime_end)
    except ValueError:
        return (None, None)

def chooseDay(chosen_date, browser):
    # Month list
    monthlist = {'一': '1', '二': '2', '三': '3', '四': '4', '五': '5', 
                 '六': '6', '七': '7', '八': '8', '九': '9', '十': '10', 
                 '十一': '11', '十二': '12'}
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
    for i in range(abs(pressnum)):
        browser.execute_script("arguments[0].scrollIntoView(false);", monthbtn)
        monthbtn.click()
        
    # Click the chosen day of that month
    calendar = browser.find_elements_by_class_name('calendar')
    days = calendar[-1].find_elements_by_class_name('day')
    days_in_month = []
    for day in days:
        if 'name' not in day.get_attribute('class') and \
           'othermonth' not in day.get_attribute('class'):
            days_in_month.append(day)
    browser.execute_script("arguments[0].scrollIntoView(false);", days_in_month[chosen_date.day - 1])
    days_in_month[chosen_date.day - 1].click()
    
def deselect(browser):
    wait = WebDriverWait(browser, 10)
    element = wait.until(EC.visibility_of_element_located((By.ID, "selAll")))
    selall_btn = browser.find_element_by_id('selAll')
    browser.execute_script("arguments[0].scrollIntoView(false);", selall_btn)
    selall_btn.click()
    selall_btn.click()

def enter_wo_period(dt1, dt2, wo_hour, browser):
    wait = WebDriverWait(browser, 10)
    element = wait.until(EC.visibility_of_element_located((By.ID, "in_praddd_xx_bt")))
    # Choose start date and end date
    start_btn = browser.find_element_by_id('in_praddd_xx_bt')
    browser.execute_script("arguments[0].scrollIntoView(false);", start_btn)
    start_btn.click()
    chooseDay(dt1, browser)
    end_btn = browser.find_element_by_id('in_pradde_xx_bt')
    browser.execute_script("arguments[0].scrollIntoView(false);", end_btn)
    end_btn.click()
    chooseDay(dt2, browser)
    # Choose start time
    stime_h = Select(browser.find_element_by_id('in_prstime_hh'))
    stime_h.select_by_visible_text(dt1.strftime('%H'))
    stime_m = Select(browser.find_element_by_id('in_prstime_mm'))
    stime_m.select_by_visible_text(dt1.strftime('%M'))
    # Choose end time
    etime_h = Select(browser.find_element_by_id('in_pretime_hh'))
    etime_h.select_by_visible_text(dt2.strftime('%H'))
    etime_m = Select(browser.find_element_by_id('in_pretime_mm'))
    etime_m.select_by_visible_text(dt2.strftime('%M'))
    # Edit work-over hour
    wo_input = browser.find_element_by_id('praddh_hh')
    browser.execute_script("arguments[0].scrollIntoView(false);", wo_input)
    wo_input.clear()
    wo_input.send_keys(wo_hour)
