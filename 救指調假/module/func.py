import re
import datetime as dt
from dateutil import relativedelta
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC



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
    browser.execute_script("arguments[0].scrollIntoView(false);", monthbtn)
    for i in range(abs(pressnum)):
        monthbtn.click()
        
    # Click the chosen day of that month
    calendar = browser.find_elements_by_class_name('calendar')
    days = calendar[-1].find_elements_by_class_name('day')
    days_in_month = []
    for day in days:
        if 'name' not in day.get_attribute('class') and \
           'othermonth' not in day.get_attribute('class'):
            days_in_month.append(day)
    sel_day = days_in_month[chosen_date.day - 1]
    browser.execute_script("arguments[0].scrollIntoView(false);", sel_day)
    sel_day.click()
    
def updateContinueLeave(lvlist, elselvlist, lvsum, elselvsum, 
                        cont_start, cont_end, clv_namelist, browser):
    wait = WebDriverWait(browser, 10)
    if (len(lvlist) + len(elselvlist) > 1) & (len(lvlist) >= 1):
    # Find a contiuous leaves issue
        total_sum = (cont_end - cont_start).days
        extra = total_sum - sum(lvsum) - sum(elselvsum)
        # Edit extra leave days
        main_window = browser.window_handles[0]
        modify_btn = lvlist[0].find_element_by_id('modifyLink')
        browser.execute_script("arguments[0].scrollIntoView(false);", modify_btn)
        modify_btn.send_keys(Keys.CONTROL + Keys.RETURN)
        # Edit in new tab
        pop_window = browser.window_handles[1]
        browser.switch_to.window(pop_window)
        element = wait.until(EC.visibility_of_element_located((By.ID, "povdays")))
        povdays = browser.find_element_by_id('povdays')
        browser.execute_script("arguments[0].scrollIntoView(false);", povdays)
        povdays.clear()
        povdays.send_keys(str(float(extra + lvsum[0])))
        
        reason = browser.find_element_by_id('reason')
        browser.execute_script("arguments[0].scrollIntoView(false);", reason)
        cstart_str = (cont_start - dt.timedelta(days = 1)).strftime('%m%d')
        cend_str = (cont_end + dt.timedelta(days = 1)).strftime('%m%d')
        add_reason = '(' + cstart_str + '-' + cend_str + '為連續休假，計為' + str(extra + lvsum[0]) + '日)'
        old_reason = reason.get_attribute('value')
        cleave_label = re.findall('(\(\d+-\d+為連續休假，計為\d+日\))', old_reason)
        if cleave_label:
            reason_str = old_reason.replace(cleave_label[0], add_reason)
        else:
            reason_str = old_reason + add_reason
        reason.clear()
        reason.send_keys(reason_str)
            
        # Store the information of continuous leave
        mem_name = browser.find_element_by_id('person').text
        clv_namelist.append(mem_name + '\t' + add_reason)
        
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
    # Clear variables and lists
    lvlist.clear()
    elselvlist.clear()
    lvsum.clear()
    elselvsum.clear()     

def str2Dates(leave_str):
    '''
    argument
        leave_str: leave period string in format "MMDD/hhmm-MMDD/hhmm"
    return
        datetime pair (start, end) for a valid period
        or
        None for an invalid period
    '''
    (str_str, str_end) = re.findall(
        r'(\d+)-(\d+-\d+)\(\w\) (\d+:\d+)', leave_str)
    try:
        str_str = ''.join(
            str(int(str_str[0]) + 1911) + '-' + str_str[1] + ' ' + str_str[2] )
        dt_str = dt.datetime.strptime(str_str, '%Y-%m-%d %H:%M')
        str_end = ''.join(
            str(int(str_end[0]) + 1911) + '-' + str_end[1] + ' ' + str_end[2] )
        dt_end = dt.datetime.strptime(str_end, '%Y-%m-%d %H:%M')
    except:
        return None

    return (dt_str, dt_end)

def isValidContinuousLeave(leave_str, cont_end):
    (dt_str, dt_end) = str2Dates(leave_str)
    dtdiff = relativedelta.relativedelta(dt_str, cont_end)
    if (dt_str.hour == 8) & (dt_str.minute == 30) & \
       (dt_end.hour == 8) & (dt_end.minute == 30) & \
       (dtdiff.days == 1 ) & (dtdiff.hours == 0) & (dtdiff.minutes == 0):
        return True
    else:
        return False

def leavePeriodDays(leave_str):
    (dt_str, dt_end) = str2Dates(leave_str)
    return (dt_end - dt_str).days