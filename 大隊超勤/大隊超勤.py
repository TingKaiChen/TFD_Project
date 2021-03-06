#!/usr/bin/env python
# coding: utf-8

# In[1]:


from time import sleep
from dateutil import relativedelta
import os, sys, re, csv, time, io, filecmp
import datetime as dt
import pandas as pd
import xlwings as xw
from xlwings.constants import AutoFillType
import numpy as np
from numpy import linalg as LA
import shutil

from GUI.大隊超勤_app import *
from PyQt5.QtWidgets import QApplication, QMainWindow

from module.SearchName import SearchNameObj


# In[2]:


# Read in the unit list
unitlist = {}
unit_table = pd.read_excel('單位名稱轉換表.xlsx')
for index, row in unit_table.iterrows():
    unitlist[row[0]] = row[1] 


# In[3]:


# Open GUI to select/read filenames
app = QtWidgets.QApplication(sys.argv)
MainWindow = QtWidgets.QMainWindow()
ui = MainWindowUIClass()
ui.setupUi(MainWindow)
MainWindow.show()
app.exec_()

(dir_path, pm_file, pm_shtname, namecorrection_file, 
     namecorrection_sht) = ui.getParam()
if not (dir_path and pm_file and pm_shtname and namecorrection_file and 
        namecorrection_sht):
    ui.clear()
    sys.exit()

pm_path = os.path.split(pm_file)[0] + '/'
pm_filename = os.path.split(pm_file)[1]
# Global variables
error_lists = {}
real_sums = {}
take_sums = {}


# In[4]:


# Supress "update linked data source" warnings
xlapp = xw.App(add_book = False, visible = False)
xlapp.display_alerts = False


# In[5]:


# Find the range and column indices of the payment file
if '-cp' not in pm_filename:
    cp_filename = pm_filename.replace('.xlsx', '-cp.xlsx')
else:
    cp_filename = pm_filename
if not os.path.isfile(pm_path + cp_filename):
    shutil.copy(pm_path + pm_filename, pm_path + cp_filename)


# In[6]:


SNO = SearchNameObj(pm_path + cp_filename, pm_shtname)
SNO.correctName(namecorrection_file, namecorrection_sht)


# In[7]:


error_lists = {}
real_sums = {}
take_sums = {}

# Process counter
filenum = len(os.listdir(dir_path))
if 'summary.xlsx' in os.listdir(dir_path):
    filenum -= 1
if 'error_log.txt' in os.listdir(dir_path):
    filenum -= 1
processnum = 1
info_str = '\r處理進度: ({}/' + str(filenum) + ') {}    ' 

for filename in os.listdir(dir_path):
    if ('summary' in filename) or ('error_log' in filename):
        continue
    if '~' not in filename:
        # Open the file
        xlapp.books.api.Open(dir_path + filename, UpdateLinks=False)
        wb = xlapp.books[-1]
        sheet1 = wb.sheets['本局超勤統計表']
        sheet2 = wb.sheets['本局印領清冊']

        try:
            # Process of 本局超勤統計表
            sheet1.activate()
            rng1 = sheet1.range('A1').current_region
            
            # Replace a string type '-' into integer 0
            for cell in rng1:
                if cell.value == '-':
                    cell.value = 0

            unit_found = False
            for i in range(len(rng1.rows)):
                # Extract the unit name
                if not unit_found and rng1.rows[i].value:
                    ulist = rng1.rows[i].value
                    unit_str = ''.join([u for u in ulist if type(u) == str])
                    if '北市政府' in unit_str:
                        unit_found = True

                        u = unit_str
                        u = re.sub(r'\s+', '', u)
                        u = re.sub(r'大隊部', '', u)
                        u = re.sub(r'救災救護', '', u)
                        u = re.findall(r'局(第\w+隊)', u)[0]
                        ul = re.findall(r'(\w+?隊)', u)
                        un = ul[-1]
                        unit_name = unitlist[un]
                        print(info_str.format(processnum, un), end = '')
                        processnum += 1
                if '實際超勤時數' in rng1.rows[i].value:
                    # Find the column index
                    cidx_sum1 = rng1.rows[i].value.index('實際超勤時數')
                if '支領超勤時數' in rng1.rows[i].value:
                    # Find the column index
                    cidx_sum2 = rng1.rows[i].value.index('支領超勤時數')
                if '姓名' in rng1.rows[i].value:
                    # Find the column index
                    cidx_name = rng1.rows[i].value.index('姓名')
                    # Find the start row index
                    rng1[i, cidx_name].select()
                    ridx_start = wb.selection.shape[0]+i
            # Find the end row index
            for i in range(ridx_start, len(rng1.rows)):
                if rng1[i, cidx_name].value != None:
                    ridx_end = i + 1

            # Summation formulas
            sum1_str = '=SUM(' + rng1[ridx_start:ridx_end, cidx_sum1].address + ')'
            sum2_str = '=SUM(' + rng1[ridx_start:ridx_end, cidx_sum2].address + ')'

            # Enter the formulas into the proper cells
            if '實際時數總和' in rng1.rows[0].value:
                sum1_label = rng1.api.Find('實際時數總和')
                sum1_cell = sheet1.range((sum1_label.row + 1, sum1_label.column))
                sum2_label = rng1.api.Find('支領時數總和')
                sum2_cell = sheet1.range((sum2_label.row + 1, sum2_label.column))
            else:
                sum1_label = sheet1.range((1, rng1.shape[1] + 1))
                sum1_label.value = '實際時數總和'
                sum1_label.autofit()
                sum1_cell = sheet1.range((sum1_label.row + 1, sum1_label.column))
                sum2_label = sheet1.range((1, sum1_label.column + 1))
                sum2_label.value = '支領時數總和'
                sum2_label.autofit()
                sum2_cell = sheet1.range((sum2_label.row + 1, sum2_label.column))
            
            sum1_cell.value = sum1_str
            real_sums[un] = int(sum1_cell.value)
            sum2_cell.value = sum2_str
            take_sums[un] = int(sum2_cell.value)
        except Exception as e:
            print('Error in ' + unit_name)
            print(type(e), end = ': ')
            print(e)
            wb.save()
            wb.close()
            continue
        except:
            pass
                    
    if '_OK' in filename:
        wb.save()
        wb.close()
    else:
        try:
            sheet2.activate()
            rng2 = sheet2.range('A1').current_region
            rng2 = sheet2.range('A1', (rng2.shape[0], 16))
            
            # Replace a string type '-' into integer 0
            for cell in rng2:
                if cell.value == '-':
                    cell.value = 0

            # Insert columns and set column index
            for i in range(len(rng2.rows)):
                if '薪俸' in rng2.rows[i].value:
                    for cell in rng2.rows[i]:
                        if type(cell.value) == str:
                            cell.value = re.sub('\s+', '', cell.value)
                    # Find the start row index
                    cidx_name = rng2.rows[i].value.index('姓名')
                    cidx_title = rng2.rows[i].value.index('職稱')
                    rng2[i, cidx_name].select()
                    ridx_start = wb.selection.shape[0] + i   

                    # Check whether columns are inserted
                    cidx1 = rng2.rows[i].value.index('薪俸')
                    cidx2 = rng2.rows[i].value.index('專業加給')        
                    if (cidx2 - cidx1) >= 3:
                        cidx_1 = rng2.rows[i].value.index('薪俸')
                        cidx_2 = rng2.rows[i].value.index('專業加給')
                        cidx_3 = rng2.rows[i].value.index('主管加給')
                        break
                    # Insert columns
                    rng2 = sheet2.range('A1').current_region
                    rng2 = sheet2.range('A1', (rng2.shape[0], 16))
                    cidx_1 = rng2.rows[i].value.index('薪俸')
                    rng2[:, cidx_1 + 1].api.Insert()
                    rng2[:, cidx_1 + 1].api.Insert()
                    rng2 = sheet2.range('A1').current_region
                    rng2 = sheet2.range('A1', (rng2.shape[0], 16))
                    cidx_2 = rng2.rows[i].value.index('專業加給')
                    rng2[:, cidx_2 + 1].api.Insert()
                    rng2[:, cidx_2 + 1].api.Insert()
                    rng2 = sheet2.range('A1').current_region
                    rng2 = sheet2.range('A1', (rng2.shape[0], 16))
                    cidx_3 = rng2.rows[i].value.index('主管加給')
                    rng2[:, cidx_3 + 1].api.Insert()
                    rng2[:, cidx_3 + 1].api.Insert()
                    rng2 = sheet2.range('A1').current_region
                    break

            # Find the end row index
            for i in range(ridx_start, len(rng2.rows)):
                if (rng2[i, cidx_title].value != None) and                    (rng2[i, cidx_name].value != None):
                    ridx_end = i + 1       

            # Fill in the ID
            rng2 = sheet2.range('A1').current_region
            rng2 = sheet2.range('A1', (rng2.shape[0], 16))
            search_rng = SNO.getDefaultRangeStr()
            cidx_pm1, cidx_pm2, cidx_pm3 = SNO.getSearchColumnIndices()
            # Formulas string
            f1_str = '=VLOOKUP({},\'{}[{}]{}\'!{},{},0)'.format(
                        rng2[ridx_start, cidx_name].address.replace('$', ''), 
                        pm_path, 
                        cp_filename, 
                        pm_shtname, 
                        search_rng, 
                        cidx_pm1)
            f2_str = '=VLOOKUP({},\'{}[{}]{}\'!{},{},0)'.format(
                        rng2[ridx_start, cidx_name].address.replace('$', ''), 
                        pm_path, 
                        cp_filename, 
                        pm_shtname, 
                        search_rng, 
                        cidx_pm2)
            f3_str = '=VLOOKUP({},\'{}[{}]{}\'!{},{},0)'.format(
                        rng2[ridx_start, cidx_name].address.replace('$', ''), 
                        pm_path, 
                        cp_filename, 
                        pm_shtname, 
                        search_rng, 
                        cidx_pm3)

            # Enter formulas into 1st row
            rng2[ridx_start, cidx_1 + 1].formula = f1_str
            rng2[ridx_start, cidx_1 + 2].value = ('={}-{}'.format(
                rng2[ridx_start, cidx_1 + 1].address.replace('$', ''), 
                rng2[ridx_start, cidx_1].address.replace('$', '')))
            rng2[ridx_start, cidx_2 + 1].formula = f2_str
            rng2[ridx_start, cidx_2 + 2].value = ('={}-{}'.format(
                rng2[ridx_start, cidx_2 + 1].address.replace('$', ''), 
                rng2[ridx_start, cidx_2].address.replace('$', '')))
            rng2[ridx_start, cidx_3 + 1].formula = f3_str
            rng2[ridx_start, cidx_3 + 2].value = ('={}-{}'.format(
                rng2[ridx_start, cidx_3 + 1].address.replace('$', ''), 
                rng2[ridx_start, cidx_3].address.replace('$', '')))

            # Autofill formulas into all cells
            rng2[ridx_start, cidx_1 + 1].api.AutoFill(
                rng2[ridx_start:ridx_end, cidx_1 + 1].api, AutoFillType.xlFillDefault)
            rng2[ridx_start, cidx_1 + 2].api.AutoFill(
                rng2[ridx_start:ridx_end, cidx_1 + 2].api, AutoFillType.xlFillDefault)
            rng2[ridx_start, cidx_2 + 1].api.AutoFill(
                rng2[ridx_start:ridx_end, cidx_2 + 1].api, AutoFillType.xlFillDefault)
            rng2[ridx_start, cidx_2 + 2].api.AutoFill(
                rng2[ridx_start:ridx_end, cidx_2 + 2].api, AutoFillType.xlFillDefault)
            rng2[ridx_start, cidx_3 + 1].api.AutoFill(
                rng2[ridx_start:ridx_end, cidx_3 + 1].api, AutoFillType.xlFillDefault)
            rng2[ridx_start, cidx_3 + 2].api.AutoFill(
                rng2[ridx_start:ridx_end, cidx_3 + 2].api, AutoFillType.xlFillDefault)
            
            # Check for duplicated names problem
            for i in range(ridx_start, ridx_end):
                pm1 = rng2[i, cidx_1].value
                pm2 = rng2[i, cidx_2].value
                pm3 = rng2[i, cidx_3].value
                if not pm1:
                    pm1 = 0
                if not pm2:
                    pm2 = 0
                if not pm3:
                    pm3 = 0
                search_rng = SNO.getDupNameRangeStr(rng2[i, cidx_name].value,
                                                    [pm1, pm2, pm3])
                # If duplicated names exist
                if search_rng:
                    # Formulas string
                    f1_str = '=VLOOKUP({},\'{}[{}]{}\'!{},{},0)'.format(
                                rng2[i, cidx_name].address.replace('$', ''), 
                                pm_path, 
                                cp_filename, 
                                pm_shtname, 
                                search_rng, 
                                cidx_pm1)
                    f2_str = '=VLOOKUP({},\'{}[{}]{}\'!{},{},0)'.format(
                                rng2[i, cidx_name].address.replace('$', ''), 
                                pm_path, 
                                cp_filename, 
                                pm_shtname, 
                                search_rng, 
                                cidx_pm2)
                    f3_str = '=VLOOKUP({},\'{}[{}]{}\'!{},{},0)'.format(
                                rng2[i, cidx_name].address.replace('$', ''), 
                                pm_path, 
                                cp_filename, 
                                pm_shtname, 
                                search_rng, 
                                cidx_pm3)
                    # Enter formulas into 1st row
                    rng2[i, cidx_1 + 1].formula = f1_str
                    rng2[i, cidx_2 + 1].formula = f2_str
                    rng2[i, cidx_3 + 1].formula = f3_str

            Yellow = (255, 255, 0)
            error_list = []
            # Check for inconsistence
            for i in range(ridx_start, ridx_end):
                err_str = ''
                err_list = []
                if rng2[i, cidx_1 + 2].value != 0:
                    if not rng2[i, cidx_1].value:
                        rng2[i, cidx_1].value = 0
                    rng2[i, cidx_1:(cidx_1 + 3)].color = Yellow
                    err_list.append('\t薪俸:\t' + str(int(rng2[i, cidx_1].value)) + 
                                '/' + str(int(rng2[i, cidx_1 + 1].value)))
                else:
                    rng2[i, cidx_1:(cidx_1 + 3)].color = None
                if rng2[i, cidx_2 + 2].value != 0:
                    if not rng2[i, cidx_2].value:
                        rng2[i, cidx_2].value = 0
                    rng2[i, cidx_2:(cidx_2 + 3)].color = Yellow
                    err_list.append('\t專業:\t' + str(int(rng2[i, cidx_2].value)) + 
                                '/' + str(int(rng2[i, cidx_2 + 1].value)))
                else:
                    rng2[i, cidx_2:(cidx_2 + 3)].color = None
                if rng2[i, cidx_3 + 2].value != 0:
                    if not rng2[i, cidx_3].value:
                        rng2[i, cidx_3].value = 0
                    rng2[i, cidx_3:(cidx_3 + 3)].color = Yellow
                    err_list.append('\t主管:\t' + str(int(rng2[i, cidx_3].value)) + 
                                '/' + str(int(rng2[i, cidx_3 + 1].value)))
                else:
                    rng2[i, cidx_3:(cidx_3 + 3)].color = None


                if err_list:
                    err_str = rng2[i, cidx_name].value + '\n'.join(err_list)
                    error_list.append(err_str)

            error_lists[unit_name] = error_list
            # Rename the file if no error exists
            if not error_list:
#                 # Copy & paste the value to fix it from the formula
#                 for _ridx in range(ridx_start, ridx_end):
#                     for _cidx in [cidx_1 + 1, cidx_2 + 1, cidx_3 + 1]:
#                         rng2[_ridx, _cidx].value = rng2[_ridx, _cidx].value
                # Close and rename the file
                wb.save()
                wb.close()
                os.rename(dir_path + filename, 
                          dir_path + filename.replace('.x', '_OK.x'))
            else:
                wb.save()
                wb.close()
            
        except KeyboardInterrupt:
            break
        except Exception as e:
            print('Error in ' + unit_name)
            print(type(e), end = ': ')
            print(e)
            wb.save()
            wb.close()
print('\n')


# In[8]:


## Create a summary file when all files are correctly done
# Check whether all files are done correctly
error_exist = False
for err in list(error_lists.values()):
    if err:
        error_exist = True
if not error_exist:
    print('表格核對完成，正在輸出時數統計...', end = '')
    # Remove the error log file if it exists
    if os.path.isfile(dir_path + 'error_log.txt'):
        os.remove(dir_path + 'error_log.txt')
    # Create a summary file
    sum_wb = xlapp.books.add()
    sum_filename = sum_wb.name + '.xlsx'
    sum_sht = sum_wb.sheets[-1]
    sum_rng = sum_sht.range((1, 1), (len(real_sums) + 3, 3))
    # Header
    sum_rng.rows[0].value = ['分隊', '實際時數', '支領時數']
    unit_names = list(real_sums.keys())
    for i in range(len(real_sums)):
        un = unit_names[i]
        sum_rng.rows[i + 1].value = [un, real_sums[un], take_sums[un]]
    # Summary
    sum_rng.rows[-1].value = ['total', 
                              '=SUM(' + sum_rng[1:len(real_sums) + 1, 1].address + ')', 
                              '=SUM(' + sum_rng[1:len(real_sums) + 1, 2].address + ')']

    sum_wb.save()
    sum_wb.close()
    # Move and rename the summary file
    shutil.move(sum_filename, dir_path + 'summary.xlsx')
    print('完成')
    print('時數統計表格儲存於「{}」資料夾中的 summary.xlsx'.format(dir_path))
    print('請手動註記於紙本超勤資料上')
else:
    print('正在輸出錯誤資訊:\n')
    # Write a error log file
    with io.open(dir_path + 'error_log.txt', 'w', encoding = 'utf8') as outf:
        for un in error_lists:
            if not error_lists[un]:
                continue
            else:
                print(un)
                outf.write(un + '\n')
            for err in error_lists[un]:
                print(err.replace('-2146826246', 'NaN'))
                outf.write(err.replace('-2146826246', 'NaN') + '\n')
            outf.write('\n')
    print('\n錯誤資訊儲存於「{}」資料夾中的 error_log.txt'.format(dir_path))
    print('請手動核對/更正錯誤資訊')


# In[9]:


for wb in xlapp.books:
    wb.save()
    wb.close()
xlapp.quit()
xlapp.kill()
SNO.quit()

input("\n按任意鍵結束")

