#!/usr/bin/env python
# coding: utf-8

# In[1]:


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot, QDate
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QInputDialog, QLineEdit, QFileDialog, QDialog, QLabel
from .救指調假_ui import Ui_MainWindow
import sys, os
import configparser as cfgparser
import ctypes
import datetime as dt

# In[ ]:


class MainWindowUIClass( Ui_MainWindow ):
    def __init__( self ):
        '''Initialize the super class
        '''
        super().__init__()
        self._year = None
        self._month = None
        self._account = None
        self._password = None
        self._start_dt = None
        self._start_dt_str = None
        self._end_dt = None
        self._end_dt_str = None

        self._start_flag = False
        self._config_filename = './module/config.ini'

        self.cfg = cfgparser.ConfigParser()
        self.cfg.read(self._config_filename)
        
    def setupUi( self, MW ):
        ''' Setup the UI of the super class, and add here code
        that relates to the way we want our UI to operate.
        '''
        super().setupUi( MW )
        self.MW = MW
        self.MW.setWindowIcon(QtGui.QIcon('./GUI/icon.png'))
        # Set different PID to show the icon
        # ref: https://stackoverflow.com/questions/1551605/how-to-set-applications-taskbar-icon-in-windows-7
        myappid = 'myapp' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

        # Set default year and month
        yearlist = [str(i) for i in range(
            dt.date.today().year-1911-20, dt.date.today().year-1911+20)]
        monthlist = [str(i) for i in range(1, 13)]
        self.year_box.clear()
        self.year_box.addItems(yearlist)
        self.year_box.setCurrentText(str(dt.date.today().year - 1911))
        self.month_box.clear()
        self.month_box.addItems(monthlist)
        self.month_box.setCurrentText(str(dt.date.today().month))
        self.account.setText(self.cfg.get('DEFAULT', 'account', fallback = ''))
        self.password.setText(self.cfg.get('DEFAULT', 'password', fallback = ''))

        default_dt_str = dt.date.today().strftime('%Y-%m-%d')
        self._start_dt_str = self.cfg.get(
            'DEFAULT', 'start_dt', fallback = default_dt_str)
        self._end_dt_str = self.cfg.get(
            'DEFAULT', 'end_dt', fallback = default_dt_str)
        qdate_start = QDate.fromString(self._start_dt_str, 'yyyy-MM-dd')
        qdate_end = QDate.fromString(self._end_dt_str, 'yyyy-MM-dd')
        self.start_dt_cal.setSelectedDate(qdate_start)
        self.end_dt_cal.setSelectedDate(qdate_end)
        self.start_dt_label.setText(self._start_dt_str)
        self.end_dt_label.setText(self._end_dt_str)

    def year_click(self):
        self._year = self.year_box.currentText()

    def month_click(self):
        self._month = self.month_box.currentText()

    def start_dt(self):
        self._start_dt_str = self.start_dt_cal.selectedDate().toString('yyyy-MM-dd')
        self.start_dt_label.setText(self._start_dt_str)
        self._start_dt = dt.datetime.strptime(self._start_dt_str, '%Y-%m-%d')

    def end_dt(self):
        self._end_dt_str = self.end_dt_cal.selectedDate().toString('yyyy-MM-dd')
        self.end_dt_label.setText(self._end_dt_str)
        self._end_dt = dt.datetime.strptime(self._end_dt_str, '%Y-%m-%d')

    def start_click(self):
        self._account = self.account.text()
        self._password = self.password.text()

        self.MW.close() 

        self._start_flag = True

        if (self._year and self._month and self._account and self._password):
            # Update the configuration
            self.cfg['DEFAULT']['year'] = self._year
            self.cfg['DEFAULT']['month'] = self._month
            self.cfg['DEFAULT']['account'] = self._account
            self.cfg['DEFAULT']['password'] = self._password
            self.cfg['DEFAULT']['start_dt'] = self._start_dt_str
            self.cfg['DEFAULT']['end_dt'] = self._end_dt_str
            with open(self._config_filename, 'w') as configfile:
                self.cfg.write(configfile)
        
    def getParam(self):
        if self._start_flag:
            return (int(self._year), int(self._month), self._account, self._password, 
                    self._start_dt, self._end_dt)
        else:
            return(None, None, None, None, None, None)
    