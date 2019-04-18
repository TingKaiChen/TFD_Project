#!/usr/bin/env python
# coding: utf-8

# In[1]:


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from .三級加班_ui import Ui_MainWindow
import sys, os
import xlwings as xw
import configparser as cfgparser
import ctypes
import datetime as dt

# In[ ]:


class MainWindowUIClass( Ui_MainWindow ):
    def __init__( self ):
        '''Initialize the super class
        '''
        super().__init__()
        self.app = xw.App(add_book = False, visible = False)
        self.app.display_alerts = False

        self._filename = None
        self._filesht = None
        self._year = None
        self._month = None
        self._account = None
        self._password = None

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
        yearlist = [str(i) for i in range(dt.date.today().year-20, dt.date.today().year+20)]
        monthlist = [str(i) for i in range(1, 13)]
        self.year_box.clear()
        self.year_box.addItems(yearlist)
        self.year_box.setCurrentText(str(dt.date.today().year))
        self.month_box.clear()
        self.month_box.addItems(monthlist)
        self.month_box.setCurrentText(str(dt.date.today().month))
        self.account.setText(self.cfg.get('DEFAULT', 'account', fallback = ''))
        self.password.setText(self.cfg.get('DEFAULT', 'password', fallback = ''))

    def filename_click(self):
        parpath = self.cfg.get('DEFAULT', 'filename_parpath', fallback = '')
        fileName, _ = QFileDialog.getOpenFileName(
            None,"選擇三級加班時數表", parpath, "All Files (*)")
        if fileName:
            self.filename.setText(fileName)
            self._filename = fileName
            self.app.books.api.Open(fileName, UpdateLinks=False)
            self.wb = self.app.books[-1]
            self.filesht_cb.clear()
            self.filesht_cb.addItems([sht.name for sht in self.wb.sheets])
            self._filesht = self.filesht_cb.currentText()
        pass

    def filesht_cb_click(self):
        self._filesht = self.filesht_cb.currentText()
    
    def year_click(self):
        self._year = self.year_box.currentText()

    def month_click(self):
        self._month = self.month_box.currentText()

    def start_click(self):
        self._account = self.account.text()
        self._password = self.password.text()

        self.MW.close() 
        for wb in self.app.books:
            wb.close()
        self.app.quit()
        self.app.kill()

        self._start_flag = True

        if (self._filename and self._filesht and self._year and self._month and \
            self._account and self._password):
            # Update the configuration
            self.cfg['DEFAULT']['filename_parpath'] = \
                os.path.abspath(os.path.join(self._filename, os.pardir))
            self.cfg['DEFAULT']['year'] = self._year
            self.cfg['DEFAULT']['month'] = self._month
            self.cfg['DEFAULT']['account'] = self._account
            self.cfg['DEFAULT']['password'] = self._password
            with open(self._config_filename, 'w') as configfile:
                self.cfg.write(configfile)
        
    def getParam(self):
        if self._start_flag:
            return (self._filename, self._filesht, int(self._year), \
                int(self._month), self._account, self._password)
        else:
            return(None, None, None, None, None, None)
    
    def clear(self):
        if not self._start_flag:
            for wb in self.app.books:
                wb.close()
            self.app.quit()
            self.app.kill()