# -*- coding: utf-8 -*-
"""
Created on Wed Dec 23 12:09:09 2020

@author: s0046998
"""

import win32com.client
import os
import glob
# import time

# for WorkingFile in os.listdir("C:\\Users\\techy\\OneDrive\\Desktop\\python\\Data_files\\cust-geo"):

File_path=r"\\rgare.net\stlusop\Claims\Reporting\Lags-Individual\Test"
os.chdir(File_path)
xlapp = win32com.client.DispatchEx("Excel.Application")
names=[]
for FileList in glob.glob('ADJ - Individual Lag.xlsm'):
    if  not FileList.startswith('~$')  :
        names.append(FileList)
  
xlapp.Quit()

print(names)

xlapp = win32com.client.DispatchEx("Excel.Application")
for i in range(len(names)):
        #print(i)
        pw_str = 'locked'
        wb = xlapp.Workbooks.Open(File_path+"\\"+names[i])
        wb.Unprotect(pw_str)
        # wb.UnprotectSharing(pw_str)
        xlapp.DisplayAlerts = False
        wb.RefreshAll()
        print(names[i]) 
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close()
    
xlapp.Quit()
