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
File_path=r"U:\Claims\Reporting\Lags-Individual\2022"#change the year in Dec
#File_path=r"M:\python\individual_lag"

os.chdir(File_path)
names=[]

xlapp = win32com.client.DispatchEx("Excel.Application")

def get_path():
    
     os.chdir(File_path)
     xlapp = win32com.client.DispatchEx("Excel.Application")
     
     for FileList in glob.glob('*.xlsm'):
        if  not FileList.startswith('~$')  :
           names.append(File_path+"\\"+FileList)
     xlapp.Quit()
     return names
   
def Refresh_file(names):
     for FileList in names:
        pw_str = 'locked'
        wb = xlapp.Workbooks.Open(FileList)
        wb.Unprotect(pw_str)
        # wb.UnprotectSharing(pw_str)
        xlapp.DisplayAlerts = False
        wb.RefreshAll()   
        print(FileList)
        wb.RefreshAll()
     
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close()
        #time.sleep(20)
     xlapp.Quit()
     return ("All files complete...")

print(Refresh_file(get_path()))
