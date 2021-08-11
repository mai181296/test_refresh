# -*- coding: utf-8 -*-
"""
Created on Mon Aug  9 12:38:01 2021

@author: ngothimai

"""
import datetime
from pathlib import Path
import sys
# import os
import win32com.client as client

win32c = client.constants

def run_refresh(f_path: Path, f_name: str) -> list:

    filename = f_path / f_name

    # create excel object
    excel = client.gencache.EnsureDispatch('Excel.Application')

    # excel can be visible or not
    excel.Visible = True
    
    # try except for file / path
    try:
        wb = excel.Workbooks.Open(filename,UpdateLinks=False)
    
    except error as e:
        if e.excepinfo[5] == -2146827284:
            print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
            sys.exit(1)
        pass
      # Refesh All
    
    wb.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()
    
    bkTime = datetime.datetime.now().strftime('%Y-%m-%d')
    newFile = r'W:\OfficeShare\15.Product development\01. WEEKLY REPORT TO GĐK\Lead Pool and Funnel\Report\Report Leadgen RF file - store for ppt_TOPUPCOVID_v1'+bkTime+".xlsx"
    
    wb.SaveAs(newFile)
    wb.Close(False)
    

    excel.Quit()
    
if __name__ == "__main__":
    # file path
    f_path = Path(r'W:\OfficeShare\15.Product development\01. WEEKLY REPORT TO GĐK\Lead Pool and Funnel\Report') # file in current working directory    #f_path1 = Path(r'W:\OfficeShare\15.Product development\98. Report\Leadgen daily report\daily TLS CR report')
    # excel filename
    f_name='Report Leadgen RF file - store for ppt_TOPUPCOVID.xlsx'


    # function calls
    run_refresh(f_path, f_name)
