import os
import win32com.client

excelfile = "C:\Full Location\To\excelsheet.xlsm"
macroToRun = "excelsheet.xlsm!modulename.macroname"

if os.path.exists(excelfile):
    xl = win32com.client.Dispatch("Excel.Application")
    
    xl.Workbooks.Open(Filename=excelfile, ReadOnly=1)
    # remove ", ReadOnly=1" for file save
    
    xl.Application.Run(macroToRun)
    #xl.Application.Save() # uncomment for file save
    
    xl.Application.Quit() # Comment this out if your excel script closes
    del xl
