import os
import win32com.client

excelFile = "C:\Full Location\To\excelsheet.xlsm"
macroToRun = "excelsheet.xlsm!modulename.macroname"

if os.path.exists(excelFile):
    xl = win32com.client.Dispatch("Excel.Application")
    
    xl.Workbooks.Open(Filename=excelFile, ReadOnly=1)
    # remove ", ReadOnly=1" for file save
    
    xl.Application.Run(macroToRun)
    #xl.Application.Save() # uncomment for file save
    
    xl.Application.Quit() # Comment this out if your excel script closes
    del xl
