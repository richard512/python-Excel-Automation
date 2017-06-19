import os
import win32com.client

excelFile = "time.xlsm"
macroToRun = "time.xlsm!Module1.Button1_Click"

if os.path.exists(excelFile):
    xl = win32com.client.Dispatch("Excel.Application")
    
    xl.Workbooks.Open(Filename=excelFile, ReadOnly=1)
    # remove ", ReadOnly=1" for file save
    
    xl.Application.Run(macroToRun)
    #xl.Application.Save() # uncomment for file save
    
    xl.Application.Quit() # Comment this out if your excel script closes
    del xl
