from win32com.client import Dispatch
wkbk1 = "worksheet1.xlsx"
wkbk2 = "worksheet2.xlsx"
excel = Dispatch("Excel.Application")
excel.Visible = 1
source = excel.Workbooks.Open(wkbk1)
excel.Range("A1:A3").Select()
excel.Selection.Copy()
copy = excel.Workbooks.Open(wkbk2)
excel.Range("A1:A3").Select()
excel.Selection.PasteSpecial(Paste=-4163)
