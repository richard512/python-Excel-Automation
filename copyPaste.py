'''
Copies A2:C4 from worksheet1.xlsx to new row(s) starting at the end of worksheet2.xlsx, then saves it as worksheet3.xlsx
'''

from win32com.client import Dispatch
wkbk1 = "C:\\path\\to\\worksheet1.xlsx"
wkbk2 = "C:\\path\\to\\worksheet2.xlsx"
wkbk3 = "C:\\path\\to\\worksheet3.xlsx"
excel = Dispatch("Excel.Application")
excel.Visible = 1
excel.ScreenUpdating = False

# copy from source
source = excel.Workbooks.Open(wkbk1)
excel.Range("A2:C4").Select()
excel.Selection.Copy()

# paste (appended) to target
target = excel.Workbooks.Open(wkbk2)

ws = target.sheets(1)
used = ws.UsedRange
maxrow = used.Row + used.Rows.Count - 1
maxcol = used.Column + used.Columns.Count - 1
newrow = maxrow + 1

destrange = "A"+str(newrow) #+":C9"
print(destrange)

excel.Range(destrange).Select()
excel.Selection.PasteSpecial(Paste=-4122)

# Excel requires the user to press enter after paste()
# It says "select destination and press enter or choose paste"
# This simulates pressing of the enter key:
shell = Dispatch("WScript.Shell")
shell.SendKeys("{ENTER}", 0)

excel.ScreenUpdating = True
ws.SaveAs(wkbk3)
