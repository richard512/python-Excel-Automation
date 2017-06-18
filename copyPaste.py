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

excel.ScreenUpdating = True
#ws.SaveAs(wkbk3)
