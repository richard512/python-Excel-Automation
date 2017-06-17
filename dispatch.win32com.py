from win32com.client import Dispatch
excel = Dispatch("Excel.Application")
wb = excel.Workbooks.Append()
range = wb.Sheets[0].Range("A1")
range.Value = 32

'''
range.Activate                 range.Merge
range.AddComment               range.NavigateArrow
range.AdvancedFilter           range.NoteText
...
range.GetOffset                range.__repr__
range.GetResize                range.__setattr__
range.GetValue                 range.__str__
range.Get_Default              range.__unicode__
range.GoalSeek                 range._get_good_object_
range.Group                    range._get_good_single_object_
range.Insert                   range._oleobj_
range.InsertIndent             range._prop_map_get_
range.Item                     range._prop_map_put_
range.Justify                  range.coclass_clsid
range.ListNames                range.__class__
...
'''
