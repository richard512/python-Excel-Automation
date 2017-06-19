import xlrd

print (xlrd.__file__)

#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
 
    # print number of sheets
    print (book.nsheets)
 
    # print sheet names
    print (book.sheet_names())
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
 
    # read a row
    print (first_sheet.row_values(0))
 
    # read a cell
    cell = first_sheet.cell(0,0)
    print (cell)
    print (cell.value)
 
    # read a row slice
    print (first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=2))
 
#----------------------------------------------------------------------
if __name__ == "__main__":
    path = "worksheet1.xlsx"
    open_file(path)
