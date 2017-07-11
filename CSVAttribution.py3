import xlrd
import csv
from datetime import datetime
import sys
import os.path

DAILY_EXCEL_FILE_PATH = 'Python Attribution File.xlsx'
MASTER_FILE_PATH = 'Master Attribution CSV.csv'


def convert_path_to_csv(path):
    return '.'.join(path.split('.')[:-1]) + '.csv'

def excel_to_csv(path):
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    csv_path = convert_path_to_csv(path)
    csv_file = open(csv_path, 'w')
    wr = csv.writer(csv_file)
    date_tuple = xlrd.xldate_as_tuple(sh.row_values(0)[-1], wb.datemode)
    date = datetime(*date_tuple).strftime('%m/%d/%Y')
    date_fields = [date for i in range(sh.nrows-1)]
    date_fields = ['Date'] + date_fields
    for rownum in range(sh.nrows):
        if rownum == 0:
            wr.writerow([date_fields[rownum]] + sh.row_values(rownum)[:-1] + ['Value'])
        else:
            wr.writerow([date_fields[rownum]] + sh.row_values(rownum))
    csv_file.close()


def add_to_master(master_path, csv_path):
    master_lines = [line.strip() for line in open(master_path) if line.strip()]
    csv_lines = [line.strip() for line in open(csv_path) if line.strip()]
    if master_lines:
        csv_lines = csv_lines[1:]
    master_lines += csv_lines
    with open(master_path, 'w') as out:
        out.write('\n'.join(master_lines))

if len(sys.argv) == 3:
    if not os.path.isfile(sys.argv[1]):
        print ("Error: File '" + sys.argv[1] + "' does not exist")
    else:
        INPUT_FILE = sys.argv[1]
        OUTPUT_FILE = sys.argv[2]
        print ("Input file = " + INPUT_FILE)
        print ("Output file = " + OUTPUT_FILE)
        excel_to_csv(INPUT_FILE)
        add_to_master(OUTPUT_FILE, convert_path_to_csv(INPUT_FILE))
else:
    print ("Error: " + sys.argv[0] + " must have 2 file arguments")