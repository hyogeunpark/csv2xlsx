import os
import csv
import pip
from glob import glob

# install xlsxwriter
is_not_installed = pip.main(['show', 'xlsxwriter'])
if is_not_installed == 1:
    pip.main(['install', '--user', 'xlsxwriter'])

from xlsxwriter.workbook import Workbook

# csv to xlsx
def all_csv_to_xlsx():
    for csvfile in glob(os.path.join('.', '*.csv')):
        workbook = Workbook(csvfile[:-4] + '.xlsx')
        sheet = workbook.add_worksheet()
        with open(csvfile, 'rt', encoding='utf8') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    sheet.write(r, c, col)
        workbook.close()

def csv_to_xlsx(file_path):
    csvfile = os.path.join('.', file_path)
    workbook = Workbook(csvfile[:-4] + '.xlsx')
    sheet = workbook.add_worksheet()
    with open(csvfile, 'rt', encoding='utf8') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                sheet.write(r, c, col)
    workbook.close()

if __name__ == '__main__':
    print('[csv to xlsx]')
    val = input('Input file path(If empty, convert all csv file) : ');
    if str(val).strip() != '':
        csv_to_xlsx(val)
    else:
        all_csv_to_xlsx()
    print('Completed!!')
