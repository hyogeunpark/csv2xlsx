import os
import glob
import csv
import pip
import codecs 

is_not_installed = pip.main(['list', 'xlsxwriter'])
if is_not_installed == 1:
    pip.main(['install', '--user', 'xlsxwriter'])

from xlsxwriter.workbook import Workbook

for csvfile in glob.glob(os.path.join('.', '*.csv')):
    workbook = Workbook(csvfile[:-4] + '.xlsx')
    worksheet = workbook.add_worksheet()
    with open(csvfile, 'rt', encoding='utf8') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()
