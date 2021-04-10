import xlsxwriter
import re
from datetime import datetime, timedelta
from Utill import OSInfo as oi  # remove

# filename =input('File name:')

filename = oi.get_basedir(__file__) + oi.getDelimit() + 'res/sample.txt'  # remove
output_filename = filename + '_output.xlsx'


def toTeraBytes(size, type):
    if type is None:
        return size

    if type == 'T':
        size = toTeraBytes(size, None)
    elif type == 'G':
        size = toTeraBytes(size / 1000, 'T')
    elif type == 'M':
        size = toTeraBytes(size / 1000, 'G')

    return size


workbook = xlsxwriter.Workbook(output_filename)
sheet = workbook.add_worksheet('Overview')
sheet.set_column('A:A', 15)
sheet.set_column('B:B', 30)
sheet.set_column('C:C', 10)

header_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'font_size': 12,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#dee6ef'
})
normal_format= workbook.add_format({
    'valign': 'vcenter',
    'border': 1,
})

row = 0
col = 0
sheet.write(row, col, 'Date', header_format)
sheet.write(row, col + 1, 'Environment', header_format)
sheet.write(row, col + 2, 'TB Usage', header_format)
sheet.write(row, col + 3, 'Month', header_format)
sheet.write(row, col + 4, 'Year', header_format)
sheet.write(row, col + 5, 'Product', header_format)
row += 1


date = datetime.now() - timedelta(days=int(datetime.now().strftime("%d")))

with open(filename) as _fh:
    for line in _fh:
        col=0
        data = re.split('\s+', line)
        client_size = toTeraBytes(int(data[0]), data[1])
        client = data[4].split('/')
        sheet.write(row, col, date.strftime('%d-%m-%Y'), normal_format)
        sheet.write(row, col + 1, client[-1], normal_format)
        sheet.write(row, col + 2, client_size, normal_format)
        sheet.write(row, col + 3, date.strftime('%m'), normal_format)
        sheet.write(row, col + 4, date.strftime('%Y'), normal_format)
        sheet.write(row, col + 5, 'FOX', normal_format)

        row += 1

workbook.close()
_fh.close()
