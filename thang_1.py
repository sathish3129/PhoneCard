import xlsxwriter
import re
from datetime import datetime, timedelta
from Utill import OSInfo as oi  # remove

# filename =input('File name:')

filename = oi.get_basedir(__file__) + oi.getDelimit() + 'res/sample2.txt'  # remove
output_filename = oi.get_basedir(__file__) + oi.getDelimit() + 'res/output2.xlsx'

client_name = ''
client_data = {}

with open(filename) as _fh:
    for line in _fh:

        if line.strip() == '' or re.match(r'^\s*(hadoop|\-)', line) or re.match(r'^\s*\(\d+\s+rows\)', line):
            continue
        elif line.startswith('Client'):
            cl_line = line.split(':')
            client_name = cl_line[1].strip()
            client_data[client_name] = {
                'DATA':[]
            }
        else:
            cl_data = line.split('|')
            cl_data = list(map(lambda x: x.strip(), cl_data))
            if cl_data[2] not in client_data[client_name]:
                client_data[client_name][cl_data[2]] = {'size': 0}
            client_data[client_name][cl_data[2]]['size'] += float(cl_data[1])
            client_data[client_name]['DATA'].append(
                {
                    'name': cl_data[0],
                    'cpu': cl_data[1],
                    'bucket': cl_data[2],
                }
            )
_fh.close()

print(client_data)

workbook = xlsxwriter.Workbook(output_filename)
sheet = workbook.add_worksheet('Overview')
sheet_summary = {
    'FOX' : workbook.add_worksheet('FOX'),
    'IRIS': workbook.add_worksheet('IRIS'),
    'FOX_IRIS': workbook.add_worksheet('FOX_IRIS')
}

header_format = workbook.add_format({
    'bold': 1,
    'font_size': 12,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#dee6ef'
})

client_header = workbook.add_format({
    'bold': 1,
    'font_size': 12,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#FFFF00'
})

normal_format= workbook.add_format({
    'valign': 'vcenter',
})

# overview
row = 0
summary_row = {
    'FOX' : {'row':0, 'col':0},
    'IRIS': {'row':0, 'col':0},
    'FOX_IRIS': {'row':0, 'col':0}
}


for i in sheet_summary.keys():
    sheet_summary[i].write(summary_row[i]['row'], summary_row[i]['col'], 'Client', header_format)
    sheet_summary[i].write(summary_row[i]['row'], summary_row[i]['col'] + 1, 'Client', header_format)
    summary_row[i]['row'] += 1


col = 0
sheet.write(row, col, 'Client', header_format)
sheet.write(row, col + 1, 'Job Type', header_format)
sheet.write(row, col + 2, 'CUP Time', header_format)
sheet.write(row, col + 3, 'Bucket', header_format)
row += 2

for _name in client_data.keys():
    col=0
    sheet.write(row, col, _name, client_header)

    if len(client_data[_name]['DATA']) == 0:
        continue

    row+=1
    for k in client_data[_name].keys():

        if k == 'DATA':
            for dat in client_data[_name]['DATA']:
                col=0
                sheet.write(row, col+1, dat['name'], normal_format)
                sheet.write(row, col+2, dat['cpu'], normal_format)
                sheet.write(row, col+3, dat['bucket'], normal_format)
                row += 1
            row += 1
        else:
            k_key = k.replace('/', '_')
            summary_row[k_key]['col'] = 0
            print(f'k:{k}')
            print('data:{0}'.format(client_data[_name][k]))
            sheet_summary[k_key].write(summary_row[k_key]['row'], summary_row[k_key]['col'], _name, normal_format)
            sheet_summary[k_key].write(summary_row[k_key]['row'], summary_row[k_key]['col'] + 1, client_data[_name][k]['size'],
                                       normal_format)
            summary_row[k_key]['row'] += 1


workbook.close()
