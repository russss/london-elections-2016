from openpyxl import load_workbook
import re
import csv
from collections import defaultdict

wb = load_workbook('results.xlsx')

candidates = defaultdict(list)
ws = wb.worksheets[-1]

for row in ws.rows[1:]:
    if row[0].value is not None:
        candidates[row[4].value.strip()].append(((row[1].value + ' ' + (row[2].value or '')).strip(),
                                                 row[3].value))


def lookup_candidates(row, competition):
    result = []
    for val in row:
        res = re.search(r'(Candidate|Party) ([0-9]+)', val)
        if res:
            result.append(candidates[competition][int(res.group(2)) - 1][0])
        else:
            result.append(val)
    return result


def write(filename, header, result):
    print("Writing %s..." % filename)
    with open("./data/" + filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(header)
        writer.writerows(result)


# Mayor, first pref:

result = []
ws = wb.worksheets[0]

header = ['Constituency', 'Ward'] + lookup_candidates([cell.value for cell in
                                                       ws.rows[2][2:13] + ws.rows[2][15:20]],
                                                      'London Mayor')
for ws in wb.worksheets:
    if ws.title == 'Keys':
        break
    constituency = ws.title
    for row in ws.rows[3:]:
        if row[0].value == 'Key':
            break
        if row[0].value is not None:
            result.append([constituency] + [cell.value for cell in row[1:13] + row[15:20]])

write('london-mayor-first-preference.csv', header, result)

# Mayor, second pref:

result = []
ws = wb.worksheets[0]

header = ['Constituency', 'Ward'] + lookup_candidates([cell.value for cell in
                                                       ws.rows[2][21:33] + ws.rows[2][34:37]],
                                                      'London Mayor')
for ws in wb.worksheets:
    if ws.title == 'Keys':
        break
    constituency = ws.title
    for row in ws.rows[3:]:
        if row[0].value == 'Key':
            break
        if row[0].value is not None:
            result.append([constituency] + [row[1].value] + [cell.value for cell in row[21:33] + row[34:37]])

write('london-mayor-second-preference.csv', header, result)


# London-wide assembly member:

result = []
ws = wb.worksheets[0]

header = ['Constituency', 'Ward'] + lookup_candidates([cell.value for cell in
                                                       ws.rows[2][38:50] + ws.rows[2][51:56]],
                                                      'London-wide Assembly')
for ws in wb.worksheets:
    if ws.title == 'Keys':
        break
    constituency = ws.title
    for row in ws.rows[3:]:
        if row[0].value == 'Key':
            break
        if row[0].value is not None:
            result.append([constituency] + [row[1].value] + [cell.value for cell in row[38:50] + row[51:56]])

write('london-member.csv', header, result)


# Constituency assembly member:

for ws in wb.worksheets:
    if ws.title == 'Keys':
        break
    result = []

    constituency = ws.title
    num_candidates = len([cell for cell in ws.rows[2][57:] if cell.value is not None]) - 7
    header = ['Ward'] + lookup_candidates([cell.value for cell in
                                           ws.rows[2][57:57 + num_candidates] +
                                           ws.rows[2][57 + num_candidates + 1:57 + num_candidates + 6]],
                                           constituency)
    for row in ws.rows[3:]:
        if row[0].value == 'Key':
            break
        if row[0].value is not None:
            result.append([row[1].value] + [cell.value for cell in row[57:57 + num_candidates] +
                                            row[57 + num_candidates + 1:57 + num_candidates + 6]])

    filename = "constituency-member-%s.csv" % constituency.replace('&', 'and').replace(' ', '-').lower()

    write(filename, header, result)
