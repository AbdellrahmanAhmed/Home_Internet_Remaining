from datetime import date, datetime

from openpyxl import load_workbook
from openpyxl.styles import Alignment

wb = load_workbook('Home Internet Remaining.xlsx')

cell = 1

today = date.today()
now = datetime.now()

month = today.strftime("%B")
time = now.strftime("%H:%M:%S")
todayN = today.strftime("%A")

ws = wb.active
cellValue = ws[f'A${cell}'].value

remaining = input('Remaining: ')

while cellValue:
    cell = cell + 1
    cellValue = ws[f'A${cell}'].value

ws[f'A${cell}'].value = today
ws[f'B${cell}'].value = month
ws[f'C${cell}'].value = todayN
ws[f'D${cell}'].value = time
ws[f'E${cell}'].value = remaining

cellList = ['A', 'B', 'C', 'D', 'E']

for cellC in cellList:
    currentCell = ws[f'${cellC}${cell}']
    currentCell.alignment = Alignment(horizontal='center')

print('Last Remaining:  ', ws[f'E${cell - 1}'].value)

wb.save('Home Internet Remaining.xlsx')
