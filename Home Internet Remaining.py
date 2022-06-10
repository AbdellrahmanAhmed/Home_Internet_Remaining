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


def writeCell(cell0):
    ws[f'C${cell0}'].value = todayN
    ws[f'B${cell0}'].value = month
    ws[f'A${cell0}'].value = today
    ws[f'D${cell0}'].value = time
    ws[f'E${cell0}'].value = remaining


ws = wb.active
cellValue = ws[f'A${cell}'].value

remaining = input('Remaining: ')

while cellValue:
    cell = cell + 1
    cellValue = ws[f'A${cell}'].value

writeCell(cell)

cellList = ['A', 'B', 'C', 'D', 'E']

for cellC in cellList:
    currentCell = ws[f'${cellC}${cell}']
    currentCell.alignment = Alignment(horizontal='center')

# print(cell)

lastRemaining = ws[f'E${cell - 1}'].value
(used) = round(float(lastRemaining) - float(remaining), 2)
# print(f'Last Used {round(used, 2)} GB')

last_Remaining = ws[f'E${cell - 1}'].value

tt = str(ws[f'D${cell - 1}'].value)

bTime = datetime.strptime(tt, "%H:%M:%S")
btTime = datetime.strptime(time, "%H:%M:%S") - bTime

print(f'''
Last Remaining: {last_Remaining}   Last time  {tt}
Used {used} GB      Duration  {btTime}
''')
wb.save('Home Internet Remaining.xlsx')
