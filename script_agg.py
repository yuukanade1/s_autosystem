import pyautogui as pg
import time
import openpyxl
import datetime
# excelbookパス C:\Users\竜馬\OneDrive\share\集計転記\syuukeiin.xlsx

dt_now = datetime.datetime.now()
print(dt_now.strftime('%Y%m%d')) 

wb = openpyxl.load_workbook(r'C:\Users\竜馬\OneDrive\share\集計転記\syuukeiin.xlsx')
ws = wb.worksheets[0]
maxClm = ws.max_column
maxRow = ws.max_row
print('maxcolumn : ' + str(maxClm))
print('maxrow    : ' + str(maxRow))

order = []
for clm in range(2, maxClm + 1, 1):
    for row in reversed(range(1,maxRow + 1)):
        if ws.cell(row, clm).value != None:
            print(str(clm) + '列の最終行 : ' + str(row))
            break

