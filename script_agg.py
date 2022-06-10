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

for j in range(2, maxClm + 1):
    for i in reversed(range(1,maxRow + 1)):
        if ws.cell(i, j).value != None:
            for row in ws.iter_cols(min_row=1, max_row=i, min_col=2, max_col=j):
                values = []
                for clm in row:
                    values.append(clm.value)
            print(str(j) + '列の最終行 : ' + str(i))
            print(values)
            break

