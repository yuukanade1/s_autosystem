import pyautogui as pg
import time
import openpyxl
import datetime
import subprocess
import win32gui
from tkinter import messagebox

subprocess.Popen(r'C:\Windows\notepad.exe')
memoapp = win32gui.FindWindow(None, '無題 - メモ帳')
time.sleep(1)
win32gui.SetForegroundWindow(memoapp) # エラーあり

dt_now = datetime.datetime.now()
print(dt_now.strftime('%Y%m%d'))

# excelファイルよりデータ取得
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
# 基本情報の入力処理
            time.sleep(1)
            pg.write(str(values[0]))
            pg.press('enter')
            pg.press('0')
            pg.press('enter')
            pg.write(dt_now.strftime('%Y%m%d'))
            pg.press('enter')
# オーダー数の処理
            for k in range(1, i):
                pg.write(str(values[k]))
                pg.press('enter')
            pg.write('F7')
            pg.press('enter')
            time.sleep(2)
            print(str(j) + '列の最終行 : ' + str(i))
            print(values)

            break
wb.close()
messagebox.showinfo('自動入力', '処理が完了しました。')