import pyautogui as pg
import time
import openpyxl
import datetime
# 対象パス C:\Users\竜馬\OneDrive\share\集計転記\syuukeiin.xlsx

dt_now = datetime.datetime.now()

print(dt_now) # 2022-06-09 01:03:52.865150(出力時)

print(dt_now.strftime('%Y%m%d')) # 20220609
