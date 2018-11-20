# -*- coding: utf-8 -*-
#import openpyxl, pprint
import openpyxl as px
from datetime import datetime

# ファイル名の指定など --- (*1)
file_master = "file_master.xlsx" # マスターデータ
touhoku= "touhoku.xlsx" # 東北のデータ
file_master2 = "file_master2.xlsx"

# 東北データを読み込む --- (*2)
wb = px.load_workbook(touhoku, data_only=True) # 数式でなく値を取り出す場合
ws = wb["Nov"] # シート名を選ぶ
list_data = ws["A14:AX324"] # 任意の範囲を取得

# マスタデータを読む --- (*3)
wb_iv = px.load_workbook('file_master.xlsx')
ws_iv = wb_iv["Sheet9"]

# 納品物を書き込む --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1+y+1, column=0+x+1, value=v)
   
# 新しく保存する --- (*6)
wb_iv.save(file_master2)
print("ok")