# -*- coding: utf-8 -*-
#import openpyxl, pprint
import openpyxl as px
from datetime import datetime

# �t�@�C�����̎w��Ȃ� --- (*1)
file_master = "file_master.xlsx" # �}�X�^�[�f�[�^
touhoku= "touhoku.xlsx" # ���k�̃f�[�^
file_master2 = "file_master2.xlsx"

# ���k�f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(touhoku, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX324"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master.xlsx')
ws_iv = wb_iv["Sheet9"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master2)
print("ok")