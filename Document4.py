# -*- coding: utf-8 -*-
#import openpyxl, pprint
import openpyxl as px
from datetime import datetime

# �t�@�C�����̎w��Ȃ� --- (*1)
file_master = "file_master.xlsx" # �}�X�^�[�f�[�^
touhoku = "touhoku.xlsx" # ���k�̃f�[�^
kita = "kita.xlsx"       # �k�֓��̃f�[�^
minami = "minami.xlsx"   # ��֓��̃f�[�^
shizuoka  = "shizuoka.xlsx" # �É��̃f�[�^
hiroshima = "hiroshima.xlsx" # �L���̃f�[�^
fukuoka ="fukuoka.xlsx"      # �����̃f�[�^
tokyo ="tokyo.xlsx"      # �����̃f�[�^
nagoya ="nagoya.xlsx"      # ���É��̃f�[�^
osaka ="osaka.xlsx"      # ���̃f�[�^


file_master2 = "file_master2.xlsx"
file_master3 = "file_master3.xlsx"
file_master4 = "file_master4.xlsx"
file_master5 = "file_master5.xlsx"
file_master6 = "file_master6.xlsx"
file_master7 = "file_master7.xlsx"
file_master8 = "file_master8.xlsx"
file_master9 = "file_master9.xlsx"
file_master10 = "file_master10.xlsx"
file_master11 = "file_master11.xlsx"
file_master12 = "file_master12.xlsx"
file_master13 = "file_master13.xlsx"
file_master14 = "file_master14.xlsx"
file_master15 = "file_master15.xlsx"
file_master16 = "file_master16.xlsx"
file_master17 = "file_master17.xlsx"
file_master18 = "file_master18.xlsx"
file_master19 = "file_master19.xlsx"
file_master20 = "file_master20.xlsx"
file_master21 = "file_master21.xlsx"
file_master22 = "file_master22.xlsx"
file_master23 = "file_master23.xlsx"
file_master24 = "file_master24.xlsx"
file_master25 = "file_master25.xlsx"
file_master26 = "file_master26.xlsx"
file_master27 = "file_master27.xlsx"
file_master28 = "file_master28.xlsx"
file_master29 = "file_master29.xlsx"
file_master30 = "file_master30.xlsx"
file_master31 = "file_master31.xlsx"
file_master32 = "file_master32.xlsx"
file_master33 = "file_master33.xlsx"
file_master34 = "file_master34.xlsx"
file_master35 = "file_master35.xlsx"
file_master36 = "file_master36.xlsx"
file_master37 = "file_master37.xlsx"
file_master38 = "file_master38.xlsx"
file_master39 = "file_master39.xlsx"
file_master40 = "file_master40.xlsx"
file_master41 = "file_master41.xlsx"
file_master42 = "file_master42.xlsx"
file_master43 = "file_master43.xlsx"
file_master44 = "file_master44.xlsx"
file_master45 = "file_master45.xlsx"
file_master46 = "file_master46.xlsx"
file_master47 = "file_master47.xlsx"
file_master48 = "file_master48.xlsx"
file_master49 = "file_master49.xlsx"
file_master50 = "file_master50.xlsx"
file_master51 = "file_master51.xlsx"


# ���k�f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(touhoku, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

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

# ���k�f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(touhoku, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Dec"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master2.xlsx')
ws_iv = wb_iv["Sheet10"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master3)
print("ok")

# �k�֓��f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(kita, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master3.xlsx')
ws_iv = wb_iv["Sheet9"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=201+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master4)
print("ok")

# �k�֓��f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(kita, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Dec"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master4.xlsx')
ws_iv = wb_iv["Sheet10"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=201+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master5)
print("ok")

# ��֓��f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(minami, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master5.xlsx')
ws_iv = wb_iv["Sheet9"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=401+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master6)
print("ok")

# ��֓��f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(minami, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Dec"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master6.xlsx')
ws_iv = wb_iv["Sheet10"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=401+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master7)
print("ok")

# �É��f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(shizuoka, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master7.xlsx')
ws_iv = wb_iv["Sheet9"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=601+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master8)
print("ok")

# �É��f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(shizuoka, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Dec"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master8.xlsx')
ws_iv = wb_iv["Sheet10"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=601+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master9)
print("ok")

# �L���f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(hiroshima, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master9.xlsx')
ws_iv = wb_iv["Sheet9"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=801+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master10)
print("ok")

# �L���f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(hiroshima, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Dec"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master10.xlsx')
ws_iv = wb_iv["Sheet10"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=801+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master11)
print("ok")

# �����f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(fukuoka, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master11.xlsx')
ws_iv = wb_iv["Sheet9"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1001+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master12)
print("ok")

# �����f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(fukuoka, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Dec"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master12.xlsx')
ws_iv = wb_iv["Sheet10"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1001+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master13)
print("ok")

# �����f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(tokyo, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master13.xlsx')
ws_iv = wb_iv["Sheet9"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1201+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master14)
print("ok")

# �����f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(tokyo, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Dec"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master14.xlsx')
ws_iv = wb_iv["Sheet10"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1201+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master15)
print("ok")

# ���É��f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(nagoya, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master15.xlsx')
ws_iv = wb_iv["Sheet9"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1401+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master16)
print("ok")

# ���É��f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(nagoya, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Dec"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master16.xlsx')
ws_iv = wb_iv["Sheet10"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1401+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master17)
print("ok")

# ���f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(osaka, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Nov"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master17.xlsx')
ws_iv = wb_iv["Sheet9"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1601+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master18)
print("ok")

# ���f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(osaka, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["Dec"] # �V�[�g����I��
list_data = ws["A14:AX200"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master18.xlsx')
ws_iv = wb_iv["Sheet10"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1601+y+1, column=0+x+1, value=v)
   
# �V�����ۑ����� --- (*6)
wb_iv.save(file_master19)
print("ok")

#�����󒍁A���������͏I���I�I�I�I�I�I�I�I�I�I�I
#��������́A����j���[�X
#���k

#�������

# ���k�f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(touhoku, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["News"] # �V�[�g����I��
list_data = ws["A3:M11"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master19.xlsx')
ws_iv = wb_iv["shi"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1+y+1, column=0+x+1, value=v)

# �V�����ۑ����� --- (*6)
wb_iv.save(file_master20)
print("ok")

#�������

# ���k�f�[�^��ǂݍ��� --- (*2)
wb = px.load_workbook(touhoku, data_only=True) # �����łȂ��l�����o���ꍇ
ws = wb["News"] # �V�[�g����I��
list_data = ws["A15:M24"] # �C�ӂ͈̔͂��擾

# �}�X�^�f�[�^��ǂ� --- (*3)
wb_iv = px.load_workbook('file_master20.xlsx')
ws_iv = wb_iv["kyo"]

# �[�i������������ --- (*5)
for y, row in enumerate(list_data):
  for x, cell in enumerate(row):
    if (cell is None) or (cell.value is None): continue
    v = cell.value
    ws_iv.cell(row=1+y+1, column=0+x+1, value=v)

# �V�����ۑ����� --- (*6)
wb_iv.save(file_master21)
print("ok")















