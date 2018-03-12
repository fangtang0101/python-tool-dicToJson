#-*-coding:utf-8 -*-
import xlrd
import sys
import chardet
from xlutils.copy import copy

import os  
import json
import xlwt

import codecs 

import json
import time



val_all = {"optionsValue1":[],"options1":[]}
target_path_model = "model.xls"
key_sheet_name    = "Sheet1"
col_target = 1
book = xlrd.open_workbook(target_path_model,formatting_info=True)
sheets=book.sheets()
sheet_A37 = book.sheet_by_name(key_sheet_name)
rows = sheet_A37.nrows
cols = sheet_A37.ncols
list_cell = []

for row in range(rows):
	cell = sheet_A37.cell_value(row,col_target) 
	cell2 = sheet_A37.cell_value(row,col_target+1) 
	if cell.strip() != '':
		val_all["optionsValue1"].append(cell)
		val_all["options1"].append(cell2)




# print(val_all)

json_str = json.dumps(val_all,ensure_ascii=False)
print(type(json_str))
print(json_str)


with open('test.json', 'w') as json_file:
	json_file.write(json.dumps(val_all,ensure_ascii=False))





