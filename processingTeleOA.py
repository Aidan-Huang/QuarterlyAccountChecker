#!/usr/bin/env python
# encoding: utf-8
from openpyxl import load_workbook
from openpyxl.styles import Font
import re

re1 = r'\/.+?(?=[\s;]|$)'

def find_the_first_not_none_value_upward(ws, row, column):
	for i in range(row - 1, 1, -1):
		if ws[column + str(i)].value is not None:
			return ws[column + str(i)]

filename = '17Q1.xlsx'

wb = load_workbook(filename = filename)

try:
    ws = wb['test']
    wb.remove(ws)
except KeyError as e:
    print ('no sheet test:' + e.__str__())

# 创建临时表单

wb.create_sheet('test')
wst = wb['test']

wst["A1"] = "部门"
wst["A1"].font = Font(bold=True)
wst["B1"] = "成员"
wst["B1"].font = Font(bold=True)
trow = 2


ws = wb['部门信息']

# 在B列找“成员”内容的单元格
# 在A列找相应的部门
# 在C列具体成员删除冗余字符串
# 拷贝 “部门”，“成员”到临时表单
for row in ws.rows:
	for cell in row:
		if cell.value == '成员':
			if cell.column == 'B':
				# print(ws['A' + str(cell.row - 1)].value)
				departName = find_the_first_not_none_value_upward(ws, cell.row, 'A').value
				# print (departName)
				usersStr = ws['C' + str(cell.row )].value
				usersStr = re.sub(re1, '', str(usersStr))
				usersStr= re.sub('admin_lxgs;', '', usersStr)
				usersStr= re.sub(r'\d', '', usersStr)
				usersStr= re.sub('_', '', usersStr)
				if usersStr != "None" and departName not in ["工会", "团委", "财务处", "党委", "公司领导"]:
					users = usersStr.split(';')
					for user in users:
						wst["A" + str(trow)] = departName
						wst["B" + str(trow)] = user
						trow += 1
					# print (replace)
wb.active = 2
wb.save(filename)

print ("Total " + str(trow - 2) + " employees.")