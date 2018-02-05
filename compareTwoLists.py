#!/usr/bin/env python
# encoding: utf-8
from openpyxl import load_workbook
from openpyxl.styles import Font, Color,colors

# 找第一个非空单元格
def find_row_count_at_column(ws, row, column):
    count = 0
    for i in range(row, 3000):
        if ws.cell(row=i, column=column).value is None:
            return count
        else:
            count += 1

# 比较文件
FILE_NAME = 'cpTest.xlsx'

# 源表单名称
SRC_SHEET_NAME = 'Sheet1'

# 比较结果表单
COMPARE_RESULT = 'result'

# 核对结果表单位置
RESULT_SHEET_NUMBER = 1

# 首行号
FIRST_LINE = 2

# 源列号
COLUMN_SRC = 1

# 目标列号
COLUMN_DES = 4

# 加载文件
wb = load_workbook(filename = FILE_NAME)

# 若比较结果表单则先删除
try:
    ws = wb[COMPARE_RESULT]
    wb.remove(ws)
except KeyError as e:
    print ('No sheet test:' + e.__str__())


wss = wb[SRC_SHEET_NAME]

# 创建比较结果表单
wb.create_sheet(COMPARE_RESULT)
wsr = wb[COMPARE_RESULT]

# 写入标题行
# 序号	部门	姓名	账号	账号类型	对账时间
wsr["A1"] = "账号"
wsr["B1"] = "姓名"

wsr["D1"] = "账号"
wsr["E1"] = "姓名"


# 标题行黑体
wsr["A1"].font = Font(bold=True)
wsr["B1"].font = Font(bold=True)

wsr["D1"].font = Font(bold=True)
wsr["E1"].font = Font(bold=True)


# 当前行号
currentRow = 2

srcRowCount = find_row_count_at_column(wss, FIRST_LINE, COLUMN_SRC)
desRowCount = find_row_count_at_column(wss, FIRST_LINE, COLUMN_DES)

# 读取源数据进数据字典

srcDict = {}
desDict = {}

for i in range(FIRST_LINE, FIRST_LINE + srcRowCount):

    srcDict[wss.cell(row=i, column=COLUMN_SRC).value] = wss.cell(row=i, column=COLUMN_SRC + 1).value

# print (srcDict)

# 读取对比数据进数据字典

for i in range(FIRST_LINE, FIRST_LINE + desRowCount):

    desDict[wss.cell(row=i, column=COLUMN_DES).value] = wss.cell(row=i, column=COLUMN_DES + 1).value

# print (desDict)


# 循环源数据字典，在对比数据字典中找到相同的，则拷贝到新的结果页，同时在源数据字典和目标数据字典中都删除


for key in srcDict.copy():
    # print (key, 'corresponds to', srcDict[key])

    if key in desDict.keys():
        # print (key, 'corresponds to', desDict[key])

        wsr.cell(row=currentRow, column=COLUMN_SRC).value = key
        wsr.cell(row=currentRow, column=COLUMN_SRC + 1).value = srcDict[key]

        wsr.cell(row=currentRow, column=COLUMN_DES).value = key
        wsr.cell(row=currentRow, column=COLUMN_DES + 1).value = desDict[key]

        del srcDict[key]
        del desDict[key]
        currentRow += 1

wsr.cell(row=currentRow, column=COLUMN_SRC).value = "差异"
wsr.cell(row=currentRow, column=COLUMN_SRC).font = Font(bold=True, color=colors.RED)
currentRow += 1

for key in srcDict.keys():
    wsr.cell(row=currentRow, column=COLUMN_SRC).value = key
    wsr.cell(row=currentRow, column=COLUMN_SRC + 1).value = srcDict[key]
    currentRow += 1

for key in desDict.keys():
    wsr.cell(row=currentRow, column=COLUMN_DES).value = key
    wsr.cell(row=currentRow, column=COLUMN_DES + 1).value = desDict[key]
    currentRow += 1

# print (srcDict)
# print (desDict)

# print (str(srcRowCount))
# print (str(desRowCount))


# print(wss.cell(row=1, column=1).value)

wb.active = RESULT_SHEET_NUMBER
wb.save(FILE_NAME)

