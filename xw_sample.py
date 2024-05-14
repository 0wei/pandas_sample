import pandas as pd
import xlwings as xw

# [超全整理｜Python 操作 Excel 库 xlwings 常用操作详解！ - 知乎](https://zhuanlan.zhihu.com/p/346813124)



# 打开原始 Excel 文件
app = xw.App(visible=True)
wb = app.books.open(f"E:\\ps\\pandas_sample\\1.xlsx")

# 读取原始数据
sheet = wb.sheets['Sheet1']
def sheet1Sum(name):
    sheet = wb.sheets['Sheet1']
    findRow = None
    for row in sheet.range("A2").expand('down'):
        if row.value == name:
            findRow = row
            break
    numbers = findRow.offset(0,1).expand("right")[:3].value
    sumResult = sum(map(lambda x: x, numbers))
    return sumResult
   
def sumAll():
    sheet = wb.sheets['Sheet2']
    for row in sheet.range('A2').expand('down'):
        sheet.range(f"B{row.row}").value = sheet1Sum(row.value)

def shee1MatchNameScore(name,score):
    sheet = wb.sheets['Sheet1']
    column = None
    row = None
    for c in sheet.range("A1").expand('down'):
        if c.value == name:
            row = c.row
            break   
    for c in sheet.range("A1").expand('right'):
        if c.value == score:
            column = c.column
            break
    print(f"find {name} {score} {row},{column}")
    return sheet[row,column].value

def shee2MathScore():
    sheet = wb.sheets['Sheet2']
    for row in sheet.range('A2').expand('down'):
        sheet.range(f"C{row.row}").value = shee1MatchNameScore(row.value,sheet[f"C1"].value)
        sheet.range(f"D{row.row}").value = shee1MatchNameScore(row.value,sheet[f"D1"].value)

data_range = sheet['B2:D10']
for row in data_range.rows:
    # 遍历每一列
    sum = 0
    for cell in row:
        # 打印单元格的值
        sum += float(cell.value)
    print(sum)    
    sheet.range(f"F{row.row}").value = sum

row_count = sheet.api.UsedRange.Rows.Count
column_count = sheet.api.UsedRange.Columns.Count
print("行数：", row_count)
print("列数：", column_count)

# 获取数据范围
data_range = sheet.used_range

# 遍历每一行
for row in data_range.rows:
    # 遍历每一列
    for cell in row:
        # 打印单元格的值
        print(cell.value)



def unitinfo():
    cell = sheet.range("A2")
    print(cell)

# 连接到新的 Excel 文件
# new_wb = xw.Book()
# new_sheet = new_wb.sheets['Sheet1']

# 将数据写入新的 Excel 文件，并保留样式
# new_sheet.range('A1').value = df

# 保存新的 Excel 文件
# new_wb.save('output.xlsx',)
# new_wb.close()

# 关闭原始 Excel 文件
wb.close()
app.quit()