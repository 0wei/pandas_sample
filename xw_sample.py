import xlwings as xw
from xlwings import Range


# [超全整理｜Python 操作 Excel 库 xlwings 常用操作详解！ - 知乎](https://zhuanlan.zhihu.com/p/346813124)

def range_find_value(rang: Range, value: str):
    for r in rang:
        if r.value == value:
            return r
    return None


# 读取原始数据
def sheet1_sum(name: str):
    sheet = wb.sheets['Sheet1']
    find_row = None
    for row in sheet.range("A2").expand('down'):
        if row.value == name:
            find_row = row
            break
    numbers = find_row.offset(0, 1).expand("right")[:3].value
    sum_result = sum(map(lambda x: x, numbers))
    return sum_result


def sum_all():
    sheet = wb.sheets['Sheet2']
    for row in sheet.range('A2').expand('down'):
        sheet.range(f"B{row.row}").value = sheet1_sum(row.value)


def sheet1_match_name_score(name: str, subjects_name: str):
    sheet = wb.sheets['Sheet1']
    column = range_find_value(sheet.range("A1").expand('right'), subjects_name).column
    row = range_find_value(sheet.range("A1").expand('down'), name).row
    value = sheet.range(row, column).value
    print(f"find {name} {subjects_name} {row},{column}, {value}")
    return value


def sheet2_math_score():
    sheet = wb.sheets['Sheet2']
    for row in sheet.range('A2').expand('down'):
        for subjects_name in sheet.range(f"C1").expand("right"):
            sheet.range(row.row, subjects_name.column).value = sheet1_match_name_score(row.value, subjects_name.value)


# sheet = wb.sheets['Sheet1']
# data_range = sheet['B2:D10']
# for row in data_range.rows:
#     # 遍历每一列
#     sum = 0
#     for cell in row:
#         # 打印单元格的值
#         sum += float(cell.value)
#     print(sum)    
#     sheet.range(f"F{row.row}").value = sum

# row_count = sheet.api.UsedRange.Rows.Count
# column_count = sheet.api.UsedRange.Columns.Count
# print("行数：", row_count)
# print("列数：", column_count)

# # 获取数据范围
# data_range = sheet.used_range

# # 遍历每一行
# for row in data_range.rows:
#     # 遍历每一列
#     for cell in row:
#         # 打印单元格的值
#         print(cell.value)


# def unitinfo():
#     cell = sheet.range("A2")
#     print(cell)

# 连接到新的 Excel 文件
# new_wb = xw.Book()
# new_sheet = new_wb.sheets['Sheet1']

# 将数据写入新的 Excel 文件，并保留样式
# new_sheet.range('A1').value = df

# 保存新的 Excel 文件
# new_wb.save('output.xlsx',)
# new_wb.close()


if __name__ == '__main__':
    # 打开原始 Excel 文件
    app = xw.App(visible=True)
    wb = app.books.open(f"1.xlsx")
    sheet2_math_score()
    # 关闭原始 Excel 文件
    wb.close()
    app.quit()
