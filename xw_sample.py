import xlwings as xw
from xlwings import Range

from xlwings_wrap import join_to_address, sheet_filter_ranges, sheet_filter_row, sheet_filter_column


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
    column = range_find_value(sheet.range(
        "A1").expand('right'), subjects_name).column
    row = range_find_value(sheet.range("A1").expand('down'), name).row
    value = sheet.range(row, column).value
    print(f"find {name} {subjects_name} {row},{column}, {value}")
    return value


def sheet2_math_score():
    sheet = wb.sheets['Sheet2']
    for row in sheet.range('A2').expand('down'):
        for subjects_name in sheet.range(f"C1").expand("right"):
            sheet.range(row.row, subjects_name.column).value = sheet1_match_name_score(
                row.value, subjects_name.value)


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
def sum_all_subjects():
    sheet_1 = wb.sheets['Sheet1']
    sheet_2 = wb.sheets['Sheet2']
    for row_name in sheet_2.range('A2').expand('down'):
        location = sheet_filter_row(sheet_1.range("A1").expand("down"), [row_name.value])
        location = sheet_filter_column(location, ["B", "C", "D"])
        # list_range = range_to_list_range(sheet_1.range("A2:D2").expand('down'))
        # list_range = filter_row(list_range, "A", row_name.value)
        # list_range = pick_columns(list_range, "B", "C", "D")
        # s = sum_list_range(list_range)
        # sheet_2.range(f"B{row_name.row}").value = f"{s}"
        # address = join_to_address(list_range)
        address = join_to_address(location)
        sheet_2.range(f"B{row_name.row}").formula = f"=SUM({address})"


# def match_cell(rows, columns, source_rows, match_rows, sour_columns, match_columns):
#     for r_cell in rows.rows:
#         location_row = r_cell.row
#         for c_cell in columns:
#             location_column = c_cell.column
#             location = sheet_location(source_rows.sheet,
#                                       source_rows, [match_rows],
#                                       sour_columns, [match_columns])
#             if location is None:
#                 continue
#             rows.sheet.range(location_row, location_column) \
#                 .formula = f"={location[0].get_address(include_sheetname=True)}"


def fill_subjects():
    sheet_1 = wb.sheets['Sheet1']
    sheet_2 = wb.sheets['Sheet2']

    for cell_title in sheet_2.range('C1').expand('right'):
        for cell_name in sheet_2.range("A2").expand("down"):
            name = cell_name.value
            subject_name = cell_title.value
            location = sheet_filter_ranges(sheet_1,
                                           sheet_1.range("A1").expand("down"), [name],
                                           sheet_1.range("A1").expand("right"), [subject_name])
            if location is None:
                continue
            sheet_2.range(cell_name.row, cell_title.column) \
                .formula = f"={location[0].get_address(include_sheetname=True)}"


if __name__ == '__main__':
    # 打开原始 Excel 文件
    app = xw.App(visible=True, add_book=False)
    wb = app.books.open("1.xlsx")
    # fill_subjects()
    sum_all_subjects()
    # sheet2_math_score()
    wb.save()  # 保存文件
    wb.close()  # 关闭文件
    app.quit()  # 关闭程序
    print("完成!")
