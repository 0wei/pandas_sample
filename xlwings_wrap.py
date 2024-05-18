import string

import xlwings as xw
from xlwings import Range, Sheet


def print_list_range(list_range: list[list]):
    print('----------')
    for row in list_range:
        for cell in row:
            print(cell.value, end=" ")
        print('')
    print('----------')


def _column_letter_to_index(column_letter: str):
    column = 0
    for letter in column_letter.upper():
        column *= 26
        column += string.ascii_uppercase.index(letter) + 1
    return column


# def find_row(list_range: list[list], column_letter: str, value):
#     rows = filter_column(list_range, column_letter, value)
#     if len(rows)>0:
#         return rows[0]
#     return None
# """
# 查找指定行
# :param list_range: 目标列表
# :param column_letter:  查找的列名称, 如 A
# :param value: 查找列的值
# :return:
# """
# column = column_letter_to_index(column_letter)
# for row in list_range:
#     for cell in row:
#         if cell.column == column and (value is None or cell.value == value):
#             return row
# return None


def list_range_filter_row(list_range: list[list], column_letter: str, *values):
    """
    过滤符合的行
    :param list_range: 目标列表
    :param column_letter:  查找的列名称, 如 A
    :param value: 查找列的值
    :return:
    """
    f = []
    column = _column_letter_to_index(column_letter)
    for row in list_range:
        for cell in row:
            if cell.column == column and (cell.value in values):
                f.append(row)
                break
    return f


def range_filter_row(row_list: Range, row_filter_values: list):
    rang_rows = find_range_by_value(row_list, *row_filter_values)
    return rang_rows


def find_range_by_value(source_range: Range, *value):
    f = []
    for cell in source_range:
        if cell.value in value:
            f.append(cell)
    if len(f) > 0:
        print(f"Find {source_range.address} {value}")
        return f
    print(f"No found {source_range.address} {value}")
    return None


def sheet_filter_ranges_row_column(sheet: Sheet, row_list: Range, row_filter_values: list, column_list: Range,
                                   column_filter_values: list):
    """
    筛选符合行列的单元格
    :param sheet:
    :param row_list: 过滤的行
    :param row_filter_values: 匹配的行名称
    :param column_list: 过滤的列
    :param column_filter_values: 匹配的列名称
    :return:
    """
    # sheet = row_list.sheet
    rang_rows = find_range_by_value(row_list, *row_filter_values)
    if rang_rows is None:
        return None
    rang_columns = find_range_by_value(column_list, *column_filter_values)
    if rang_columns is None:
        return None
    f = []
    for row in rang_rows:
        for column in rang_columns:
            f.append(sheet.range(row.row, column.column))
    if len(f) > 0:
        print(f"Find location {"".join([x.address for x in f])} row:{row_filter_values} column:{column_filter_values}")
        return f
    print(f"Not Find location row:{row_filter_values} column:{column_filter_values}")

    return None


def list_range_filter_column(row_list: list[list], columns: list):
    """
    筛选需要的列
    :param row_list:
    :param columns:
    :return:
    """
    f = []
    for row in row_list:
        r = []
        row_0 = row[0]
        sheet = row_0.sheet
        for c in columns:
            r.append(sheet.range(f"{c}{row_0.row}"))
        f.append(r)
    if len(f) > 0:
        # print(f"Find location {"".join([x for x in f])}")
        return f
    print(f"Not Find location column")
    return None


# def list_range_pick_columns(list_range: list[list], *column_letter):
#     """
#     筛选需要的列
#     :param list_range:  目标列表
#     :param column_letter: 列名称
#     :return:
#     """
#     columns = list(map(lambda letter: column_letter_to_index(letter), column_letter))
#     # column = column_letter_to_index(column_letter)
#     # return map(lambda row: filter(lambda cell: cell.column in columns, row), list_range)
#     f = []
#     # column = column_letter_to_index(column_letter)
#     for row in list_range:
#         filter_row = []
#         for cell in row:
#             if cell.column in columns:
#                 filter_row.append(cell)
#         if len(filter_row) > 0:
#             f.append(filter_row)
#     return f


def unfold_list_range(list_range: list[list]):
    """
    展开二维数组
    :param list_range:
    :return:
    """
    return sum(list_range, [])
    # f = []
    # for row in list_range:
    #     for cell in row:
    #         f.append(cell)
    # return f


def range_to_list_range(rang: Range):
    """
    将 Range 对象的行转为二维列表
    :param rang: 目标 range
    :return:
    """
    return [row for row in rang.rows]


def list_range_sum_list_range(list_range: list[list]):
    """
    求目标列表的总和
    :param list_range:
    :return:
    """

    def cell_value_to_float(cell):
        try:
            return float(cell.value)
        except Exception as e:
            print(f"cell_value_to_float error: {e}")
            return 0.0

    return sum(map(cell_value_to_float, unfold_list_range(list_range)))


def list_range_join_to_address(list_range: list[list]):
    address = [cell.get_address(include_sheetname=True) for cell in unfold_list_range(list_range)]
    return ", ".join(address)


if __name__ == '__main__':
    # 打开原始 Excel 文件
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(f"1.xlsx")
    sheet_1 = wb.sheets['Sheet1']
    used_range = [row for row in sheet_1.used_range.rows]
    # print_list_range(used_range)
    # # used_range = filter_column(used_range, 'B', 86)
    # print_list_range(used_range)
    # used_range = filter_row(used_range, 'A', '李四')
    # print_list_range(used_range)
    # used_range = pick_columns(used_range, 'B', 'C', 'D')
    # print(f"李四: sum:{sum_list_range(used_range)}")
    # 关闭原始 Excel 文件
    wb.close()
    app.quit()
