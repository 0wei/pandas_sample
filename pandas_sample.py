
import pandas as pd
import string
from openpyxl import load_workbook

def column_to_index(column):
    if len(column) == 1:
        return string.ascii_uppercase.index(column)
    else:
        return (string.ascii_uppercase.index(column[0]) + 1) * 26 + string.ascii_uppercase.index(column[1])

# 示例用法
# column_index = column_to_index('AQ')
# print(column_index)


# 读取 Excel 文件，跳过表头行
file_path = "1.xlsx"
table = "Sheet1"
df = pd.read_excel(file_path,sheet_name=table,header=None)
# 更新单元格数据
df.at[10,0] = "aaabb"
# 将列 SN 数据写入到 Username 中
# df['Username'] = df['SN'].apply(lambda x: f"NN-{x}")

# 写入
def save():
    # book = load_workbook(file_path)
    writer = pd.ExcelWriter(file_path,mode='a',if_sheet_exists='replace')
    # writer.book = book
    df.to_excel(writer, sheet_name=table,header=None)
    # writer.save()
    writer.close()


def test():

    shape = df.shape
    # 输出 DataFrame 的行数和列数
    print("行数:", shape[0])
    print("列数:", shape[1])

    print(df)

    # 行打印
    for c in df.loc[2]:
        print(f"row:{c}")

    # 列打印
    for c in df[0]:
        print(f"colum:{c}")


    # 这里的 42 表示 AQ 列的索引，索引从 0 开始
    print(df.iloc[1, column_to_index('D')] )


    # 遍历表格中的每一行
    for index, row in df.iterrows():
        # 遍历行中的每个单元格
        for column_name, cell_value in row.items():
            # 在这里处理每个单元格的数据
            # 一行的数据
            print(f"行名: {index}, 列名: {column_name}, 值: {cell_value}")

    print(df[2][0])
