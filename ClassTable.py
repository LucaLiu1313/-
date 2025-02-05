import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os

# 语文战字典存储时间段与班级
语文战 = {}
时间段 = []
班级 = []

# 读取 Excel 数据
df = pd.read_excel(r"大课表的地址")

# 遍历 DataFrame 查找 "语文战"
for row_index, row in df.iloc[4:].iterrows():
    for col_index, col in enumerate(df.columns[1:], start=1):
        value = row[col]
        first_col_value = df.iloc[row_index, 0]  # 该行的第一列（时间段）
        first_row_value = df.iloc[0, col_index]  # 该列的第一行（班级）

        if value == '语文战':
            时间段.append(first_col_value)
            班级.append(first_row_value)
            语文战[first_col_value] = first_row_value  # 确保字典有数据

# 构建 DataFrame
data = {"时间段": 时间段, "班级": 班级}
df1 = pd.DataFrame(data)

# Excel 文件路径
excel_path = r"小课表excel的地址"

# 确保 Excel 文件存在
if not os.path.exists(excel_path):
    df1.to_excel(excel_path, index=False, engine="openpyxl")  # 先创建文件
print(df1)
# 载入 Excel
wb = load_workbook(excel_path, read_only=False, keep_vba=True)
ws = wb.active

# 遍历 "语文战" 数据并写入 Excel
for key, value in 语文战.items():
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            if str(cell.value).strip() == key:  # 确保 key 匹配
                row_index = cell.row  # 获取匹配的行索引
                ws.cell(row=row_index, column=8, value=value)  # 在第8列写入
                print(f"成功写入: {key} -> 行 {row_index}, 班级: {value}")


# 保存 Excel
wb.save(excel_path)
wb.close()
