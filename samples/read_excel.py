import os
import xlrd3
excel_path = os.path.join(os.path.dirname(__file__), "data/test_data.xlsx")
wb = xlrd3.open_workbook(excel_path)  # 创建工作簿对象
sheet = wb.sheet_by_name("Sheet1")   # 创建表格对象
merged = sheet.merged_cells   # 返回一个列表，起始行、结束行、起始列、结束列

def get_merged_cell_value(row_index, col_index):
    cell_value = None
    for (rlow, rhigh, clow, chigh) in merged:
        if (row_index >= rlow and row_index < rhigh):
            if (col_index >= clow and col_index < chigh):
                cell_value = sheet.cell_value(rlow, clow)
                break; # 防止循环去进行判断出现值覆盖的情况
            else:
                cell_value = sheet.cell_value(row_index, col_index)
        else:
            cell_value = sheet.cell_value(row_index, col_index)
    return cell_value
for i in range(1, 9):
    for j in range(0,4):
        print(get_merged_cell_value(i, j))