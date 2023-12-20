from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter


def excel_format():
    global cell, cell
    workbook = load_workbook("考生去向表.xlsx")
    # 设置边框样式
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    # 针对每个工作表进行格式调整
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]

        col_width_ratios = [4.0, 6, 5, 4.0, 7, 5, 5, 6, 6]
        first_col_width = 8  # 第一列宽度为7

        # 设置列宽
        for idx, ratio in enumerate(col_width_ratios, start=1):
            column_width = first_col_width * (ratio / col_width_ratios[0])
            worksheet.column_dimensions[get_column_letter(idx)].width = column_width

        # 设置行高
        for row in worksheet.iter_rows():
            worksheet.row_dimensions[row[0].row].height = 22

            # 设置单元格的字体、边框和对齐方式
            for cell in row:
                # 设置字体
                if cell.row == 1:  # 标题行
                    cell.font = Font(name='微软雅黑', size=12)
                else:  # 数据行
                    cell.font = Font(name='微软雅黑', size=11)

                # 设置边框
                cell.border = thin_border

                # 设置对齐
                cell.alignment = Alignment(horizontal='center', vertical='center')
    # 保存更改
    workbook.save("考生去向表_格式调整.xlsx")
    print("调整后的表格已保存到“考生去向表_格式调整.xlsx”文件中。")


