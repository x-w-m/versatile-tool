import pandas as pd
from reportlab.lib import pagesizes
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm

# 注册微软雅黑字体
pdfmetrics.registerFont(TTFont('微软雅黑', '字体文件/msyh.ttc'))
pdfmetrics.registerFont(TTFont('微软雅黑粗体', '字体文件/msyhbd.ttc'))

# 读取Excel文件
df = pd.read_excel('考生去向表.xlsx', sheet_name='考生去向表')

# 根据班级进行分组
grouped = df.groupby('班级')

# 页面大小和边距
page_width, page_height = A4
left_margin = right_margin = 2 * cm
top_margin = bottom_margin = 1 * cm

# 创建PDF文档
pdf = SimpleDocTemplate('考生去向表.pdf', pagesize=A4, rightMargin=right_margin, leftMargin=left_margin,
                        topMargin=top_margin, bottomMargin=bottom_margin)

# 创建一个空列表来存放流式内容
elements = []

# 列宽比例
col_width_ratios = [4.0, 6, 5, 4.0, 7, 5, 5, 6, 6]
total_ratio = sum(col_width_ratios)

# 计算可用宽度
available_width = page_width - (left_margin + right_margin)
scaled_col_widths = [ratio / total_ratio * available_width for ratio in col_width_ratios]

# 遍历每个班级
for class_name, group in grouped:
    num_rows = len(group) + 1  # 数据行数加上标题行

    num_rows = len(group) + 1  # 数据行数加上标题行

    # 计算可用高度
    available_height = page_height - (top_margin + bottom_margin)
    # 计算最大行高
    max_row_height = available_height / num_rows
    # 反向计算最大字体大小（假设行高是字体大小的0.8倍）
    max_font_size = max_row_height - 1
    print(max_font_size)
    # 限定字体大小的下限为6
    font_size = max(6, min(7, max_font_size))
    # 计算实际行高
    row_height = font_size * 0.8

    # 设置表格样式
    table_style = TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), '微软雅黑粗体'),  # 标题行使用微软雅黑粗体
        ('FONTNAME', (0, 1), (-1, -1), '微软雅黑'),  # 数据行使用微软雅黑
        ('FONTSIZE', (0, 0), (-1, -1), font_size),  # 字体大小
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # 水平居中
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # 垂直居中
        ('INNERGRID', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('BOX', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
    ])

    # 创建表格
    data = [group.columns.to_list()] + group.values.tolist()
    t = Table(data, colWidths=scaled_col_widths)
    t.setStyle(table_style)
    elements.append(t)
    elements.append(PageBreak())  # 每个班级新起一页

# 构建PDF
pdf.build(elements)
