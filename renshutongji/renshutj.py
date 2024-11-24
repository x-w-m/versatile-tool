import pandas as pd
from pyecharts.charts import Bar
from pyecharts import options as opts
import os

# 考试文件和对应的工作表信息
exam_files = [
    {'file': '考试成绩/2323级分科分班成绩.xlsx', 'sheets': ['物理', '历史'], 'exam_name': '分班'},
    {'file': '考试成绩/2024年10月高二月考班级赋分成绩.xlsx', 'sheets': ['物理', '历史'], 'exam_name': '10月月考'},
    {'file': '考试成绩/2024年11月高二期中考试班级赋分成绩.xlsx', 'sheets': ['物理', '历史'], 'exam_name': '期中考试'},
    # 添加更多考试信息 {'file': 'exam3.xlsx', 'sheets': ['物理', '历史'], 'exam_name': '考试3'},
]

# 定义列名
headers = ['班级', '姓名', '考号', '选科', '原始总分', '赋分总分', '班排名', '校排名', '语文', '数学', '外语',
           '选科', '化学原始分', '化学', '生物原始分', '生物', '政治原始分', '政治', '地理原始分', '地理']


# 统计函数
def count_students_in_rank(data, rank_thresholds, rank_col='校排名'):
    results = {}
    for rank in rank_thresholds:
        filtered_data = data[data[rank_col] <= rank]
        group_counts = filtered_data.groupby('班级')[rank_col].count()
        results[f'前{rank}名'] = group_counts
    return pd.DataFrame(results)


# 动态柱状图生成函数
def create_bar_chart(data, titile):
    """
    创建动态柱状图，支持切换显示不同的项目。
    :param data: DataFrame，索引为班级，列为项目，值为人数。
    """
    # 创建图表对象
    bar = Bar(init_opts=opts.InitOpts(width="1200px", height="500px"))
    bar.add_xaxis(data.index.tolist())  # X 轴为班级

    # 添加所有项目
    for project in data.columns:
        bar.add_yaxis(
            series_name=project,
            y_axis=data[project].tolist()  # 默认只显示第一个项目
        )

    # 设置全局选项
    bar.set_global_opts(
        title_opts=opts.TitleOpts(title=titile),
        legend_opts=opts.LegendOpts(
            is_show=True,
            pos_top="5%",
            type_="scroll"  # 图例可滚动，适合项目多的情况
        ),
        toolbox_opts=opts.ToolboxOpts(),  # 工具栏选项
    )
    return bar


# 创建一个空的字典，用于存储所有统计结果
all_statistics = {}

# 遍历考试文件和工作表
for exam in exam_files:
    file = exam['file']
    sheets = exam['sheets']
    exam_name = exam['exam_name']

    for sheet_name in sheets:
        # 读取表格数据
        data = pd.read_excel(file, sheet_name=sheet_name, header=None, skiprows=3)
        data.columns = headers

        # 根据科目设置排名分段
        if sheet_name == '物理':
            rank_thresholds = [50, 100, 200, 500, 800]
        elif sheet_name == '历史':
            rank_thresholds = [30, 60, 120, 300, 480]
        else:
            continue

        # 统计结果
        result = count_students_in_rank(data, rank_thresholds)
        bar = create_bar_chart(result, f'{exam_name}_{sheet_name}')
        bar.render(f'charts/{exam_name}_{sheet_name}.html')
        # 将结果存入字典，以方便统一保存
        all_statistics[f'{exam_name}_{sheet_name}'] = result

# 保存所有统计结果到一个 Excel 文件
# output_file = '人数统计.xlsx'
# with pd.ExcelWriter(output_file) as writer:
#     for sheet_name, result in all_statistics.items():
#         result.to_excel(writer, sheet_name=sheet_name)
#
# print(f"统计结果已保存到 {output_file}")

import os

# 获取所有生成的 HTML 文件
output_directory = "charts/"
html_files = [f for f in os.listdir(output_directory) if f.endswith(".html")]



for html_file in html_files:
    file_path = os.path.join(output_directory, html_file)

    with open(file_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    # 添加 CSS 样式以使图表居中
    centered_html_content = html_content.replace(
        '<body >',
        '<body style="display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0;">'
    )

    with open(file_path, "w", encoding="utf-8") as f:
        f.write(centered_html_content)

print("所有图表已设置为居中显示")

# 创建索引页面
index_path = os.path.join(output_directory, "index.html")
with open(index_path, "w", encoding="utf-8") as index_file:
    index_file.write("<html>\n<head><title>图表索引</title></head>\n<body>\n")
    index_file.write("<h1>图表索引</h1>\n<ul>\n")

    # 添加每个 HTML 文件的链接
    for html_file in html_files:
        link = f'<li><a href="{html_file}" target="_blank">{html_file}</a></li>\n'
        index_file.write(link)

    index_file.write("</ul>\n</body>\n</html>")

print(f"索引页面已生成，路径为：{index_path}")