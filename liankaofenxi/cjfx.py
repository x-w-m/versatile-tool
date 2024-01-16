import pandas as pd

# 读取Excel文件，不处理标题行
df = pd.read_excel("处理后的成绩表.xlsx")
# 分组按学校计算各排名范围内的学生人数
rank_ranges = [100, 150, 200, 250, 500, 1000]
schools = df['学校'].unique()
summary = pd.DataFrame(index=schools)

for rank in rank_ranges:
    column_name = f'前{rank}名'
    summary[column_name] = df[df['总分联名'] <= rank].groupby('学校').size()

# 删除在所有排名范围内人数都为空的学校
summary = summary.dropna(how='all')
# 根据前1000名学生人数降序排序
summary = summary.sort_values(by='前1000名', ascending=False)
# 保存结果到Excel文件
# summary.to_excel('前1000统计.xlsx')

# 显示部分结果以确认
print(summary.head())

# 定义分数段
score_ranges = [650, 620, 600, 570, 550, 520, 500, 480]

# 创建一个新的DataFrame来保存结果
school_stats = pd.DataFrame(index=schools)

# 计算每个学校的总人数
total_students_per_school = df.groupby('学校').size()

# 对每个分数段计算学生人数和百分比
for score in score_ranges:
    # 计算分数段内的学生人数
    num_students_column = f'{score}及以上人数'
    school_stats[num_students_column] = df[df['赋分总分'] >= score].groupby('学校').size()

    # 计算分数段内的学生人数占总人数的百分比
    percent_column = f'{score}及以上百分比'
    school_stats[percent_column] = (school_stats[num_students_column] / total_students_per_school) * 100

# 用0填充NaN值
school_stats.fillna(0, inplace=True)

# 将百分比列的格式设置为百分比形式
for score in score_ranges:
    percent_column = f'{score}及以上百分比'
    school_stats[percent_column] = school_stats[percent_column].apply(lambda x: f'{x:.2f}%')

# 显示部分结果以确认
print(school_stats.head())

# 保存结果到Excel文件
school_stats.to_excel('各分数段学生人数及百分比统计.xlsx')
