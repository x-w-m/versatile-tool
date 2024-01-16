import pandas as pd

# 读取Excel文件，不处理标题行
df = pd.read_excel('高一赋分_六科报表.xlsx', header=None, skiprows=2)

# 手动设置列标题
column_titles = ["学校", "年级", "班级", "姓名", "学号", "考号", "原始总分", "赋分总分", "总分班名", "总分校名",
                 "总分联名",
                 "语文原始分数", "语文班名", "语文校名", "语文联名", "数学原始分数", "数学班名", "数学校名", "数学联名",
                 "英语原始分数", "英语班名", "英语校名", "英语联名", "物理原始分数", "物理班名", "物理校名", "物理联名",
                 "化学原始分数", "化学赋分分数", "化学班名", "化学校名", "化学联名", "生物原始分数", "生物赋分分数",
                 "生物班名", "生物校名", "生物联名"]
df.columns = column_titles
print("计算总排名")
# 计算总分联名
df['总分联名'] = df['赋分总分'].rank(method='min', ascending=False)
print("计算校排名")
# 计算总分校名
df['总分校名'] = df.groupby('学校')['赋分总分'].rank(method='min', ascending=False)
print("计算班排名")
# 计算总分班名
df['总分班名'] = df.groupby(['学校', '班级'])['赋分总分'].rank(method='min', ascending=False)

# 显示前几行数据以确认排名是否计算正确
print(df.head())
print("保存文件")
# 将处理后的DataFrame写入新的Excel文件
df.to_excel('处理后的成绩表.xlsx', index=False)
