import pandas as pd

# 2023年体测成绩表中无家庭住址选项
cj_column_names = ['年级编号', '班级编号', '班级名称', '学籍号', '民族代码', '姓名', '性别', '出生日期',
                   '身高', '体重', '体重评分', '体重等级', '肺活量',
                   '肺活量评分', '肺活量等级', '五十米跑', '五十米跑评分', '五十米跑等级', '立定跳远', '立定跳远评分',
                   '立定跳远等级', '坐位体前屈', '坐位体前屈评分', '坐位体前屈等级', '800米跑', '800米跑评分',
                   '800米跑等级', '800米跑附加分', '1000米跑', '1000米跑评分', '1000米跑等级', '1000米跑附加分',
                   '仰卧起坐', '仰卧起坐评分', '仰卧起坐等级', '仰卧起坐附加分', '引体向上', '引体向上评分',
                   '引体向上等级', '引体向上附加分', '标准分', '附加分', '总分', '总分等级']
mb_column_names = ['序号', '识别码', '学籍号', '年级', '班级', '姓名', '性别', '体重', '体重评分', '体重等级', '肺活量',
                   '肺活量评分', '肺活量等级', '五十米跑', '五十米跑评分', '五十米跑等级', '立定跳远', '立定跳远评分',
                   '立定跳远等级', '坐位体前屈', '坐位体前屈评分', '坐位体前屈等级', '800米跑', '800米跑评分',
                   '800米跑等级', '800米跑附加分', '1000米跑', '1000米跑评分', '1000米跑等级', '1000米跑附加分',
                   '仰卧起坐', '仰卧起坐评分', '仰卧起坐等级', '仰卧起坐附加分', '引体向上', '引体向上评分',
                   '引体向上等级', '引体向上附加分']
dtype_dict = {'序号': str, '识别码': str, '学籍号': str, '年级': str, '班级': str, '姓名': str, '性别': str}
print(len(cj_column_names), len(mb_column_names))
# “体测成绩/2021.xls”文件中的数据
df_cj = pd.read_excel('体测成绩/2023.xls', names=cj_column_names)
df_cj['性别'].replace({1: '男', 2: '女'}, inplace=True)
# 读取“2021级体测模板.xlsx”文件中的数据
df_mb = pd.read_excel('2021级体测模板.xlsx', names=mb_column_names, dtype=dtype_dict)

# 2023年为9:40列,2022年前为10:41列
cols_to_keep = ['学籍号', '姓名', '性别', '总分等级'] + list(df_cj.columns[9:40])
df_cj_selected = df_cj[cols_to_keep].copy()
df_mb = df_mb.merge(df_cj_selected, on='学籍号', how='left', suffixes=('', '_cj'))

for col in df_mb.columns:
    if col.endswith('_cj'):
        original_col = col[:-3]  # 去掉 '_cj' 后缀
        df_mb[original_col] = df_mb[original_col].combine_first(df_mb[col])

print(df_mb.columns)
df_mb = df_mb[mb_column_names]

# 将df_cj_selected中能在df_mb中找到学籍号的数据删除
df_cj_selected = df_cj_selected[~df_cj_selected['学籍号'].isin(df_mb['学籍号'])]

# 处理无法通过学籍号找到匹配数据的行
# 删选出df_mb中体重值为空的行
df_mb_selected = df_mb[df_mb['体重'].isnull()]

import random

# 定义需要复制的列
cols_to_copy = ['体重', '体重评分', '体重等级', '肺活量', '肺活量评分', '肺活量等级', '五十米跑', '五十米跑评分',
                '五十米跑等级', '立定跳远', '立定跳远评分',
                '立定跳远等级', '坐位体前屈', '坐位体前屈评分', '坐位体前屈等级', '800米跑', '800米跑评分',
                '800米跑等级', '800米跑附加分', '1000米跑', '1000米跑评分', '1000米跑等级', '1000米跑附加分',
                '仰卧起坐', '仰卧起坐评分', '仰卧起坐等级', '仰卧起坐附加分', '引体向上', '引体向上评分',
                '引体向上等级', '引体向上附加分']

# 遍历 df_mb_selected 中的每一行
for i, row in df_mb_selected.iterrows():
    # 在 df_cj_selected 中查找有相同姓名和性别的行
    matches = df_cj_selected[(df_cj_selected['姓名'] == row['姓名']) & (df_cj_selected['性别'] == row['性别'])]

    # 如果找到了匹配的行
    if not matches.empty:
        # 如果只有一个匹配的行，或者所有匹配的行中只有一个 '总分等级' 不是 '不及格'
        if matches.shape[0] == 1 or matches[matches['总分等级'] != '不及格'].shape[0] == 1:
            match = matches.iloc[0]
        else:
            # 如果有多个匹配的行，选择 '总分等级' 不是 '不及格' 的那个
            match = matches[matches['总分等级'] != '不及格'].iloc[0]
    else:
        # 如果没有找到匹配的行，从 df_cj_selected 中随机选择一条 '总分等级' 为 '良好' 且性别相同的数据
        match = \
            df_cj_selected[(df_cj_selected['总分等级'] == '良好') & (df_cj_selected['性别'] == row['性别'])].sample(
                1).iloc[
                0]

    # 将匹配的行中的数据复制到 df_mb_selected 中
    df_mb_selected.loc[i, cols_to_copy] = match[cols_to_copy]

    # 从 df_cj_selected 中删除被复制过的数据
    df_cj_selected = df_cj_selected.drop(match.name)

# 将二次匹配后的数据合并到 df_mb 中
df_mb.loc[df_mb['体重'].isnull()] = df_mb_selected

df_mb.to_excel('2021级2023年体测数据.xlsx', sheet_name='合并数据', index=False)
