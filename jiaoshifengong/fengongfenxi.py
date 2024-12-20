# 读取分工表，生成教师任课数据
# 分工表数据格式要求：大标题一行，科目标题一行；表名：高一、高二、高三；同年级无同名教师
# 科目名称要求：['语文', '数学', '外语', '物理', '化学', '生物', '政治', '历史', '地理', '体育', '声乐', '舞蹈', '美术','信息', '通用', '研学', '心理']
from collections import Counter

import numpy as np
import pandas as pd


# 读取表格，表名必须是年级
def read_excel_file(file_name, sheet_index):
    # 读取Excel文件
    df = pd.read_excel(file_name, sheet_name=sheet_index, header=1)
    # 打印数据
    print(df)
    return df


# 提取教师和科目
def create_teachers_df(df):
    # 创建一个新的DataFrame来存储老师的信息
    teachers_df = pd.DataFrame(columns=['教师', '科目', '班主任'])
    # 需要统计的科目列
    subjects = ['语文', '数学', '外语', '物理', '化学', '生物', '政治', '历史', '地理', '体育', '声乐', '舞蹈', '美术',
                '信息', '通用', '研学', '心理']
    # 遍历原始DataFrame的每一行和每一列
    for index, row in df.iterrows():
        for subject in subjects:
            # 如果这一行列表示的单元格的值不是NaN，就把它添加到结果集新的行中
            if subject in row and pd.notna(row[subject]):
                new_row = {'教师': row[subject], '科目': subject, '班主任': None}
                teachers_df.loc[len(teachers_df)] = new_row
    # 删除重复行,并重置索引
    teachers_df = teachers_df.drop_duplicates(ignore_index=True)
    print(teachers_df)
    return teachers_df


# 标记班主任，并按科目顺序排序
def update_class_teacher_info(df, teachers_df):
    # 遍历原始DataFrame的'班主任'列
    for index, row in df.iterrows():
        class_teacher = row['班主任']  # 班主任姓名
        class_name = row['班级']  # 对应的班级

        # 在teachers_df中找到班主任，并更新他们的'Class'列
        teachers_df.loc[(teachers_df['教师'] == class_teacher), '班主任'] = class_name
    # 定义科目的排序顺序
    subject_order = {'语文': 1, '数学': 2, '外语': 3, '物理': 4, '化学': 5, '生物': 6, '政治': 7, '历史': 8, '地理': 9,
                     '体育': 10, '美术': 11, '声乐': 12, '舞蹈': 13, '信息': 14, '通用': 15, '研学': 16, '心理': 17}
    # 使用map函数将科目列转换为排序顺序
    teachers_df['SubjectOrder'] = teachers_df['科目'].map(subject_order)
    # 根据科目的排序顺序对DataFrame进行排序
    teachers_df.sort_values(by=['SubjectOrder'], ignore_index=True, inplace=True)
    # 删除排序顺序列
    teachers_df.drop('SubjectOrder', axis=1, inplace=True)
    return teachers_df


# 提取教师所有任教班级，并计算数量
def update_class_info(df, teachers_df):
    # 创建一个新的列'班级'，并初始化为None
    teachers_df['班级'] = None
    # 需要统计的科目列
    subjects = ['语文', '数学', '外语', '物理', '化学', '生物', '政治', '历史', '地理', '体育', '声乐', '舞蹈', '美术',
                '信息', '通用', '研学', '心理']
    # 遍历原始DataFrame的每一行
    for index, row in df.iterrows():
        class_name = row['班级']  # 获取'班级'列的值
        for subject in subjects:
            # 如果这一列的值不是NaN，就把班级添加到对应老师的'班级'列中
            if subject in row and pd.notna(row[subject]):
                teacher = row[subject]
                # 如果一个老师已经有了班级，我们需要在现有的班级后面添加","，然后再添加新的班级
                if pd.notna(teachers_df.loc[(teachers_df['教师'] == teacher), '班级'].iloc[0]):
                    teachers_df.loc[(teachers_df['教师'] == teacher), '班级'] += ',' + str(class_name)
                else:
                    teachers_df.loc[(teachers_df['教师'] == teacher), '班级'] = str(class_name)
    # 添加班级数列，记录教师任教班级的数量
    teachers_df["班级数"] = teachers_df['班级'].str.count(',') + 1
    return teachers_df


# 分别提取高考班和学考班，并计算数量
def update_exam_class_info(df, teachers_df):
    # 判断分工表df中是否有“科目”这一列，如果没有说明是高一，直接退出函数
    if '科目' not in df.columns:
        return teachers_df

    # 创建两个新的列'高考班'和'学考班'，并初始化为None
    teachers_df['高考班'] = None
    teachers_df['学考班'] = None

    # 将语文、数学、外语三科的班级列的值复制到它们的高考班列中,其他科目根据
    for subject in ['语文', '数学', '外语']:
        teachers_df['高考班'] = np.where(
            teachers_df['科目'] == subject,
            teachers_df['班级'],
            teachers_df['高考班']
        )
    # 高考科目的简写
    gaokao_subjects = {'物': '物理', '化': '化学', '生': '生物', '政': '政治', '历': '历史', '地': '地理'}
    # 遍历原始DataFrame的每一行
    for index, row in df.iterrows():
        class_name = row['班级']  # 获取'班级'列的值
        gaokao_subjects_str = row['科目']  # 获取'科目'列的值，即科目组合
        for subject_abbr, subject in gaokao_subjects.items():
            # 如果'科目'列的值包含某个科目的简写，就把班级添加到对应老师的'高考班'列中
            if subject_abbr in gaokao_subjects_str:
                teacher = row[subject]  # 高考科目一定有安排教师，不需要额外判断。
                if pd.notna(teachers_df.loc[(teachers_df['教师'] == teacher), '高考班'].iloc[0]):
                    teachers_df.loc[(teachers_df['教师'] == teacher), '高考班'] += ',' + str(class_name)
                else:
                    teachers_df.loc[(teachers_df['教师'] == teacher), '高考班'] = str(class_name)
            # 否则，就把班级添加到对应老师的'学考班'列中
            else:
                teacher = row[subject]  # 高二年级学考科目存在无教师情况，需额外进行nan值判断。
                if pd.notna(teacher) and pd.notna(teachers_df.loc[(teachers_df['教师'] == teacher), '学考班'].iloc[0]):
                    teachers_df.loc[(teachers_df['教师'] == teacher), '学考班'] += ',' + str(class_name)
                else:
                    teachers_df.loc[(teachers_df['教师'] == teacher), '学考班'] = str(class_name)

    # 添加两个新的列来记录高考班和学考班的个数
    teachers_df['高考班数'] = teachers_df['高考班'].str.count(',') + 1
    teachers_df['学考班数'] = teachers_df['学考班'].str.count(',') + 1

    return teachers_df


# 跨楼分析，计算的是所有班级跨楼情况，如果要只计算高考班，需要删除分工表中的学考班级任课。
def cross_floor_analysis(df, teachers_df):
    # 判断分工表df中是否有“教学楼”这一列，如果没有则不分析跨楼
    if '教学楼' not in df.columns:
        return teachers_df
    # 创建班级-楼栋映射
    class_building_map = dict(zip(df['班级'], df['教学楼']))
    print(class_building_map)
    # 分析教师跨楼情况
    results = []

    for _, row in teachers_df.iterrows():
        teacher = row['教师']
        # 将班级字符串转换为整数列表，匹配字典键的数据类型
        classes = [int(cls.strip()) for cls in row['班级'].split(',')]
        print(classes)

        if len(classes) < 3:
            continue

        buildings = [class_building_map.get(cls, 'Unknown') for cls in classes]
        building_counts = Counter(buildings)

        cross_building = len(building_counts) > 1

        if cross_building:
            main_building = max(building_counts, key=building_counts.get)
            cross_building_classes = [cls for cls, building in zip(classes, buildings) if building != main_building]
        else:
            cross_building_classes = []

        results.append({
            '教师': teacher,
            '是否跨楼': '是' if cross_building else '否',
            '跨楼班级': ','.join(str(cls) for cls in cross_building_classes) if cross_building else ''
        })

    result_df = pd.DataFrame(results)
    # 将跨楼结果合并到教师信息表中
    teachers_df = pd.merge(teachers_df, result_df, how="left", on="教师")
    return teachers_df


# 表名必须是年级
def main(file_dir, file_name, sheet_name):
    df = read_excel_file(file_dir + file_name, sheet_name)
    teachers_df = create_teachers_df(df)
    teachers_df = update_class_teacher_info(df, teachers_df)
    teachers_df = update_class_info(df, teachers_df)
    teachers_df = update_exam_class_info(df, teachers_df)
    teachers_df = cross_floor_analysis(df, teachers_df)
    # 在第一列插入年级
    teachers_df.insert(0, "年级", sheet_name)
    teachers_df.to_excel(f"{file_dir}教师信息{sheet_name}.xlsx", index=False)


# 将三个年级的信息合并到一个文件中
def merge(file_dir, file_list):
    df_list = []
    for file in file_list:
        df = pd.read_excel(file_dir + file)
        df_list.append(df)
    merge_df = pd.concat(df_list)
    merge_df.to_excel(f"{file_dir}教师任课信息.xlsx", index=False)


if __name__ == '__main__':
    # 目录
    file_dir = "2024年下/1014/"
    file_name = "2024年下学期教学分工总表 - 10.14.xlsx"
    # 年级列表
    grade_list = ["高一", "高二", "高三"]
    for grade in grade_list:
        main(file_dir, file_name, grade)

    # 需要合并的文件列表
    file_list = ["教师信息高一.xlsx", "教师信息高二.xlsx", "教师信息高三.xlsx"]
    merge(file_dir, file_list)
