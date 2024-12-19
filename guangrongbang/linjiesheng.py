"""
各班各科临界生，成绩表表名必须为“物理”，“历史”。
"""
import pandas as pd


def analyze_critical_students(grades_file, thresholds_file):
    # 设置 pandas 显示选项
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    # 读取成绩表和临界线分数表
    grades_df = pd.ExcelFile(grades_file)
    thresholds_df = pd.read_excel(thresholds_file)

    # 初始化结果字典
    results = {}

    # 初始化 ExcelWriter
    output_file = thresholds_file.replace("临界分数线", "临界学生名单")

    with (pd.ExcelWriter(output_file, engine='openpyxl') as writer):
        for sheet_name in grades_df.sheet_names:
            # 读取每个方向的成绩表
            df = pd.read_excel(grades_df, sheet_name=sheet_name, header=None, skiprows=3)

            # 定义列名
            headers = ['班级', '姓名', '考号', '选科', '原始总分', '赋分总分', '班排名', '校排名', '语文', '数学',
                       '外语',
                       sheet_name, '化学原始分', '化学', '生物原始分', '生物', '政治原始分', '政治', '地理原始分',
                       '地理']
            df.columns = headers

            # 根据方向筛选对应的临界分数
            thresholds = thresholds_df[thresholds_df['方向'] == sheet_name]

            # 初始化结果 DataFrame
            critical_students = pd.DataFrame()

            # 初始化统计结果 DataFrame
            critical_summary = pd.DataFrame()

            # 确保科目顺序为总分、语文、数学、外语、物理、历史、化学、生物、政治、地理
            subject_order = ['赋分总分', '语文', '数学', '外语', sheet_name, '化学', '生物', '政治', '地理']

            # 遍历每个科目，按照指定顺序统计
            for subject in subject_order:
                column_name = subject

                # 特殊处理“物理/历史”列名
                if subject in ['物理', '历史']:
                    threshold_column = '物理/历史'
                else:
                    threshold_column = subject

                if column_name in df.columns:
                    # 获取科目临界线
                    threshold = thresholds[threshold_column].values[
                        0] if threshold_column in thresholds.columns else None

                    if threshold is not None:
                        # 筛选临界学生
                        mask = df[column_name].between(threshold - 15, threshold + 5)
                        subject_critical = df.loc[mask, ['班级', '姓名', '考号', column_name]]

                        # 添加科目列标识
                        subject_critical['项目'] = subject

                        # 调整列顺序，将“项目”列放在第一列
                        subject_critical = subject_critical[['项目', '班级', '姓名', '考号', column_name]]

                        # 合并结果
                        critical_students = pd.concat([critical_students, subject_critical], ignore_index=True)

            # 按项目排序时保持指定顺序，同时对班级进行排序
            if not critical_students.empty:
                # 获取实际类别
                actual_categories = critical_students['项目'].unique()
                # 合并实际类别与预定义顺序
                all_categories = [cat for cat in subject_order if cat in actual_categories] + [cat for cat in actual_categories if cat not in subject_order]
                # 确保“项目”列按指定顺序排序
                critical_students['项目'] = pd.Categorical(critical_students['项目'], categories=all_categories,
                                                           ordered=True)
                # 按“项目”顺序和“班级”升序排序
                critical_students = critical_students.sort_values(by=['项目', '班级'], key=lambda
                    col: col if col.name == '班级' else col.cat.codes).reset_index(drop=True)

            # 按班级和项目统计临界生人数并调整表格形式
            if not critical_students.empty:
                critical_summary = critical_students.groupby(['班级', '项目'], observed=True).size().unstack(
                    fill_value=0)
                # 按指定科目顺序重新排列列
                critical_summary = critical_summary.reindex(columns=subject_order, fill_value=0)

            # 按班级分组
            grouped = critical_students.groupby('班级')

            # 保存到结果字典
            results[sheet_name] = grouped

            # 将数据写入 Excel 文件
            critical_students.to_excel(writer, sheet_name=sheet_name, index=False)

            # 将统计结果写入 Excel 文件
            critical_summary.to_excel(writer, sheet_name=f"{sheet_name}_人数统计", index=True)


if __name__ == "__main__":
    # 成绩表
    grades_file = './班级成绩/2024年12月高二月考班级赋分成绩.xlsx'

    # 临界线分数表
    thresholds_file = './临界生/临界分数线/2024年12月高二月考临界分数线.xlsx'

    # 执行分析
    analyze_critical_students(grades_file, thresholds_file)
