import pandas as pd


def calculate_ranks(file_path, sheet_name, output_file):
    # 定义列名
    columns_scores = ['班级', '姓名', '考号', '选科', '原始总分', '赋分总分', '班排名', '校排名', '语文', '数学',
                      '外语', sheet_name, '化学原始分', '化学', '生物原始分', '生物', '政治原始分', '政治',
                      '地理原始分', '地理']

    # 读取数据
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, skiprows=3, names=columns_scores)

    # 处理赋分科目班排名和校排名
    for subject in ['语文', '数学', '外语', sheet_name, '化学', '生物', '政治', '地理']:
        if subject not in df.columns or df[subject].isnull().all():
            continue

        # 班排名：按班级分组计算排名
        df[f'{subject}_班排名'] = df.groupby('班级')[subject].rank(method='min', ascending=False)

        # 校排名：直接按全校范围排名
        df[f'{subject}_校排名'] = df[subject].rank(method='min', ascending=False)

    # 保存结果到新文件
    df.to_excel(output_file, index=False)
    print(f"排名计算完成，结果已保存到 {output_file}")


if __name__ == '__main__':
    # 本次考试成绩文件路径
    file2 = './班级成绩/2024年11月高二期中班级赋分成绩.xlsx'
    # 指定工作表名称
    sheet_names = ['物理', '历史']

    # 为每个工作表添加排名
    for sheet_name in sheet_names:
        output_file = f'./{sheet_name}_排名.xlsx'
        calculate_ranks(file2, sheet_name, output_file)
