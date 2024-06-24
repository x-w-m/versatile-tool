from decimal import Decimal

import numpy as np
# 高考奖金计算
import pandas as pd


# file1：原始分班成绩
# file2：高考成绩
# file3：高考分数线
def analysis_of_college_entrance_examination_awards(file1, file2, file3):
    # 设置控制台的宽度和每列数据的最大宽度，None表示不限制
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', None)

    # 打开Excel文件
    gkcjb = pd.ExcelFile(file2)

    # 打开分数线表格,标题有两行
    fsx = pd.read_excel(file3, header=[0, 1])

    # 遍历所有工作表的名称(物理、历史)
    for sheet_name in gkcjb.sheet_names:
        # 读取每个工作表的内容到一个DataFrame
        df_gkcj = pd.read_excel(gkcjb, sheet_name=sheet_name, header=None, skiprows=3)

        # 定义数据列名，四选二赋分省略“赋分”直接使用科目名，便于后续标记科目
        headers = ['班级', '姓名', '考号', '选科', '原始总分', '赋分总分', '班排名', '校排名', '语文', '数学', '外语',
                   sheet_name, '化学原始分', '化学', '生物原始分', '生物', '政治原始分', '政治', '地理原始分', '地理']

        # 重命名列名
        df_gkcj.columns = headers

        # 获取分数线中对应科目方向的本科线分数
        # 本科线分数
        bkx = fsx.loc[0, (sheet_name, '本科线')]

        # 划分分数区间
        score_ranges = np.arange(bkx, 751, 20)
        # 为分数区间定义加权系数，第一个为1，第二个为1.1，以此类推
        k = [Decimal('1') + Decimal('0.1') * i for i in range(len(score_ranges))]
        # 创建区间
        # 创建第一个区间（左闭右闭）
        first_interval = pd.Interval(left=bkx, right=bkx + 20, closed='both')
        # 创建后续区间（左开右闭）
        subsequent_intervals = pd.interval_range(start=bkx + 20, end=751, freq=20, closed='right')
        # 合并区间
        intervals = pd.Index([first_interval]).union(subsequent_intervals)
        # 直接创建的区间第一个区间无法做到左闭
        # intervals = pd.interval_range(start=bkx, end=751, freq=20)

        # 将区间转换为字符串格式的列表。intervals本身也是各序列，不转换也可以传递给cut()方法的labels参数。
        # str_intervals = [str(interval) for interval in intervals]
        # 创建一个字典，其中键是分数范围（区间），值是对应的系数
        k_dict = dict(zip(intervals, k))
        # 使用pandas的cut函数将“赋分总分”列划分为分数段,inclusive= True表示包含最小值,即第一个区间左侧闭合
        df_gkcj['分数区间'] = pd.cut(df_gkcj['赋分总分'], labels=intervals, bins=score_ranges, include_lowest=True)
        # 计算各分数段各班级的人数，结果为一个Series
        gkpjrs = df_gkcj.groupby(['分数区间', '班级'], observed=True).size()

        # 重置索引，将Series转换为DataFrame，索引转换为普通列，size()结果保存为“人数”列
        df_gkdf = gkpjrs.reset_index(name="人数")
        # 为每个分数段的人数添加对应的系数,系数需要转换为浮点数，否则会出现TypeError错误，无法进行乘法运算。
        df_gkdf["系数"] = df_gkdf["分数区间"].map(k_dict).astype(float)
        # 计算每个班每分数段的高考评价分
        df_gkdf["高考评价分"] = df_gkdf["人数"] * df_gkdf["系数"]
        # 将评价分按班级求和，结果为一个Series
        gkpjf_b = df_gkdf.groupby('班级')["高考评价分"].sum()

        # 生成一个新的排名列，其中每个人的排名都是唯一的,降序排序
        df_gkcj['校排名_无重复'] = df_gkcj['赋分总分'].rank(method='first', ascending=False).astype(int)
        # 使用groupby函数和agg函数来获取每个分数段的最小和最大校排名
        df_rank = df_gkcj.groupby('分数区间', observed=True)['校排名_无重复'].agg(['min', 'max'])
        df_rank["系数"] = df_rank.index.map(k_dict).astype(float)
        # 创建一个新的区间索引，其中每个区间的左右边界分别是'min'和'max'列的值
        df_rank['排名区间'] = pd.IntervalIndex.from_arrays(df_rank['min'], df_rank['max'], closed='both')

        # 读取原始分班成绩
        df_ysfbcj = pd.read_excel(file1, sheet_name=sheet_name, header=None, skiprows=3)
        # 重命名列名
        df_ysfbcj.columns = headers

        # 使用遍历的方式来判断排名所处区间,可以有效避免排名区间不连续，cut方法无法正确处理的问题。
        # results是一个元组列表，每个元组包含两个元素，第一个是区间，第二个是系数
        results = df_ysfbcj['校排名'].apply(lambda x: find_interval(df_rank['排名区间'], df_rank['系数'], x))
        # 对results进行解压，将区间和系数分别赋值给两个新列
        # TODO:由于results元组中含有None值，pandas会将区间边界类型转换为float64。如何让结果保持原始类型？
        df_ysfbcj['排名区间'], df_ysfbcj['系数'] = zip(*results)
        # 统计每个区间每个班级的人数'排名区间'与'系数'等价。加上“系数”用于保留系数值。
        yspjrs = df_ysfbcj.groupby(['排名区间', '系数', '班级']).size()
        # 重置索引，将Series转换为DataFrame，索引转换为普通列，size()结果保存为“人数”列
        df_ysdf = yspjrs.reset_index(name="人数")

        # 根据各区间班级人数与对应区间系数的乘积来计算原始评价分
        df_ysdf["原始评价分"] = df_ysdf["人数"] * df_ysdf["系数"]
        # 将评价分按班级求和
        yspjf_b = df_ysdf.groupby('班级')['原始评价分'].sum()

        # 计算评价分差值
        diff = gkpjf_b - yspjf_b
        # 保留两位小数
        diff = diff.round(2)
        # 将“班级、原始评价分、高考评价分评价分差值”四项数据保存到新的表格中
        # 班级是三个Series索引，数据会根据索引对齐，同时index默认True，保存时会显性保存为一列。
        df_df = pd.DataFrame(
            {'原始评价分': yspjf_b, '高考评价分': gkpjf_b, '评价分差值': diff})
        # 创建一个ExcelWriter对象
        with pd.ExcelWriter(f'{sheet_name}评价分.xlsx') as writer:
            # 将df_df写入一个名为'评价分'的工作表
            df_df.to_excel(writer, sheet_name=f'{sheet_name}组评价分')
            # 将df_gkdf（即高考得分表）写入一个名为'高考得分表'的工作表
            df_gkdf.to_excel(writer, sheet_name=f'{sheet_name}组高考得分表', index=False)
            # 将df_ysdf（即原始得分表）写入一个名为'原始得分表'的工作表
            df_ysdf.to_excel(writer, sheet_name=f'{sheet_name}组原始得分表', index=False)
            # 将df_rank（即排名区间表）写入一个名为'分数排名对照表'的工作表
            df_rank.to_excel(writer, sheet_name=f'{sheet_name}组分数排名对照表')
            # 选择df_gkcj的特定列并将其保存到一个名为'高考区间表'的工作表
            columns = ['班级', '姓名', '选科', '赋分总分', '校排名', '分数区间']
            df_gkcj[columns].to_excel(writer, sheet_name=f'{sheet_name}组高考区间表', index=False)
            # 选择df_ysfbcj的特定列并将其保存到一个名为'原始区间表'的工作表
            columns = ['班级', '姓名', '选科', '赋分总分', '校排名', '排名区间']
            df_ysfbcj[columns].to_excel(writer, sheet_name=f'{sheet_name}组原始区间表', index=False)

        print(f'{sheet_name}评价分.xlsx已生成')


def find_interval(rank_ranges, coefficients, value):
    for interval, coefficient in zip(rank_ranges, coefficients):
        # 迭代时判断是否是区间，避免迭代到NaN值。
        if isinstance(interval, pd.Interval) and value in interval:
            return interval, coefficient
        # 如果没有找到匹配的区间，可以选择设置为None或其他默认值
    return None, None


if __name__ == '__main__':
    analysis_of_college_entrance_examination_awards('23级分班成绩.xlsx', '23级高考成绩.xlsx', '本科分数线.xlsx')
