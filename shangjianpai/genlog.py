import pandas as pd

# 读取两个Excel文件的数据
df_workdays = pd.read_excel("工作日上.xlsx")
df_schedule = pd.read_excel("上机安排（22年上）.xlsx")

# 准备用来存储机房使用记录的列表
records = []

# 对于上机安排中的每一行（每一条上机记录），生成对应的机房使用记录
for idx, row in df_schedule.iterrows():
    teacher_name = row['教师姓名']
    day_of_week = row['星期']
    period = row['节次']
    classroom = row['机房']
    week_type = row['单双周']

    # 在“工作日.xlsx”中筛选对应星期和单双周的日期
    selected_dates = df_workdays[(df_workdays['星期'] == day_of_week) & (df_workdays['单双周'] == week_type)]

    for _, date_row in selected_dates.iterrows():
        record = {
            "日期": date_row["日期"].date(),
            "星期": date_row["星期"],
            "周次": date_row["周次"],
            "单双周": date_row["单双周"],
            "教师姓名": teacher_name,
            "节次": period,
            "机房": classroom
        }
        records.append(record)

# 将机房使用记录转换为DataFrame并保存到Excel文件
df_records = pd.DataFrame(records)
df_records.to_excel("机房使用记录（22年上）.xlsx", index=False)
