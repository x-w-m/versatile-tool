import pandas as pd


# 定义一个函数来生成日期范围内的工作日信息
def generate_workdays(start_date, end_date):
    # 生成日期范围
    date_range = pd.date_range(start_date, end_date)

    # 筛选周一到周五的日期
    workdays = date_range[date_range.weekday < 5]

    # 创建一个DataFrame来存储日期、星期、周次、单双周信息
    df = pd.DataFrame({
        "日期": workdays.date,  # 转为日期格式
        "星期": workdays.strftime("%A"),
        "周次": (workdays.isocalendar().week - workdays[0].isocalendar().week + 1).astype(int),
        "单双周": (((workdays.isocalendar().week - workdays[0].isocalendar().week) % 2) + 1).astype(int)
    })

    # 替换英文的星期为中文
    df["星期"] = df["星期"].replace({
        "Monday": "周一",
        "Tuesday": "周二",
        "Wednesday": "周三",
        "Thursday": "周四",
        "Friday": "周五"
    })

    return df


# 调用函数并保存结果
df_workdays = generate_workdays("2022-02-15", "2022-06-03")
df_workdays.to_excel("工作日.xlsx", index=False)
# 调用函数并保存结果
# df_workdays = generate_workdays("2022-09-04", "2023-01-20")
# df_workdays.to_excel("工作日下.xlsx", index=False)
