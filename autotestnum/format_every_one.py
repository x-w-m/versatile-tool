# 使用pandas将去向表和考室座次表进行拆分
# 然后使用openpyxl对拆分后的表格进行格式调整
import pandas as pd

from autotestnum import formatexcel


def format_every_one():
    df = pd.read_excel("考生去向表/考生去向表.xlsx", sheet_name=0, dtype=str)
    # 插入一列“年级”到第一列，值为“高一”
    df.insert(0, "年级", "高一")
    # 删除“分数”列
    df.drop(columns=["分数"], inplace=True)

    # 找出需要拆分的列的所有唯一值
    unique_values = df['班级'].unique()
    # 创建一个新的ExcelWriter对象
    with pd.ExcelWriter('考生去向表/考生去向表_拆分.xlsx') as writer:
        # 对于每个唯一值，创建一个新的DataFrame，其中只包含该值的行
        for value in unique_values:
            new_df = df[df['班级'] == value]
            # 使用pandas的to_excel函数将每个新的DataFrame作为一个新的工作表写入Excel文件
            new_df.to_excel(writer, sheet_name=str(value) + " C", index=False)

    # 根据考号排序
    df = df.sort_values("考号")
    # "考号"列的所有唯一值
    unique_values = df['考室号'].unique()
    # 创建一个新的ExcelWriter对象
    with pd.ExcelWriter('考生去向表/考室座次表_拆分.xlsx') as writer:
        # 对于每个唯一值，创建一个新的DataFrame，其中只包含该值的行
        for value in unique_values:
            new_df = df[df['考室号'] == value]
            # 使用pandas的to_excel函数将每个新的DataFrame作为一个新的工作表写入Excel文件
            new_df.to_excel(writer, sheet_name=str(value).zfill(2) + "考室", index=False)
    formatexcel.excel_format("考生去向表/考生去向表_拆分.xlsx")
    formatexcel.excel_format("考生去向表/考室座次表_拆分.xlsx")
