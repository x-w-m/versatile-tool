# 工资短信生成
# 工资明细表中必须包含应发合计、应扣合计、实发工资三列。
import pandas as pd


# 读取并合并数据
def read_excel(file1, file2):
    # 读取通讯录表格
    contacts_df = pd.read_excel(file1, usecols=['姓名', '联系方式'])

    # 读取工资明细表格
    salary_details_df = pd.read_excel(file2, dtype=str)

    # 输出工资明细的相关信息
    print("工资明细中的员工数量（数据行数）:", salary_details_df.shape[0])  # 数据行数

    # 计算姓名列之后的列数（条目数量）
    name_column_index = salary_details_df.columns.get_loc('姓名')
    num_of_entries = len(salary_details_df.columns) - (name_column_index + 1)
    print("工资明细中的条目数量（姓名列后的列数）:", num_of_entries)

    # 输出各条目名称（除姓名后的列标题）
    entry_titles = salary_details_df.columns[name_column_index + 1:].tolist()
    print("各条目名称（姓名列后的列标题）:", entry_titles)

    # 确定应发和应扣工资的列
    columns = salary_details_df.columns.tolist()
    start_index = columns.index('姓名') + 1
    end_index = columns.index('应发合计')
    earning_titles_d = columns[start_index:end_index]

    deduction_start_index = columns.index('应发合计') + 1
    deduction_end_index = columns.index('应扣合计')
    deduction_titles_d = columns[deduction_start_index:deduction_end_index]

    # 合并通讯录和工资明细表格,右链接保留工资明细中的所有条目
    merged_df_d = pd.merge(contacts_df, salary_details_df, on='姓名', how='right')
    # 替换 NaN 值为 0
    merged_df_d = merged_df_d.fillna(0)
    return merged_df_d, earning_titles_d, deduction_titles_d


# 生成工资条文本
def generate_payslip(name, contact, details, month_d, earning_titles_d, deduction_titles_d):
    # earnings = [f"{title}{details[title]}元" for title in earning_titles if pd.notna(details[title])]
    # deductions = [f"{title}{details[title]}元" for title in deduction_titles if pd.notna(details[title])]
    # 针对应发工资的部分，只保留数值非空的项，并使用中文逗号和空格进行分隔
    earnings = '，'.join([f"{title}{details[title]}元" for title in earning_titles_d if details[title] != 0])

    # 针对应扣工资的部分，操作同上
    deductions = '，'.join([f"{title}{details[title]}元" for title in deduction_titles_d if details[title] != 0])

    payslip = (f"尊敬的{name}老师，您好，您{month_d}月份学校工资明细如下："
               f"应发合计{details['应发合计']}元：【{earnings}】；"
               f"应扣合计{details['应扣合计']}元：【{deductions}】；"
               f"实发工资{details['实发工资']}元。")
    return contact, payslip


if __name__ == '__main__':
    month = 9
    gzmx = f"{month}月工资明细.xlsx"
    txl = "通讯录.xlsx"
    merged_df, earning_titles, deduction_titles = read_excel(txl, gzmx)
    # 对每位员工生成工资条
    payslip_data = [generate_payslip(row['姓名'], row['联系方式'], row, month, earning_titles, deduction_titles)
                    for index, row in merged_df.iterrows()]

    # 创建一个 DataFrame 用于保存到 Excel
    payslip_df = pd.DataFrame(payslip_data, columns=['联系方式', '工资条'], dtype=str)
    # 输出合并后的工资条数量
    print("合并后的工资条数量:", len(payslip_df))

    # 将 DataFrame 保存到名为“工资短信.xlsx”的 Excel 文件中
    payslip_df.to_excel(f'{month}月工资短信.xlsx', index=False)

    print(f'工资条已保存到 "{month}月工资短信.xlsx"。')
