# 匹配体质健康数据
# 读取“21级体测模板”目录下的所有xls文件，将数据合并到一个文件中，数据标题是第三行
import pandas as pd
import os

# “体测成绩/2021.xls”文件中的数据
df_cj = pd.read_excel('体测成绩/2021.xls')
df_cj['性别'].replace({1: '男', 2: '女'}, inplace=True)
# 初始化一个空的DataFrame用于存储所有文件的数据
all_data = pd.DataFrame()

# 遍历指定目录下的所有xlsx文件
for filename in os.listdir("21级体测模板"):
    if filename.endswith(".xls"):
        # 读取文件内容，设置数据标题为第三行
        df = pd.read_excel(os.path.join("21级体测模板", filename), header=2)

        # 将处理后的数据添加到all_data中
        all_data = pd.concat([all_data, df])

# 将DataFrame写入Excel文件
all_data.to_excel('2021级体测模板.xlsx', sheet_name='合并数据', index=False)
