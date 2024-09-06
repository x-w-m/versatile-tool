# 检查未交名单
import pandas as pd
import os
from pypinyin import lazy_pinyin
from fuzzywuzzy import process

# Read the Excel file
df = pd.read_excel('教师任课信息.xlsx')

# Define the directory you want to search
search_dir = "24年上学期教学总结汇总"

# Use os.walk to iterate through all subdirectories and files
file_names = [os.path.join(root, file) for root, _, files in os.walk(search_dir) for file in files]
# Extract names from file names
name_from_file = [os.path.splitext(os.path.basename(file_name))[0] for file_name in file_names]
print(name_from_file)
# Mark names that appear in the file names in the DataFrame
df['是否提交'] = df['教师'].apply(lambda name: any(name in file_name for file_name in name_from_file))
df.to_excel('未交名单.xlsx', index=False)
