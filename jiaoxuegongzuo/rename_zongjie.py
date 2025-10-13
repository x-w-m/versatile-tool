# 检查工作总结文件名，并重命名
import pandas as pd
import os
from pypinyin import lazy_pinyin
from fuzzywuzzy import process

# 读取Excel文件
df = pd.read_excel('教师信息.xlsx')

# 定义要搜索的目录
search_dir = "240916"

# 使用os.walk遍历所有子目录和文件
file_names = [os.path.join(root, file) for root, _, files in os.walk(search_dir) for file in files]

# 定义关键字
grade_keywords = ['高一', '高二', '高三']
subject_keywords = ['语文', '数学', '外语', '英语', '物理', '化学', '生物', '政治', '历史', '体育', '地理', '声乐',
                    '舞蹈', '美术', '信息技术', '通用', '研学']
group_keywords = ['教研组', '备课组', '实验室', '培训', '高考班', '学考班']
type_keywords = ['工作总结', '教学总结', '工作计划', '教学计划']

# 从DataFrame中获取名称
name_keywords = df['姓名'].tolist()

# 遍历所有文件名
for file_name_d in file_names:
    file_dir, file_name = os.path.split(file_name_d)
    # Create a dictionary to store the keywords found in each category
    found_keywords = {'name': '', 'grade': '', 'subject': ''}
    file_keywords = {'grade': '', 'subject': '', 'group': '', 'type': ''}

    # 检查文件名是否包含任何关键字
    for keyword in name_keywords:
        if keyword in file_name:
            found_keywords['name'] = keyword
            # 从DataFrame中获取匹配名称的年级和科目
            matched_row = df[df['姓名'] == found_keywords['name']]
            found_keywords['grade'] = matched_row['年级'].values[0]
            found_keywords['subject'] = matched_row['科目'].values[0]
            break

    if not found_keywords['name']:
        name_keywords_pinyin = [''.join(lazy_pinyin(name)) for name in name_keywords]
        # Convert the file name to pinyin
        file_name_pinyin = ''.join(lazy_pinyin(file_name))
        # If no exact match was found, use fuzzywuzzy to find the closest matches
        closest_matches = process.extract(file_name_pinyin, name_keywords_pinyin, limit=10)
        # 过滤掉匹配度低于50的结果
        closest_matches = [match for match in closest_matches if match[1] >= 50]
        if closest_matches:
            print(f"原文件名: {file_name}")
            print("0: 跳过")
            for i, match in enumerate(closest_matches, start=1):
                # Get the index of the match
                match_index = name_keywords_pinyin.index(match[0])
                # Get the corresponding name from the original name list
                print(f"{i}:{name_keywords[match_index]}", end=" ")
            choice = input("请选择最接近的姓名，或直接输入姓名，或输入0跳过: ")
            if choice.isdigit():
                choice = int(choice)
                if choice == 0:
                    print(f"跳过文件: {file_name}")
                    continue
                else:
                    choice -= 1
                    found_keywords['name'] = name_keywords[name_keywords_pinyin.index(closest_matches[choice][0])]
            else:
                found_keywords['name'] = choice
            # Get the grade and subject for the matched name from the DataFrame
            matched_row = df[df['姓名'] == found_keywords['name']]
            found_keywords['grade'] = matched_row['年级'].values[0]
            found_keywords['subject'] = matched_row['科目'].values[0]

    # Check if the file name contains any of the grade, subject, and type keywords
    for keyword in grade_keywords:
        if keyword in file_name:
            file_keywords['grade'] = keyword
    for keyword in subject_keywords:
        if keyword in file_name:
            file_keywords['subject'] = keyword
    for keyword in group_keywords:
        if keyword in file_name:
            file_keywords['group'] = keyword
    for keyword in type_keywords:
        if keyword in file_name:
            file_keywords['type'] = keyword

    # If all types are not found, record it and do not rename
    if not found_keywords['name']:
        print(f"未找到匹配的姓名: {file_name}")
        continue

    # Print whether the grade and subject found in the file name match the grade and subject for the matched name
    print(f"年级匹配: {found_keywords['grade'] == file_keywords['grade']}")
    print(f"科目匹配: {found_keywords['subject'] == file_keywords['subject']}")

    for keyword in type_keywords:
        if keyword in file_name:
            file_keywords['type'] = keyword

    # Get the extension of the original file

    _, ext = os.path.splitext(file_name)

    # Create the new file name by joining the found keywords and adding the original file extension
    print(found_keywords['name'])
    new_file_name = "_".join(
        [str(found_keywords['grade']), found_keywords['name'], str(found_keywords['subject']), file_keywords['group'],
         file_keywords['type']]) + ext

    # If the new file name is the same as the original file name, skip it
    if file_name == new_file_name:
        continue

    counter = 1
    new_file_base, new_file_ext = os.path.splitext(new_file_name)
    original_file_name = new_file_name
    while os.path.exists(os.path.join(file_dir, new_file_name)):
        new_file_name = f"{new_file_base}_{counter}_{new_file_ext}"
        counter += 1

    # Rename the file
    os.rename(os.path.join(file_dir, file_name), os.path.join(file_dir, new_file_name))
    print(f"新文件名: {new_file_name}")
