# 遍历指定目录下的所有文件，根据文件名中的关键字将文件移动到对应的文件夹中。
import os
import shutil

# 定义关键字
subject_keywords = ['语文', '数学', '外语', '物理', '化学', '生物', '政治', '历史', '地理', '体育', '声乐', '舞蹈',
                    '美术', '信息', '通用', '研学', '心理']
group_keywords = ['教研', '备课']
special_keywords = ['培训']

# 定义要搜索的目录
search_dir = "汇总"

# 使用os.walk遍历所有子目录和文件
file_paths = [os.path.join(root, file) for root, _, files in os.walk(search_dir) for file in files]

# 遍历所有文件
for file_path in file_paths:
    file_dir, file_name = os.path.split(file_path)
    # 检查文件名是否包含关键字
    for keyword in group_keywords:
        if keyword in file_name:
            # 如果文件名包含关键字，将文件移动到相应的目录
            dest_dir = os.path.join(file_dir, "教研组计划")
            os.makedirs(dest_dir, exist_ok=True)
            shutil.move(file_path, dest_dir)
            break
    else:
        for keyword in special_keywords:
            if keyword in file_name:
                dest_dir = os.path.join(file_dir, "培训计划")
                os.makedirs(dest_dir, exist_ok=True)
                shutil.move(file_path, dest_dir)
                break
        else:
            for keyword in subject_keywords:
                if keyword in file_name:
                    dest_dir = os.path.join(file_dir, keyword)
                    os.makedirs(dest_dir, exist_ok=True)
                    shutil.move(file_path, dest_dir)
                    break
