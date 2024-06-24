# 对体测数据结果进行重命名和整理
# eeid体测数据目录下有多个文件夹，每个文件夹下有多个文件，文件名为班级名
# 子目录名格式为学年名
# 文件名中班级名原为：”高中2021级01班.xls“现改为”2101.xls“

import os
import shutil

main_dir = 'eeid体测数据'  # replace with your main directory

# iterate over all directories in the main directory
for school_year in os.listdir(main_dir):
    school_year_dir = os.path.join(main_dir, school_year)

    # ensure it's a directory
    if os.path.isdir(school_year_dir):
        # iterate over all files in the school year directory
        for file in os.listdir(school_year_dir):
            # ensure it's a file
            if os.path.isfile(os.path.join(school_year_dir, file)):
                # extract class name from file name
                class_name = file[4:6] + file[7:9]

                # create new directory if it doesn't exist
                new_dir = os.path.join(main_dir, class_name)
                if not os.path.exists(new_dir):
                    os.makedirs(new_dir)

                # construct new file name and path
                new_file_name = class_name + "_" + school_year + '.xls'
                new_file_path = os.path.join(new_dir, new_file_name)

                # move and rename the file
                shutil.move(os.path.join(school_year_dir, file), new_file_path)
