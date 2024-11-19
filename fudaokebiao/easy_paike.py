# 完整的排课软件的Python实现

import pandas as pd

file_path = '分工.xlsx'
df = pd.read_excel(file_path)

# Step 1: 构建教师与班级的映射
teacher_classes = {}

for index, row in df.iterrows():
    yuwen_teacher = row['语文']
    waiyu_teacher = row['外语']
    class_name = row['班级']

    if yuwen_teacher not in teacher_classes:
        teacher_classes[yuwen_teacher] = {'科目': '语文', '班级': []}
    teacher_classes[yuwen_teacher]['班级'].append(class_name)

    if waiyu_teacher not in teacher_classes:
        teacher_classes[waiyu_teacher] = {'科目': '外语', '班级': []}
    teacher_classes[waiyu_teacher]['班级'].append(class_name)

# Step 2: 初始化排课表，记录每个时间段的课程安排
schedule_first_period = {}  # 第一节课，key为班级，value为教师信息
schedule_second_period = {}  # 第二节课

for teacher, data in teacher_classes.items():
    classes = data['班级']
    num_classes = len(classes)

    # 如果教师只教1个班级，直接分配到第一节课
    if num_classes == 1:

        schedule_first_period[classes[0]] = [teacher]

    # 如果教师教2个班级，分配到不同节课
    elif num_classes == 2:
        schedule_first_period.append({'Teacher': teacher, 'Class': classes[0], 'Subject': data['subject']})
        schedule_second_period.append({'Teacher': teacher, 'Class': classes[1], 'Subject': data['subject']})

    # 如果教师教3个或更多班级，分配到两节课
    else:
        half = num_classes // 2
        # 分配前半部分到第一节课
        for cls in classes[:half]:
            schedule_first_period.append({'Teacher': teacher, 'Class': cls, 'Subject': data['subject']})
        # 分配后半部分到第二节课
        for cls in classes[half:]:
            schedule_second_period.append({'Teacher': teacher, 'Class': cls, 'Subject': data['subject']})

# Step 3: 输出排课结果
df_first_period = pd.DataFrame(schedule_first_period)
df_second_period = pd.DataFrame(schedule_second_period)
