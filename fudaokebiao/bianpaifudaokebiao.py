import pandas as pd
from collections import defaultdict


# 读取Excel数据
def read_teacher_data(file_path):
    df = pd.read_excel(file_path)
    return df


# 初始化时间表
def initialize_schedule(classes):
    schedule = defaultdict(lambda: defaultdict(lambda: ['自习', '自习']))
    return schedule


# 检测潜在的冲突并生成合班方案
def analyze_conflicts(teacher_data, class_subject_mapping):
    conflict_info = defaultdict(lambda: defaultdict(list))

    for day, subjects in class_subject_mapping.items():
        available_teachers = teacher_data[teacher_data['科目'].isin(subjects)]

        for subject in subjects:
            teachers_for_subject = available_teachers[available_teachers['科目'] == subject]

            for _, row in teachers_for_subject.iterrows():
                classes = row['班级'].split(',')
                for cls in classes:
                    conflict_info[day][cls].append(subject)

    return conflict_info


# 按照合班方案分配课程
def assign_courses_with_conflict_handling(schedule, teacher_data, class_subject_mapping, conflict_info):
    for day, subjects in class_subject_mapping.items():
        available_teachers = teacher_data[teacher_data['科目'].isin(subjects)]

        for subject in subjects:
            teachers_for_subject = available_teachers[available_teachers['科目'] == subject]
            for _, row in teachers_for_subject.iterrows():
                classes = row['班级'].split(',')

                if len(classes) == 1:
                    # 如果只有一个班级，直接安排课程
                    schedule[day][classes[0]][0] = subject
                    schedule[day][classes[0]][1] = subject
                elif len(classes) == 2:
                    # 如果有两个班级，分别安排在第一节和第二节
                    schedule[day][classes[0]][0] = subject
                    schedule[day][classes[1]][1] = subject
                else:
                    # 三个及以上班级处理合班情况
                    slot_assigned = False
                    for slot in [0, 1]:
                        if all(schedule[day][cls][slot] == '自习' for cls in classes):
                            for cls in classes:
                                schedule[day][cls][slot] = f"{subject}（合班）"
                            slot_assigned = True
                            break

                    if not slot_assigned:
                        for cls in classes:
                            if '自习' in schedule[day][cls]:
                                slot = schedule[day][cls].index('自习')
                                schedule[day][cls][slot] = subject


# 保存课表到Excel文件
def save_schedule_to_excel(schedule, classes, file_path):
    days = ['星期一', '星期二', '星期三', '星期四', '星期五']
    columns = pd.MultiIndex.from_product([days, ['第1节', '第2节']], names=['星期', '节次'])

    # 初始化空的DataFrame
    df = pd.DataFrame(index=classes, columns=columns)

    # 填充DataFrame
    for day in days:
        for class_name in classes:
            df.loc[class_name, (day, '第1节')] = schedule[day][class_name][0]
            df.loc[class_name, (day, '第2节')] = schedule[day][class_name][1]

    # 将DataFrame保存到Excel
    df.to_excel(file_path)


# 主函数
def main():
    file_path = '教师信息高二.xlsx'
    output_file = 'schedule.xlsx'
    teacher_data = read_teacher_data(file_path)

    # 所有班级
    classes = set()
    for class_list in teacher_data['班级']:
        classes.update(class_list.split(','))
    classes = list(classes)

    # 每天的课程安排规则
    class_subject_mapping = {
        '星期一': ['语文', '外语'],
        '星期二': ['数学'],
        '星期三': ['物理', '历史'],
        '星期四': ['化学', '政治'],
        '星期五': ['生物', '地理']
    }

    # 初始化时间表
    schedule = initialize_schedule(classes)

    # 冲突分析
    conflict_info = analyze_conflicts(teacher_data, class_subject_mapping)

    # 分配课程并处理冲突
    assign_courses_with_conflict_handling(schedule, teacher_data, class_subject_mapping, conflict_info)

    # 保存课表到Excel
    save_schedule_to_excel(schedule, classes, output_file)
    print(f"课表已保存到 {output_file}")


if __name__ == "__main__":
    main()
