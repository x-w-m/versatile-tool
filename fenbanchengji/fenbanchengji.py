import pandas as pd

# Column names as provided
column_names = ['班级', '姓名', '考号', '总分分数', '总分班排名', '总分校排名',
                '语文', '数学', '外语', '物理', '历史', '化学', '生物', '政治', '地理']

# Load the Excel file
data = pd.ExcelFile("三次成绩.xlsx")

# Load individual sheets with the specified column names
oct_scores = pd.read_excel(data, '10月', skiprows=3, header=None, names=column_names)
mid_scores = pd.read_excel(data, '期中', skiprows=3, header=None, names=column_names)
dec_scores = pd.read_excel(data, '12月', skiprows=3, header=None, names=column_names)
attendance = pd.read_excel(data, '参考情况')


def update_attendance():
    # Extract exam numbers
    oct_exam_numbers = oct_scores['考号'].unique()
    mid_exam_numbers = mid_scores['考号'].unique()
    dec_exam_numbers = dec_scores['考号'].unique()

    # Update attendance
    attendance['10月参考'] = attendance['10月考号'].apply(lambda x: 1 if x in oct_exam_numbers else 0)
    attendance['期中参考'] = attendance['期中考号'].apply(lambda x: 1 if x in mid_exam_numbers else 0)
    attendance['12月参考'] = attendance['12月考号'].apply(lambda x: 1 if x in dec_exam_numbers else 0)

    # Save to a new Excel file
    # output_file = '参考情况.xlsx'
    # attendance.to_excel(output_file, index=False)

    return attendance


def calculate_weights(row):
    # Checking attendance and calculating weights
    attended_exams = [row['10月参考'], row['期中参考'], row['12月参考']]
    attended_count = sum(attended_exams)

    if attended_count == 3:
        # Attended all exams
        return 0.3, 0.4, 0.3
    elif attended_count == 1:
        # Attended only one exam
        return [float(attended) for attended in attended_exams]
    else:
        # Attended two exams
        if row['期中参考']:
            # Attended mid-term
            return (0.3 if row['10月参考'] else 0, 0.7, 0.3 if row['12月参考'] else 0)
        else:
            # Did not attend mid-term
            return (0.5 if row['10月参考'] else 0, 0, 0.5 if row['12月参考'] else 0)


def update_weights(attendance):
    # Calculate weights for each row
    weights = attendance.apply(calculate_weights, axis=1)

    # Update the attendance DataFrame
    attendance[['10月权重', '期中权重', '12月权重']] = pd.DataFrame(weights.tolist(), index=attendance.index)

    return attendance


# 科目列表
subjects = ['语文', '数学', '外语', '物理', '历史', '化学', '生物', '政治', '地理']


# 合并表格并计算复合成绩
def calculate_composite_score(scores, attendance, exam_col, weight_col):
    scores.fillna(0, inplace=True)
    merged = pd.merge(attendance, scores[['考号'] + subjects], left_on=exam_col, right_on='考号', how="left")
    for subject in subjects:
        merged[subject] = merged[subject] * merged[weight_col]
    return merged[['班级', '姓名', '学号'] + subjects]


attendance = update_attendance()
updated_attendance = update_weights(attendance)
# 分别计算每次考试的复合成绩
oct_composite = calculate_composite_score(oct_scores, updated_attendance, '10月考号', '10月权重')
mid_composite = calculate_composite_score(mid_scores, updated_attendance, '期中考号', '期中权重')
dec_composite = calculate_composite_score(dec_scores, updated_attendance, '12月考号', '12月权重')

# 合并加权成绩
final_composite = pd.concat([oct_composite, mid_composite, dec_composite])
print(final_composite.head())
final_composite = final_composite.groupby(['班级', '姓名', '学号']).sum().reset_index()

updated_attendance.to_excel('final_output.xlsx', index=False)
final_composite.to_excel('final_weighted_scores.xlsx', index=False)
