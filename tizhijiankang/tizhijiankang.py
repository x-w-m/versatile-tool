# 对学生体测数据无法识别的部分进行数据清洗
import random
import re

import pandas as pd

data = pd.read_excel("隆回第一中学体测-数据清洗.xlsx", dtype={"50米跑": str, "立定跳远": str})


def clean_height(height, gender):
    try:
        # 尝试将字符串转换为浮点数
        height = float(height)

        # 将以米为单位的数值转换为厘米
        if 1.0 <= height < 2.1:
            height = height * 100

        # 处理可能由于疏忽而忘记前面的“1”的情况
        if 55 <= height <= 85:
            height = float('1' + str(height))

        if 1000 <= height < 2100:
            height = height / 10

        # 检查转换后的数值是否在合理范围内
        if 100 <= height <= 210:
            return height
        else:
            raise ValueError  # 触发异常，以便生成基于性别的随机数
    except ValueError:
        # 对于无法识别为数字的数据或不在合理范围内的数据，根据性别生成随机数
        if gender == 1:  # 男生
            return random.randint(150, 180)
        elif gender == 2:  # 女生
            return random.randint(140, 160)
        else:
            return random.randint(140, 190)  # 对于性别未知的情况


# 应用清洗函数到'身高'列，并保存结果到新列
data['身高_清洗'] = data.apply(lambda row: clean_height(row['身高'], row['性别']), axis=1)


def estimate_weight(height, gender):
    try:
        height = float(height)
        # 根据性别和身高估计体重
        if gender == 1:  # 男性
            estimated_weight = random.uniform(height - 100 - 10, height - 100 + 10)
        elif gender == 2:  # 女性
            estimated_weight = random.uniform(height - 105 - 10, height - 105 + 10)
        else:
            # 身高数据错误或性别未知，随机生成45到65之间的体重
            estimated_weight = random.uniform(45, 65)

        # 四舍五入到一位小数
        return round(estimated_weight, 0)
    except ValueError:
        # 身高数据错误，随机生成45到65之间的体重
        return round(random.uniform(45, 65), 0)


def clean_weight(weight, height, gender):
    try:
        # 尝试将字符串转换为浮点数
        weight = float(weight)

        # 检查转换后的数值是否在合理范围内
        if 30 <= weight <= 150:
            return weight
        else:
            raise ValueError  # 触发异常，以便生成基于身高和性别的体重
    except ValueError:
        # 对于无法识别为数字的数据或不在合理范围内的数据，根据身高和性别估算体重
        return estimate_weight(height, gender)


# 应用清洗函数到'体重'列，并保存结果到新列
data['体重_清洗'] = data.apply(lambda row: clean_weight(row['体重'], row['身高'], row['性别']), axis=1)


def clean_vital_capacity(vc):
    try:
        # 尝试将字符串转换为整数
        vc = int(float(vc))

        # 检查转换后的数值是否在合理范围内
        if 1800 <= vc <= 9999:
            return vc
        else:
            raise ValueError  # 触发异常，以便生成随机数
    except ValueError:
        # 对于无法识别为数字的数据或不在合理范围内的数据，随机生成一个数
        return random.randint(1900, 3500)


# 应用清洗函数到'肺活量'列，并保存结果到新列
data['肺活量_清洗'] = data['肺活量'].apply(clean_vital_capacity)


def clean_50m_run(time):
    if isinstance(time, float) or isinstance(time, int):
        if 6.0 <= time <= 15.0:
            return time
        else:
            return round(random.uniform(6.0, 15.0), 2)
    else:
        time_str = str(time)
        # 使用正则表达式提取所有数字
        numbers = re.findall(r'\d+', time_str)

        if len(numbers) == 1:
            # 如果只有一个数字，直接使用这个数字作为秒数
            time = float(numbers[0])
        elif len(numbers) >= 2:
            # 如果有两个或更多数字，将第二个数字作为小数部分
            time_str = numbers[0] + '.' + numbers[1]
            time = float(time_str)
        else:
            return round(random.uniform(6.0, 15.0), 2)

        # 检查时间是否在合理范围内
        if 6.0 <= time <= 15.0:
            return time
        else:
            return round(random.uniform(6.0, 15.0), 2)


# 应用清洗函数到'50米跑'列，并保存结果到新列
data['50米跑_清洗'] = data['50米跑'].apply(clean_50m_run)

# 应用清洗函数到'50米跑'列，并保存结果到新列
data['50米跑_清洗'] = data['50米跑'].apply(clean_50m_run)


def clean_standing_jump(distance):
    # 确保数据是字符串类型
    distance_str = str(distance)

    # 将字母l、i和o替换为数字1和0，忽略大小写
    distance_str = distance_str.replace('l', '1').replace('i', '1').replace('o', '0')
    distance_str = distance_str.replace('L', '1').replace('I', '1').replace('O', '0')

    try:
        # 尝试将字符串转换为浮点数
        distance = float(distance_str)

        # 将以米为单位的数值转换为厘米
        if 0.8 <= distance < 4.0:
            distance = distance * 100
        # 将以毫米为单位的数值转换为厘米
        elif 1000 <= distance < 10000:
            distance = distance / 10
        # 对于大于等于10000的数值，截取前三位数字
        elif distance >= 10000:
            distance = int(str(distance)[:3])

        # 检查转换后的数值是否在合理范围内
        if 80 <= distance <= 400:
            return int(distance)
        else:
            return random.randint(150, 250)
    except ValueError:
        # 对于无法识别为数字的数据，随机生成一个150到250之间的整数
        return random.randint(150, 250)


# 应用清洗函数到'立定跳远'列，并保存结果到新列
data['立定跳远_清洗'] = data['立定跳远'].apply(clean_standing_jump)


def clean_sit_and_reach(distance):
    # 确保数据是字符串类型
    distance_str = str(distance)

    # 将字母l、i和o替换为数字1和0，忽略大小写
    distance_str = distance_str.replace('l', '1').replace('i', '1').replace('o', '0')
    distance_str = distance_str.replace('L', '1').replace('I', '1').replace('O', '0')

    try:
        # 尝试将字符串转换为浮点数
        distance = float(distance_str)

        # 对于太小的数值乘以10
        if distance < 5:
            distance *= 10
        # 对于太大的数值，截取前两位数
        elif distance >= 100:
            distance = int(str(distance)[:2])

        # 检查转换后的数值是否在合理范围内
        if 5 <= distance <= 40:
            return round(distance, 1)
        else:
            # 生成5到25之间的随机整数
            return random.randint(5, 25)
    except ValueError:
        # 对于无法识别为数字的数据，随机生成一个5到25之间的整数
        return random.randint(5, 25)


# 应用清洗函数到 '立定跳远' 列，并保存结果到新列
data['坐位体前屈_清洗'] = data['坐位体前屈'].apply(clean_sit_and_reach)


def clean_800m_run(time_record, gender):
    # 如果性别为男，则成绩为空
    if gender == 1:
        return None

    # 确保时间记录是字符串类型
    time_str = str(time_record)

    # 使用正则表达式提取数字
    numbers = re.findall(r'\d+', time_str)

    # 根据数字数量和内容确定时间起始点
    if len(numbers) == 3 and numbers[0] == '0':  # 秒和毫秒格式
        minutes, seconds = int(numbers[1]), int(numbers[2])
    elif len(numbers) >= 2:  # 分和秒格式
        minutes, seconds = int(numbers[0]), int(numbers[1])
    else:
        # 数据不合理时，随机生成一个合理的时间
        return f"{random.choice([3, 4])}'{str(random.randint(0, 59)).zfill(2)}\""

    # 分别处理分钟和秒钟的合理性
    minutes_valid = minutes in [3, 4, 5, 6]
    seconds_valid = 0 <= seconds < 60

    if not minutes_valid:
        minutes = random.choice([3, 4])
    if not seconds_valid:
        seconds = random.randint(0, 59)

    # 格式化秒数，确保它始终有两位数字
    seconds_str = str(seconds).zfill(2)

    return f"{minutes}'{seconds_str}\""


# 假设 data 是您的 DataFrame
# 应用清洗函数到 '800米跑' 列，并保存结果到新列
data['800米跑_清洗'] = data.apply(lambda row: clean_800m_run(row['800米跑'], row['性别']), axis=1)


def clean_1000m_run(time_record, gender):
    # 如果性别为女，则成绩为空
    if gender != 1:
        return None

    # 确保时间记录是字符串类型
    time_str = str(time_record)

    # 使用正则表达式提取数字
    numbers = re.findall(r'\d+', time_str)

    # 根据数字数量和内容确定时间起始点
    if len(numbers) == 3 and numbers[0] == '0':  # 秒和毫秒格式
        minutes, seconds = int(numbers[1]), int(numbers[2])
    elif len(numbers) >= 2:  # 分和秒格式
        minutes, seconds = int(numbers[0]), int(numbers[1])
    else:
        # 数据不合理时，随机生成一个合理的时间
        return f"{random.choice([3, 4])}'{str(random.randint(0, 59)).zfill(2)}\""

    # 分别处理分钟和秒钟的合理性
    minutes_valid = minutes in [3, 4, 5, 6]
    seconds_valid = 0 <= seconds < 60

    if not minutes_valid:
        minutes = random.choice([3, 4])
    if not seconds_valid:
        seconds = random.randint(0, 59)

    # 格式化秒数，确保它始终有两位数字
    seconds_str = str(seconds).zfill(2)

    return f"{minutes}'{seconds_str}\""


# 假设 data 是您的 DataFrame
# 应用清洗函数到 '1000米跑' 列，并保存结果到新列
data['1000米跑_清洗'] = data.apply(lambda row: clean_1000m_run(row['1000米跑'], row['性别']), axis=1)


def clean_sit_ups(sit_ups_count, gender):
    # 如果性别不是女性（2），则成绩为空
    if gender != 2:
        return None

    try:
        # 尝试将数据转换为整数
        sit_ups_count = int(sit_ups_count)
        # 检查数据是否在0到99之间
        if 0 <= sit_ups_count <= 99:
            return sit_ups_count
        else:
            # 如果不在范围内，则生成一个25到45之间的随机整数
            return random.randint(25, 45)
    except ValueError:
        # 如果数据不是整数，同样生成一个25到45之间的随机整数
        return random.randint(25, 45)


# 假设 data 是您的 DataFrame
# 应用清洗函数到 '一分钟仰卧起坐' 列，并保存结果到新列
data['一分钟仰卧起坐_清洗'] = data.apply(lambda row: clean_sit_ups(row['一分钟仰卧起坐'], row['性别']), axis=1)


def clean_pull_ups(pull_ups, gender):
    if gender != 1:
        return None
    try:
        # 尝试将数据转换为整数
        pull_ups_count = int(pull_ups)
    except ValueError:
        # 如果不是整数，提取数字
        numbers = re.findall(r'\d+', str(pull_ups))
        if numbers:
            # 尝试使用提取的前1位或2位数字
            pull_ups_count = int(numbers[0][:2])
        else:
            # 如果没有数字，随机生成一个8到16之间的整数
            return random.randint(8, 16)

    # 检查数据是否在5到35之间
    if 5 <= pull_ups_count <= 35:
        return pull_ups_count
    else:
        # 如果不在范围内，则生成一个8到16之间的随机整数
        return random.randint(8, 16)


# 假设 data 是您的 DataFrame
# 应用清洗函数到 '引体向上' 列，并保存结果到新列
data['引体向上_清洗'] = data.apply(lambda row: clean_pull_ups(row['引体向上'], row['性别']), axis=1)

# 显示结果
print(data[['引体向上', '引体向上_清洗']])
# 可以选择保存整个DataFrame或仅保存需要的列
data.to_excel('体测成绩_清洗后.xlsx')
print("清洗完成")
