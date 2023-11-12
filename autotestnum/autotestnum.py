# 自动编排考号
# 需准备考室安排表及参考名单表两个表。
# 考室安排表需要有三列：考室号、楼层、人数、科目组。
# 参考名单表包含基础三列：班级、姓名、科目组、分数；以及待生成四列：考生号、考室号、座位号、楼层。


# 导入所需的库
import pandas as pd

# 读取两个excel文件中的数据，存储为pandas的DataFrame对象
df_room = pd.read_excel("考室安排表.xlsx")  # 考室安排表
df_student = pd.read_excel("参考名单表.xlsx", dtype=str)  # 参考名单表

# 按照科目组和分数对学生进行排序
df_student.sort_values(by=["科目组", "分数"], ascending=[True, False], inplace=True)


def mode_for_one():
    kshp = input("请输入考生号前缀：")
    # 定义一个列表，用于保存已经编排好的考生信息。
    rows = []
    # 新建一个字典，用于跟踪每个考室已经分配的最后一个座位号
    room_seat_tracker = {}

    # 遍历每个科目组
    for group in df_room["科目组"].unique():
        # 筛选出该科目组对应的考室和学生
        df_room_group = df_room[df_room["科目组"] == group]
        df_student_group = df_student[df_student["科目组"] == group]
        # 初始化一个变量，用于记录已经安排了多少个学生
        count = 0
        # 遍历该科目组对应的每个考室
        for index, room_row in df_room_group.iterrows():
            # 获取该考室的编号和人数
            room_number = room_row["考室号"]
            room_size = room_row["人数"]
            room_position = room_row["楼层"]

            # 如果考室号在tracker中没有记录，说明是新的考室或者新的科目组的开始，座位号从1开始
            if room_number not in room_seat_tracker:
                room_seat_tracker[room_number] = 0

            # 从之前的座位号继续开始分配
            start_seat_number = room_seat_tracker[room_number] + 1

            # 筛选出该考室需要安排的学生
            df_student_room = df_student_group.iloc[count:count + room_size]
            # 给每个学生分配座位号，并生成考号
            seat_number = start_seat_number  # 设置本次循环开始时的座位号。
            for i, r in df_student_room.iterrows():
                # 在迭代过程中，r是一个Series对象，它是从df_student_room中复制出来的一行数据，
                # 对r所做的修改不会反映到原始的DataFrame对象df_student_room中。
                # 想要修改原始的DataFrame对象，可以使用.loc[]方法来定位行和列，并直接修改DataFrame对象中的值。
                # 或者将被修改的行“r”保存到列表中，最后将其转换为新的DataFrame对象。
                r["座位号"] = str(seat_number).zfill(2)  # 座位号用两位数表示，不足补零
                r["考室号"] = str(room_number).zfill(2)
                r["考生号"] = kshp + str(room_number).zfill(2) + str(seat_number).zfill(2)  # 考号由“1000”、考室号、座位号拼接而成
                r["楼层"] = room_position

                # 使用.loc[]方法来定位行和列，从而直接修改DataFrame对象中的值。
                # 此方法会出现SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
                # 暂时没有找到解决方法
                # 座位号用两位数表示，不足补零
                # df_student_room.loc[i, "座位号"] = str(seat_number).zfill(2)
                # df_student_room.loc[i, "考室号"] = str(room_number).zfill(2)
                # # 考号由“1000”、考室号、座位号拼接而成
                # df_student_room.loc[i, "考生号"] = "1000" + str(room_number).zfill(2) + str(seat_number).zfill(2)

                seat_number += 1
                # 将编排后的行添加到列表中
                rows.append(r)
            # 将该考室安排好的学生添加到结果中
            # df_result = pd.concat([df_result, df_student_room])
            # 更新该考室的最后一个座位号
            room_seat_tracker[room_number] = seat_number - 1
            # 更新已经安排了多少个学生的变量
            count += room_size

    # 将rows列表转换为DataFrame对象，用于存储最终的结果。
    df_result = pd.DataFrame(rows, columns=df_student.columns, dtype=str)

    # 将结果按照班级、姓名进行排序，并重置索引
    df_result.sort_values(by=["班级", "考生号"], inplace=True)
    df_result.reset_index(drop=True, inplace=True)

    # 将结果保存为一个新的excel文件，命名为“考生去向表.xlsx”
    df_result.to_excel("考生去向表.xlsx", index=False)

    # 打印一条提示信息，表示程序运行结束
    print("程序运行结束，已经生成“考生去向表.xlsx”文件，请查看。")


def mode_for_two():
    # Prompt user for the names of the experimental classes
    experimental_classes = input("请输入实验班班级名称，多个班级使用“、”分隔：").split("、")
    # Prompt user for the corresponding rooms
    assigned_rooms = input("请输入需要安排到的考室，多个考室使用“、”分隔：").split("、")
    kshp = input("请输入考生号前缀：")
    # Assuming the data consistency is already ensured, proceed with the assignment
    # Define a list to save the arranged student information
    rows = []
    # Initialize a dictionary to track the last seat number assigned in each room
    room_seat_tracker = {}

    # Iterate over the experimental classes and the corresponding rooms
    for class_name, room_number in zip(experimental_classes, assigned_rooms):
        # Filter the students belonging to the current experimental class
        df_student_class = df_student[df_student["班级"] == class_name]
        # Filter the room information for the current assigned room
        df_room_assigned = df_room[df_room["考室号"] == room_number]

        # If the room has not been assigned before, start seat numbering from 1
        if room_number not in room_seat_tracker:
            room_seat_tracker[room_number] = 0

        # Continue seat assignment from the last seat number for this room
        start_seat_number = room_seat_tracker[room_number] + 1

        # Assign a seat number and generate exam number for each student
        seat_number = start_seat_number
        for index, student_row in df_student_class.iterrows():
            # Assign seat number, room number, and exam number to the student
            student_row["座位号"] = str(seat_number).zfill(2)  # Seat number with leading zero
            student_row["考室号"] = str(room_number).zfill(2)  # Room number with leading zero
            # Exam number consists of a prefix, room number, and seat number
            student_row["考生号"] = kshp + str(room_number).zfill(2) + str(seat_number).zfill(2)
            student_row["楼层"] = df_room_assigned.iloc[0]["楼层"]  # Assuming one floor per room

            # Increment the seat number for the next student
            seat_number += 1
            # Append the arranged student information to the list
            rows.append(student_row)

        # Update the last seat number assigned in the room
        room_seat_tracker[room_number] = seat_number - 1

    # Convert the list of arranged students to a DataFrame
    df_result = pd.DataFrame(rows, columns=df_student.columns, dtype=str)
    # Sort the result by class and exam number, then reset the index
    df_result.sort_values(by=["班级", "考生号"], inplace=True)
    df_result.reset_index(drop=True, inplace=True)

    # Save the result to a new Excel file
    df_result.to_excel("考生去向表.xlsx", index=False)

    # Print a message indicating the end of the process
    print("程序运行结束，已经生成“考生去向表.xlsx”文件，请查看。")


if __name__ == '__main__':
    # 使用提示
    print("# 自动编排考号")
    print("# 需准备\"考室安排表.xlsx\"及\"参考名单表.xlsx\"两个表。")
    print("# 考室安排表需要有三列：考室号、楼层、人数、科目组。")
    print("# 参考名单表包含基础三列：班级、姓名、科目组、分数；以及待生成三列：考生号、考室号、座位号、楼层。")

    # 模式选择
    mode = input("准备好后请使用数字键选择模式：\n1：简单科目组模式；\n2：实验班指定考室模式。\n")
    if mode == "1":
        mode_for_one()
    elif mode == "2":
        mode_for_two()
