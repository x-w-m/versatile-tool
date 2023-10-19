# 自动编排考号
# 需准备考室安排表及参考名单表两个表。
# 考室安排表需要有三列：考室号、楼层、人数、科目组。
# 参考名单表包含基础三列：班级、姓名、科目组、分数；以及待生成三列：考生号、考室号、座位号、楼层。


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

    # 遍历每个科目组
    for group in df_room["科目组"].unique():
        # 筛选出该科目组对应的考室和学生
        df_room_group = df_room[df_room["科目组"] == group]
        df_student_group = df_student[df_student["科目组"] == group]
        # 初始化一个变量，用于记录已经安排了多少个学生
        count = 0
        # 遍历该科目组对应的每个考室
        for index, row in df_room_group.iterrows():
            # 获取该考室的编号和人数
            room_number = row["考室号"]
            room_size = row["人数"]
            room_position = row["楼层"]
            # 筛选出该考室需要安排的学生
            df_student_room = df_student_group.iloc[count:count + room_size]
            # 给每个学生分配座位号，并生成考号
            seat_number = 1
            for i, r in df_student_room.iterrows():
                # 在迭代过程中，r是一个Series对象，它是从df_student_room中复制出来的一行数据，
                # 对r所做的修改不会反映到原始的DataFrame对象df_student_room中。
                # 想要修改原始的DataFrame对象，可以使用.loc[]方法来定位行和列，并直接修改DataFrame对象中的值。
                # 或者将被修改的行“r”保存到列表中，最后将其转换为新的DataFrame对象。
                r["座位号"] = str(seat_number).zfill(2)  # 座位号用两位数表示，不足补零
                r["考室号"] = str(room_number).zfill(2).zfill(2)
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
    pass


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
