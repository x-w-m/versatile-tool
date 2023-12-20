# 自动编排考号
# 需准备考室安排表及参考名单表两个表。
# 考室安排表需要有三列：考室号、教室、楼层、人数、科目组。
# 参考名单表包含基础三列：班级、姓名、科目组、分数；以及待生成四列：考号、考室号、座位号、楼层、教室。
import sys

# 导入所需的库
import pandas as pd
from colorama import init, Fore, Style


def mode_for_one():
    # 读取两个excel文件中的数据，存储为pandas的DataFrame对象
    df_room = pd.read_excel("考室安排表.xlsx")  # 考室安排表
    df_student = pd.read_excel("参考名单表.xlsx", dtype=str)  # 参考名单表

    # 按照科目组和分数对学生进行排序
    df_student.sort_values(by=["科目组", "分数"], ascending=[True, False], inplace=True)
    kshp = input("请输入考号前缀：")
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
            room = room_row["教室"]
            room_position = room_row["楼层"]
            room_size = room_row["人数"]

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
                r["考号"] = kshp + str(room_number).zfill(2) + str(seat_number).zfill(2)  # 考号由“1000”、考室号、座位号拼接而成
                r["考室号"] = str(room_number).zfill(2)
                r["座位号"] = str(seat_number).zfill(2)  # 座位号用两位数表示，不足补零
                r["楼层"] = room_position
                r["教室"] = room

                # 使用.loc[]方法来定位行和列，从而直接修改DataFrame对象中的值。
                # 此方法会出现SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
                # 暂时没有找到解决方法
                # 座位号用两位数表示，不足补零
                # df_student_room.loc[i, "座位号"] = str(seat_number).zfill(2)
                # df_student_room.loc[i, "考室号"] = str(room_number).zfill(2)
                # # 考号由“1000”、考室号、座位号拼接而成
                # df_student_room.loc[i, "考号"] = "1000" + str(room_number).zfill(2) + str(seat_number).zfill(2)

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
    df_result = pd.DataFrame(rows,
                             columns=["班级", "姓名", "科目组", "分数", "考号", "考室号", "座位号", "楼层", "教室"],
                             dtype=str)

    # 将结果按照班级、考号进行排序，并重置索引
    # inplace=True意味着排序操作将直接在原始的 df_result 对象上进行，默认为False，会返回新的DataFrame
    df_result.sort_values(by=["班级", "考号"], inplace=True)
    df_result.reset_index(drop=True, inplace=True)

    # 创建一个ExcelWriter对象
    with pd.ExcelWriter("考生去向表.xlsx") as writer:
        # 保存“考生去向表”为第一个工作表
        df_result.to_excel(writer, sheet_name="考生去向表", index=False)

        # 对df_result根据考号进行排序，为创建第二个工作表做准备
        df_seating = df_result.sort_values(by="考号")
        df_seating.reset_index(drop=True, inplace=True)

        # 保存排序后的数据为第二个工作表，这里命名为“考室座次表”
        df_seating.to_excel(writer, sheet_name="考室座次表", index=False)

    print("数据已保存至'考生去向表.xlsx'，包含'考生去向表'和'考室座次表'两个工作表。")


def generate_template(directory="模板文件"):
    # 在“模板文件”目录下创建考室安排表及参考名单表这两个表格的模板文件
    import os
    # 确保目录存在
    if not os.path.exists(directory):
        os.makedirs(directory)

    # 创建考室安排表模板
    df_room_template = pd.DataFrame(columns=["考室号", "教室", "楼层", "人数", "科目组"])
    room_template_path = os.path.join(directory, "考室安排表模板.xlsx")
    df_room_template.to_excel(room_template_path, index=False)

    # 创建参考名单表模板
    df_student_template = pd.DataFrame(
        columns=["班级", "姓名", "科目组", "分数", "考号", "考室号", "座位号", "楼层", "教室"])
    student_template_path = os.path.join(directory, "参考名单表模板.xlsx")
    df_student_template.to_excel(student_template_path, index=False)

    print("模板已生成，请在“{}”目录下查看。".format(directory))
    os.system("pause")
    sys.exit()


def auto_num():
    init()  # 初始化colorama
    # 使用提示
    print("# 自动编排考号")
    print("# 需准备\"考室安排表.xlsx\"及\"参考名单表.xlsx\"两个文件，与本程序.exe文件同目录。")
    print("# 考室安排表需要有三列：考室号、楼层、人数、科目组。混合考室不要使用合并单元格。")
    print("# 参考名单表包含基础三列：班级、姓名、科目组、分数；以及待生成五列：考号、考室号、座位号、楼层、教室。")
    print(
        Fore.RED + "#请确保文件名和扩展名与要求的一致；表格标题包含上面指出的标题，名称需一致，顺序随意；科目组不能为空。" + Style.RESET_ALL)
    # 模式选择
    mode = input("准备好后请使用数字键选择模式：\n1：考号编排模式；\n2：生成模板。\n")
    if mode == "1":
        mode_for_one()
    elif mode == "2":
        generate_template()