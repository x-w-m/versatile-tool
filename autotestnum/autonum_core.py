# 自动编排考号
# 需准备考室安排表及参考名单表两个表。
# 考室安排表需要有三列：考室号、教室、楼层、人数、科目组。
# 参考名单表包含基础三列：班级、姓名、科目组、分数；以及待生成四列：考号、考室号、座位号、楼层、教室。
import sys

# 导入所需的库
import pandas as pd
from colorama import init, Fore, Style


def mode_for_one():
    print("正在进行考号编排......")
    # 读取两个excel文件中的数据，存储为pandas的DataFrame对象
    df_room = pd.read_excel("考室安排表.xlsx")  # 考室安排表
    df_student = pd.read_excel("参考名单表.xlsx", dtype=str)  # 参考名单表

    # 将分数转换为float
    df_student["分数"] = df_student["分数"].astype(float)
    # 按照科目组和分数对学生进行排序,使用稳定的归并排序，不改变相同值的相对顺序
    df_student.sort_values(by=["科目组", "分数"], ascending=[True, False], kind="mergesort", inplace=True)
    kshp = input("请输入考号前缀，输入0表示使用编号作为考号：")
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
            room_position = room_row["楼层"]
            room = room_row["教室"]
            room_size = room_row["人数"]

            # 如果考室号在tracker中没有记录，说明是新的考室或者新的科目组的开始，初始化座位号为0，下一步会加1
            if room_number not in room_seat_tracker:
                room_seat_tracker[room_number] = 0

            # 从之前的座位号继续开始分配
            start_seat_number = room_seat_tracker[room_number] + 1

            # 筛选出该考室需要安排的学生
            df_student_room = df_student_group.iloc[count:count + room_size]
            # 给每个学生分配座位号，并生成考号
            seat_number = start_seat_number  # 设置本次循环开始时的座位号。
            for i, r in df_student_room.iterrows():
                # 添加原始索引，用于保留原来顺序
                r["原始索引"] = i
                # 在迭代过程中，r是一个Series对象，它是从df_student_room中复制出来的一行数据，
                # 对r所做的修改不会反映到原始的DataFrame对象df_student_room中。
                # 想要修改原始的DataFrame对象，可以使用.loc[]方法来定位行和列，并直接修改DataFrame对象中的值。
                # 或者将被修改的行“r”保存到列表中，最后将其转换为新的DataFrame对象。
                temp_ksh = str(room_number).zfill(2)  # 考室号用两位数表示，不足补零
                temp_zwh = str(seat_number).zfill(2)  # 座位号

                # 检测是否需要自定义考号，如果输入“0”则表示不需要自定义考号，使用默认的编号作为考号。
                if kshp == "0":
                    r["考号"] = r["编号"]
                else:
                    r["考号"] = kshp + temp_ksh + temp_zwh

                r["考室号"] = temp_ksh
                r["座位号"] = temp_zwh
                r["楼层"] = room_position
                r["教室"] = room

                seat_number += 1
                # 将编排后的行添加到列表中，rows是Series对象的列表，每个Series带有索引。
                rows.append(r)

                # 使用.loc[]方法来定位行和列，从而直接修改DataFrame对象中的值。
                # 此方法会出现SettingWithCopyWarning: A value is trying to be set on a copy of a slice from a DataFrame.
                # 原因是df_student_room是由切片操作获取的，可能是一个副本而不是视图，对它进行的修改可能不会影响原始DataFrame
                # 为了消除不确定性，应该避免使用这种方式。
                # 座位号用两位数表示，不足补零
                # df_student_room.loc[i, "座位号"] = str(seat_number).zfill(2)
                # df_student_room.loc[i, "考室号"] = str(room_number).zfill(2)
                # # 考号由“1000”、考室号、座位号拼接而成
                # df_student_room.loc[i, "考号"] = "1000" + str(room_number).zfill(2) + str(seat_number).zfill(2)

            # 将该考室安排好的学生添加到结果中
            # df_result = pd.concat([df_result, df_student_room])
            # 更新该考室的最后一个座位号
            room_seat_tracker[room_number] = seat_number - 1
            # 更新已经安排了多少个学生的变量
            count += room_size

    # 将rows列表转换为DataFrame对象，用于存储最终的结果。columns参数会匹配Series索引，按columns中的顺序调整列的顺序。
    df_result = pd.DataFrame(rows,
                             columns=["原始索引", "班级", "姓名", "科目组", "分数", "考号", "考室号", "座位号", "楼层",
                                      "教室"],
                             dtype=str)
    print("数据编排结束。\n正在处理特殊科目组。。。")
    # 将“科目组”列中的数据进行转换操纵，去除内容中的数字。
    df_result["科目组"] = df_result["科目组"].str.replace(r"\d+", "", regex=True)
    print(Fore.RED + "已去除特殊科目组中的数字。" + Style.RESET_ALL)

    df_result['原始索引'] = df_result['原始索引'].astype(int)
    df_result.sort_values(by=["原始索引"], inplace=True)
    df_result.to_excel("编排结果/编排结果_保留顺序.xlsx", index=False)
    df_result.drop(columns=["原始索引"], inplace=True)

    # 创建一个ExcelWriter对象，使用with语句管理上下文，无需手动调用close方法关闭文件。
    with pd.ExcelWriter("考生去向表/考生去向表.xlsx") as writer:
        # 将结果按照班级、考号进行排序，并重置索引
        # inplace=True意味着排序操作将直接在原始的 df_result 对象上进行，默认为False，会返回新的DataFrame
        df_result.sort_values(by=["班级", "考室号", "座位号"], inplace=True)
        df_result.reset_index(drop=True, inplace=True)
        # 保存“考生去向表”为第一个工作表
        df_result.to_excel(writer, sheet_name="考生去向表", index=False)

        # 对df_result根据考号进行排序，为创建第二个工作表做准备
        df_seating = df_result.sort_values(by=["考室号", "座位号"])
        df_seating.reset_index(drop=True, inplace=True)

        # 保存排序后的数据为第二个工作表，这里命名为“考室座次表”
        df_seating.to_excel(writer, sheet_name="考室座次表", index=False)

    print("数据已保存至“考生去向表”目录下'考生去向表.xlsx'，包含'考生去向表'和'考室座次表'两个工作表。")


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
    print("# 自动编排考号控制台版本 V1.2")
    print("# 需准备\"考室安排表.xlsx\"及\"参考名单表.xlsx\"两个文件，与本程序.exe文件同目录。")
    print("# 考室安排表需要有三列：考室号、楼层、人数、科目组。混合考室每个科目组占一行。")
    print("# 参考名单表包含基础三列：班级、姓名、科目组、分数；以及待生成五列：考号、考室号、座位号、楼层、教室。")
    print(
        "# 对考室顺序和考生顺序没要求，即不需要提前排序。若要指定班级排到指定考室请同时对班级和考室的科目组进行重命名，如“物化生1”。")
    print(
        "# 对于相同科目组的考室，可以通过将指定考室放在前面的方式来实现优先编排该考室。如尖子考室放前面可以将成绩靠前的考生排入该考室。")
    print(
        Fore.RED + "# 请确保文件名和扩展名与要求的一致；表格标题包含上面指出的标题，名称需一致，顺序随意；科目组不能为空。" + Style.RESET_ALL)
    # 模式选择
    mode = input("准备好后请使用数字键选择模式：\n1：考号编排模式；\n2：生成模板。\n")
    if mode == "1":
        mode_for_one()
    elif mode == "2":
        generate_template()
