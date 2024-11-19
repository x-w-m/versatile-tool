# 读取分工，并分析是否跨楼
from collections import Counter, defaultdict
import pandas as pd


class Teacher:
    def __init__(self, name, kemu_tpye):
        self.name = name
        self.kemu_tpye = kemu_tpye

    # 重写__eq__，根据姓名和科目判断是否是同一个老师
    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.name == other.name and self.kemu_tpye == other.kemu_tpye
        return False

    def __hash__(self):
        return hash((self.name, self.kemu_tpye))


# 任课，包含科目、和班级
class RenKe:
    def __init__(self, banji, kemu_tpye, teacher):
        self.banji = banji
        self.kemu_tpye = kemu_tpye
        self.teacher = teacher
        self.position = None


# 班级，包含班级名、教学楼。并预留两节课的位置
class BanJi:
    def __init__(self, name, jiaoxuelou):
        self.name = name
        self.jiaoxuelou = jiaoxuelou
        self.w1 = None
        self.w2 = None

    # 将
    def place_kechen(self, renke, position):
        if position == "w1" and self.w1 is None:
            self.w1 = renke
            renke.position = "w1"
            return True
        elif position == "w2" and self.w2 is None:
            self.w2 = renke
            renke.position = "w2"
            return True
        return False

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.name == other.name and self.jiaoxuelou == other.jiaoxuelou
        return False

    def __hash__(self):
        return hash((self.name, self.jiaoxuelou))


# 教师分工,包含教师姓名、科目、任教表
class TeacherFengong:
    def __init__(self, teacher, kemu_tpye, renke_list):
        self.teacher = teacher
        self.kemu_tpye = kemu_tpye
        # 合班标记
        self.heban = False
        # 跨楼标记,默认不跨楼
        self.kualou = False
        # 放最后，避免标记修改不生效
        self.renke_lists = self.split_fengong(renke_list)

    # 切割任教，多个班级合班上课
    def split_fengong(self, renke_list):
        if len(renke_list) == 1:
            renke_lists = [renke_list]
            return renke_lists
        if len(renke_list) == 2:
            renke_lists = [renke_list[:1], renke_list[1:]]
            return renke_lists
        # 如果任课班级大于1，则需要进行分组
        if len(renke_list) > 2:
            # 合班上课
            self.heban = True
            # 提取任教班级的班级
            banji_list = [renke.banji for renke in renke_list]
            # 统计各教学楼出现的次数
            jiaoxuelou_countis = Counter(banji.jiaoxuelou for banji in banji_list)
            # 判断有几个教学楼
            if len(jiaoxuelou_countis) > 1:
                # 跨楼
                self.kualou = True
                # 主要教学楼,结果是个元组('教学楼',2)
                main_jiaoxuelou_tuple = max(jiaoxuelou_countis.items(), key=lambda x: x[1])
                main_jiaoxuelou = main_jiaoxuelou_tuple[0]
                # 获取主要教学楼的任课表
                main_renke_list = [renke for renke in renke_list if renke.banji.jiaoxuelou == main_jiaoxuelou]
                # 次要教学楼的任课表
                minor_renke_list = [renke for renke in renke_list if renke.banji.jiaoxuelou != main_jiaoxuelou]
                # 如果次要教学楼中只有一个任课班级，且主教学楼比次要楼多1个班以上，则分一个班到次要教学楼组。
                if len(minor_renke_list) <= 2 and len(main_renke_list) - len(minor_renke_list) >= 2:
                    renke_temp = main_renke_list.pop(0)
                    minor_renke_list.append(renke_temp)
                renke_lists = [main_renke_list, minor_renke_list]
                return renke_lists
            else:
                # 中间元素下标
                mid_index = len(banji_list) // 2
                renke_lists = [renke_list[:mid_index], renke_list[mid_index:]]
                return renke_lists


# 获取班级列表；获取各科任课表，保存到字典中；
def process_ball_groups_from_excel(file_path):
    # 读取 Excel 文件中的数据
    df = pd.read_excel(file_path)

    # 获取班级
    banji_name_list = df["班级"].astype(str).tolist()
    banji_jxl_list = df["教学楼"].astype(str).tolist()
    banji_list = [BanJi(name, jxl) for name, jxl in zip(banji_name_list, banji_jxl_list)]
    # 提取语文教师教师和外语教师教师的编号,初始化键的值为列表
    renke_dict = defaultdict(list)

    kemu_list = ["语文", "数学", "外语", "物理", "化学", "生物", "政治", "历史", "地理"]
    for kemu in kemu_list:
        # 遍历每一行数据
        for _, row in df.iterrows():
            banji_name = row['班级']
            teacher_name = row[kemu]
            jxl = row['教学楼']
            # 将当前班级的编号加入到对应的语文教师中
            if pd.notna(teacher_name):
                banji = BanJi(banji_name, jxl)
                teacher = Teacher(teacher_name, kemu)
                renke = RenKe(banji, kemu, teacher)
                renke_dict[kemu].append(renke)

    return banji_list, renke_dict


# 对教师任课进行分组，并获取所有分组情况
def group_teacher_renke(renke_dict):
    teacher_fengong_list = []
    # 遍历每一个科目和它的所有任课情况
    for kemu, all_reke_list in renke_dict.items():
        # 获取此科目的任课教师列表,使用dict.fromkeys去重并保留顺序,前提是需要重写类Teacher的__hash__和__eq__方法。
        teacher_list = list(dict.fromkeys(renke.teacher for renke in all_reke_list))

        for teacher in teacher_list:
            # 获取此教师的任课列表
            reke_list = [renke for renke in all_reke_list if renke.teacher.name == teacher.name]
            # 教师分工分组
            teacher_fengong = TeacherFengong(teacher, kemu, reke_list)
            teacher_fengong_list.append(teacher_fengong)
    return teacher_fengong_list


# 示例使用
file_path = '分工.xlsx'
banji_list, renke_dict = process_ball_groups_from_excel(file_path)

teacher_fengong_list = group_teacher_renke(renke_dict)
# for teacher_fengong in teacher_fengong_list:
#     for i in teacher_fengong.renke_lists:
#         for j in i:
#             print(j.teacher.name, j.banji.name, j.kemu_tpye, end="\t")
#         print(" ", )


kumu1 = "数学"
kumu2 = "外语"

# 提取语文外语教师分工,结果是分工对象列表
kemu1_teacher_fg_list = [teacher_fengong for teacher_fengong in teacher_fengong_list if
                         teacher_fengong.kemu_tpye == kumu1]
kemu2_teacher_fg_list = [teacher_fengong for teacher_fengong in teacher_fengong_list if
                         teacher_fengong.kemu_tpye == kumu2]


# 检查当前合班上课分工是否会和科目2的分工产生冲突。
# 此方法设计标准是合班分工对比所有只有两个班的分工。不和其他合班分工比较
def jianchachongtu(kemu1_teacher_fg, kemu2_teacher_fg_list):
    if kemu1_teacher_fg.kualou:
        print("跨楼")
    # 遍历当前分工任课表
    for i, renke_list in enumerate(kemu1_teacher_fg.renke_lists):
        # 判断当前任课组班级数，如果为1则跳过
        if len(renke_list) == 1:
            continue
        else:
            # 遍历科目2的分工表，逐个检查每个分工与待检查分工的冲突情况
            for kemu2_teacher_fg in kemu2_teacher_fg_list:
                # 科目1分工小组包含的班级列表
                banji_list1 = [renke.banji for renke in renke_list]

                # 获取科目2当前分工任课表
                kemu2_teacher_renke_lists = kemu2_teacher_fg.renke_lists
                # 只有一个班跳过检查
                if len(kemu2_teacher_renke_lists) == 1:
                    continue
                # 提取分工表中的任课信息，获取科目2的任课班级
                banji_list2 = [renke.banji for renke_list in kemu2_teacher_renke_lists for renke in renke_list]

                # 判断两个班是否被合班到一起，并记录该班在数组中的位置
                pos1, pos2 = None, None
                if banji_list2[0] in banji_list1:
                    pos1 = (i, banji_list1.index(banji_list2[0]))
                if banji_list2[1] in banji_list1:
                    pos2 = (i, banji_list1.index(banji_list2[1]))

                # 判断是否存在冲突
                if pos1 and pos2 and pos1[0] == pos2[0]:
                    print("存在冲突！")
                    # 交换第二个班到另一个小组
                    other_row = 1 - pos2[0]
                    renke_lists = kemu1_teacher_fg.renke_lists
                    # 使用多重赋值从两个子列表中交换元素
                    # 只会固定交换另一个子列表的第一个元素，如果交换后依旧有冲突可能会进入死循环。
                    renke_lists[pos1[0]][pos1[1]], renke_lists[other_row][0] = renke_lists[other_row][0], \
                        renke_lists[pos1[0]][pos1[1]]
                    print("完成班级分组交换。")
                    # 递归检查，直到没有冲突
                    jianchachongtu(kemu1_teacher_fg, kemu2_teacher_fg_list)
    return True


def chongtuqingkuapanduan(kemu1_teacher_fg_list, kemu2_teacher_fg_list):
    # 分别提取两个科目中有合班情况的分工，结果是分工对象列表
    kemu1_heban_teacher_fg_list = [teacher_fengong for teacher_fengong in kemu1_teacher_fg_list if
                                   teacher_fengong.heban]
    kemu2_heban_teacher_fg_list = [teacher_fengong for teacher_fengong in kemu2_teacher_fg_list if
                                   teacher_fengong.heban]
    l1 = len(kemu1_heban_teacher_fg_list)
    l2 = len(kemu2_heban_teacher_fg_list)
    # 两个科目均不存在合班情况
    if l1 == 0 and l2 == 0:
        return "00"
    if l1 != 0 and l2 == 0:
        # 遍历有合班上课的分工，每个元素都是分工对象
        for kemu1_teacher_fg in kemu1_heban_teacher_fg_list:
            jianchachongtu(kemu1_teacher_fg, kemu2_teacher_fg_list)

        return "10"
    if l1 == 0 and l2 != 0:
        # 与前一种情况相反，传参时交换科目1与科目2
        for kemu2_teacher_fg in kemu2_heban_teacher_fg_list:
            jianchachongtu(kemu2_teacher_fg, kemu1_teacher_fg_list)
        return "01"
    # 两个科目均有合班情况
    if l1 != 0 and l2 != 0:
        # 提取两种科目教师没有合班情况的分工
        kemu1_buheban_teacher_fg_list = [teacher_fengong for teacher_fengong in kemu1_teacher_fg_list if
                                         not teacher_fengong.heban]
        kemu2_buheban_teacher_fg_list = [teacher_fengong for teacher_fengong in kemu2_teacher_fg_list if
                                         not teacher_fengong.heban]
        
        # 分别检查该科目合班时与另一个科目不合班是否存在冲突
        for kemu1_teacher_fg in kemu1_heban_teacher_fg_list:
            jianchachongtu(kemu1_teacher_fg, kemu2_buheban_teacher_fg_list)
        for kemu2_teacher_fg in kemu2_heban_teacher_fg_list:
            jianchachongtu(kemu2_teacher_fg, kemu1_buheban_teacher_fg_list)
        #检查两个科目合班分工间是否有冲突

        #
        return "11"


status_code = chongtuqingkuapanduan(kemu1_teacher_fg_list, kemu2_teacher_fg_list)
print(status_code)
