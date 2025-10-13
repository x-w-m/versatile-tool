#拆分规则详见readme文件
import itertools
import json
from collections import defaultdict

import pandas as pd


class DSU:
    # ... (并查集类代码保持不变) ...
    def __init__(self, elements):
        self.parent = {el: el for el in elements}
        self.rank = {el: 0 for el in elements}

    def find(self, i):
        if self.parent[i] == i:
            return i
        self.parent[i] = self.find(self.parent[i])
        return self.parent[i]

    def union(self, i, j):
        root_i = self.find(i)
        root_j = self.find(j)
        if root_i != root_j:
            if self.rank[root_i] > self.rank[root_j]:
                self.parent[root_j] = root_i
            else:
                self.parent[root_i] = root_j
                if self.rank[root_i] == self.rank[root_j]:
                    self.rank[root_j] += 1
            return True
        return False


def is_group_valid(group, separation_constraints):
    """检查一个组是否违反了“不能链接”的约束"""
    if len(group) < 2:
        return True
    for el1, el2 in itertools.combinations(group, 2):
        if frozenset({el1, el2}) in separation_constraints:
            return False
    return True


def split_lists_with_bidirectional_checks(input_dict: dict) -> dict:
    """
    最终版算法：同时处理 must-link 和 cannot-link 约束。
    """
    if not input_dict:
        return {}

    all_elements = set(el for lst in input_dict.values() for el in lst)
    dsu = DSU(all_elements)
    separation_constraints = set()

    sorted_items = sorted(input_dict.items(), key=lambda item: len(item[1]))

    result_dict = {}

    for key, current_list in sorted_items:
        # 简单情况
        if len(current_list) <= 1:
            result_dict[key] = [current_list]
            continue

        # 根据 DSU 将元素分组
        groups = defaultdict(list)
        for element in current_list:
            root = dsu.find(element)
            groups[root].append(element)
        atomic_groups = list(groups.values())

        # 如果已在同一组，无法拆分
        if len(atomic_groups) == 1:
            result_dict[key] = [current_list]
            continue

        # 寻找所有有效的拆分方案
        valid_splits = []
        for i in range(1, len(atomic_groups) // 2 + 1):
            for part_a_groups in itertools.combinations(atomic_groups, i):
                part_b_groups = [g for g in atomic_groups if g not in part_a_groups]

                sublist_a = [item for sublist in part_a_groups for item in sublist]
                sublist_b = [item for sublist in part_b_groups for item in sublist]

                # 关键：有效性检查
                if is_group_valid(sublist_a, separation_constraints) and \
                        is_group_valid(sublist_b, separation_constraints):
                    diff = abs(len(sublist_a) - len(sublist_b))
                    valid_splits.append((diff, (sorted(sublist_a), sorted(sublist_b))))

        # 决策与更新
        if not valid_splits:
            # 没有找到任何有效的拆分方式，意味着该列表必须保持整体
            result_dict[key] = [sorted(current_list)]
            # 更新 DSU，强制它们未来都在一起
            if len(current_list) > 1:
                first_el = current_list[0]
                for i in range(1, len(current_list)):
                    dsu.union(first_el, current_list[i])
        else:
            # 选择最平衡的有效拆分
            valid_splits.sort(key=lambda x: x[0])
            _, (final_a, final_b) = valid_splits[0]

            result_dict[key] = [final_a, final_b]

            # 更新 must-link 约束 (DSU)
            if len(final_a) > 1:
                first_el = final_a[0]
                for i in range(1, len(final_a)):
                    dsu.union(first_el, final_a[i])
            if len(final_b) > 1:
                first_el = final_b[0]
                for i in range(1, len(final_b)):
                    dsu.union(first_el, final_b[i])

            # 更新 cannot-link 约束
            for el_a in final_a:
                for el_b in final_b:
                    separation_constraints.add(frozenset({el_a, el_b}))

    return result_dict


def save_to_excel(input_dict: dict, result_dict: dict, filename: str = "split_results.xlsx"):
    """
    将输入和拆分结果保存到 Excel 文件中。

    Args:
        input_dict (dict): 原始的输入字典。
        result_dict (dict): 算法处理后的结果字典。
        filename (str): 要保存的 Excel 文件名。
    """
    # 准备一个列表，用于存储每一行的数据
    data_for_excel = []

    # 遍历结果字典，为每一条记录创建一行数据
    for key, split_lists in result_dict.items():
        row_data = {
            '原始键': key,
            '原始列表': str(input_dict[key])  # 将列表转换为字符串以便在单元格中显示
        }

        # 将拆分后的子列表放入不同的列中
        for i, sublist in enumerate(split_lists):
            row_data[f'子列表 {i + 1}'] = str(sublist)

        data_for_excel.append(row_data)

    # 使用 pandas 创建 DataFrame (可以理解为一个内存中的表格)
    df = pd.DataFrame(data_for_excel)

    # 为了美观，我们重新排列一下列的顺序
    # 获取所有列名
    columns = list(df.columns)
    # 确保'原始键'和'原始列表'在最前面
    if '原始键' in columns:
        columns.remove('原始键')
        columns.insert(0, '原始键')
    if '原始列表' in columns:
        columns.remove('原始列表')
        columns.insert(1, '原始列表')

    df = df[columns]

    # 将 DataFrame 写入 Excel 文件
    # index=False 表示不将 DataFrame 的行索引写入 Excel 文件
    try:
        df.to_excel(filename, index=False)
        print(f"\n✅ 结果已成功保存到文件: {filename}")
    except Exception as e:
        print(f"\n❌ 保存文件时出错: {e}")
        print("请确保您有权限在此位置创建文件，并且文件名是有效的。")

# 1. 定义您的输入数据字典
#    键可以是任何你喜欢的名字，比如 'A', 'B', 'C' 或者 'list_alpha'
#    值是包含元素的列表
my_input_data = {
    'C': [1, 2],
    'k':[2,4],
    'A': [1, 2, 3],
    'B': [1, 2, 4],
    'D': [6],
    'E': [3, 5],
    'F':[1,2,5],
    'G':[1,4,5],
    'H':[2,4,6,7],
    'i':[1,2,5,4],
    'j':[2,4,5,7]
}

# 2. 调用主函数，将您的数据作为参数传入
#    函数入口就是这里！
final_result = split_lists_with_bidirectional_checks(my_input_data)

# 4. 调用新函数，将输入和结果一起保存到 Excel 文件
#    您可以在这里修改想要保存的文件名
output_filename = "my_split_output.xlsx"




save_to_excel(my_input_data, final_result, filename=output_filename)