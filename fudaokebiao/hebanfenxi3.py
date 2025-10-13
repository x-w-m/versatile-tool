import itertools
from collections import defaultdict
import copy  # 用于深拷贝状态
import pandas as pd  # 导入 pandas 用于读取 Excel
from numpy import dtype

from fudaokebiao.hebanfenxi2 import save_to_excel


# ==============================================================================
# DSU 和 is_group_valid 函数保持不变
# ==============================================================================
class DSU:
    def __init__(self, elements):
        self.parent = {el: el for el in elements}
        self.rank = {el: 0 for el in elements}

    def find(self, i):
        if self.parent[i] == i: return i
        self.parent[i] = self.find(self.parent[i])
        return self.parent[i]

    def union(self, i, j):
        root_i, root_j = self.find(i), self.find(j)
        if root_i != root_j:
            if self.rank[root_i] > self.rank[root_j]:
                self.parent[root_j] = root_i
            else:
                self.parent[root_i] = root_j
                if self.rank[root_i] == self.rank[root_j]: self.rank[root_j] += 1
            return True
        return False

    def copy(self):
        """创建一个DSU对象的副本"""
        new_dsu = DSU([])
        new_dsu.parent = self.parent.copy()
        new_dsu.rank = self.rank.copy()
        return new_dsu


def is_group_valid(group, separation_constraints):
    if len(group) < 2: return True
    for el1, el2 in itertools.combinations(group, 2):
        if frozenset({el1, el2}) in separation_constraints: return False
    return True


# ==============================================================================
# 新增：从 Excel 读取数据的函数
# ==============================================================================
def read_data_from_excel(file_path, sheet_name):
    """
    从 Excel 文件读取教师和班级数据，并转换为字典格式。

    Args:
        file_path (str): Excel 文件的路径。
        sheet_name (str): 要读取的工作表的名称。

    Returns:
        dict: 老师为键，班级列表为值的字典。
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        # 确保 '教师' 和 '班级' 列存在
        if '教师' not in df.columns or '班级' not in df.columns:
            raise ValueError(f"工作表 '{sheet_name}' 中缺少 '教师' 或 '班级' 列。")

        teacher_classes = defaultdict(list)
        for index, row in df.iterrows():
            teacher = row['教师']
            class_str = row['班级']

            # 确保老师和班级信息不为空
            if pd.notna(teacher) and pd.notna(class_str):
                class_list = [int(x) for x in class_str.split(',')]
                teacher_classes[teacher] = class_list
        return dict(teacher_classes)
    except FileNotFoundError:
        print(f"错误：找不到文件 {file_path}")
        return None
    except Exception as e:
        print(f"读取 Excel 文件时发生错误: {e}")
        return None


# ==============================================================================
# 核心优化类：OptimalSplitter
# ==============================================================================
class OptimalSplitter:
    def __init__(self, input_dict, eval_group1=None, eval_group2=None):
        """
        初始化时接收额外的评估列表。

        Args:
            input_dict (dict): 待拆分的字典。
            eval_group1 (list, optional): 第一个评估群组。Defaults to None.
            eval_group2 (list, optional): 第二个评估群组。Defaults to None.
        """
        self.input_dict = input_dict
        self.sorted_items = sorted(input_dict.items(), key=lambda item: len(item[1]))
        self.all_elements = set(el for lst in input_dict.values() for el in lst)

        # 为了快速查找，将评估列表转换为集合
        self.eval_group1 = set(eval_group1) if eval_group1 else set()
        self.eval_group2 = set(eval_group2) if eval_group2 else set()

        self.best_solution = None
        self.min_cost = float('inf')

    def _calculate_cost(self, solution):
        """
        计算一个完整解决方案的总成本，包含不平衡度和群组分裂惩罚。
        """
        # 1. 计算不平衡度成本
        balancing_cost = 0
        for key, splits in solution.items():
            # 只对被拆分成两部分的列表计算不平衡度
            if len(splits) == 2:
                cost = (len(splits[0]) - len(splits[1])) ** 2
                balancing_cost += cost

        # 2. 计算群组分裂成本
        splitting_cost = 0
        # 遍历解决方案中的每一个最终子列表
        all_sublists = [sublist for splits in solution.values() for sublist in splits]

        for sublist in all_sublists:
            # 只关心包含多个元素的子列表
            if len(sublist) < 2:
                continue

            sublist_set = set(sublist)
            # 检查子列表是否完全被任一评估群组包含
            is_pure_group1 = sublist_set.issubset(self.eval_group1)
            is_pure_group2 = sublist_set.issubset(self.eval_group2)

            # 如果一个子列表同时包含两个评估群组的成员，则施加惩罚
            if is_pure_group1 and is_pure_group2:
                splitting_cost += 8

        return balancing_cost + splitting_cost

    def find_optimal_split(self):
        """启动回溯搜索寻找最优解"""
        initial_dsu = DSU(self.all_elements)
        initial_constraints = set()

        self._backtrack_search(0, initial_dsu, initial_constraints, {})

        return self.best_solution

    def _backtrack_search(self, item_index, dsu, constraints, current_solution):
        """
        递归的回溯搜索函数

        Args:
            item_index (int): 当前正在处理的 sorted_items 的索引。
            dsu (DSU): 当前路径的 DSU 状态。
            constraints (set): 当前路径的 separation_constraints 状态。
            current_solution (dict): 到目前为止构建的部分解决方案。
        """
        # --- Base Case: 所有列表都处理完毕 ---
        if item_index == len(self.sorted_items):
            cost = self._calculate_cost(current_solution)
            if cost < self.min_cost:
                self.min_cost = cost
                self.best_solution = current_solution
                print(f"找到一个新优解，成本: {cost}, 解: {current_solution}")  # 调试信息
            return

        # --- Recursive Step ---
        key, current_list = self.sorted_items[item_index]

        # 简单情况或已完全约束的情况
        if len(current_list) <= 1:
            current_solution[key] = [current_list]
            self._backtrack_search(item_index + 1, dsu, constraints, current_solution)
            return

        groups = defaultdict(list)
        for element in current_list:
            groups[dsu.find(element)].append(element)
        atomic_groups = list(groups.values())

        if len(atomic_groups) == 1:
            current_solution[key] = [sorted(current_list)]
            self._backtrack_search(item_index + 1, dsu, constraints, current_solution)
            return

        # 寻找所有有效的拆分方案
        valid_splits = []
        for i in range(1, len(atomic_groups) // 2 + 1):
            for part_a_groups in itertools.combinations(atomic_groups, i):
                part_b_groups = [g for g in atomic_groups if g not in part_a_groups]
                sublist_a = [item for sublist in part_a_groups for item in sublist]
                sublist_b = [item for sublist in part_b_groups for item in sublist]
                if is_group_valid(sublist_a, constraints) and is_group_valid(sublist_b, constraints):
                    valid_splits.append((sorted(sublist_a), sorted(sublist_b)))

        # --- 决策与递归 ---
        if not valid_splits:
            # 没有有效拆分，只能保持整体
            current_solution[key] = [sorted(current_list)]
            if len(current_list) > 1:
                first_el = current_list[0]
                for i in range(1, len(current_list)): dsu.union(first_el, current_list[i])
            self._backtrack_search(item_index + 1, dsu, constraints, current_solution)
        else:
            # 遍历所有有效的拆分选择（决策树的分支）
            for split_choice in valid_splits:
                final_a, final_b = split_choice

                # 为这个新的分支创建状态的深拷贝
                next_dsu = dsu.copy()
                next_constraints = constraints.copy()
                next_solution = current_solution.copy()

                # 应用当前选择的拆分，并更新新状态
                next_solution[key] = [final_a, final_b]
                if len(final_a) > 1:
                    for i in range(1, len(final_a)): next_dsu.union(final_a[0], final_a[i])
                if len(final_b) > 1:
                    for i in range(1, len(final_b)): next_dsu.union(final_b[0], final_b[i])
                for el_a in final_a:
                    for el_b in final_b: next_constraints.add(frozenset({el_a, el_b}))

                # 递归进入下一层
                self._backtrack_search(item_index + 1, next_dsu, next_constraints, next_solution)


# ==============================================================================
# 主程序入口
# ==============================================================================
if __name__ == "__main__":
    # --- 从 Excel 文件读取数据 ---
    excel_file_path = '教师信息高二.xlsx'  # <--- 请用户替换为他们的 Excel 文件名
    sheet_name="化学政治"

    # 假设 Excel 文件中有名为 '物理', '化政', '生地' 的工作表
    data_to_process = read_data_from_excel(excel_file_path, sheet_name)
    output_filename = f"{sheet_name}.xlsx"

    # 如果读取失败，则 data_to_process 为 None，程序应能处理这种情况或退出
    if data_to_process is None:
        print(f"无法从 Excel 文件 '{excel_file_path}' 中读取数据，程序将退出。")
        exit()

    # 群组1: 探索楼
    evaluation_group_1 = [2407, 2408, 2409, 2410, 2411, 2412, 2413, 2414, 2415, 2416, 2417, 2418, 2419, 2420, 2421,
                          2422, 2423, 2424, 2428, 2429, 2430, 2431, 2432, 2433]
    # 群组2: 求知楼
    evaluation_group_2 = [2402, 2403, 2404, 2405, 2406]

    print("输入数据:", data_to_process)
    print("\n--- 开始寻找最优拆分方案 ---\n")

    # 1. 创建 splitter 实例
    splitter = OptimalSplitter(data_to_process, evaluation_group_1, evaluation_group_2)

    # 2. 运行优化算法
    optimal_result = splitter.find_optimal_split()

    # 3. 打印最终的最优结果
    print("\n--- 搜索完成 ---")
    if optimal_result:
        # 注意：这里的 save_to_excel 函数的第一个参数是原始数据
        save_to_excel(data_to_process, optimal_result, filename=output_filename)
    else:
        print("未能找到任何有效的拆分方案。")
