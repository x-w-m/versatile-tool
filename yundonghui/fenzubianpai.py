# 运动会分组编排
import pandas as pd

# 设置控制台的宽度和每列数据的最大宽度，None表示不限制
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
# 读取 Excel 文件
df = pd.read_excel("参赛名单.xlsx", dtype=str)

# 随机打乱数据
df_shuffled = df.sample(frac=1).reset_index(drop=True)
# 分组算法
groups = []
for _, row in df_shuffled.iterrows():
    # 查找可以加入的组
    eligible_groups = [group for group in groups if len(group) < 8 and row['班级'] not in [r['班级'] for r in group]]

    if eligible_groups:
        # 选择人数最少的组
        min_group = min(eligible_groups, key=len)
        min_group.append(row)
    else:
        # 创建一个新组
        groups.append([row])

# 最后的平衡调整
while True:
    group_lengths = [len(group) for group in groups]
    min_length, max_length = min(group_lengths), max(group_lengths)

    if max_length - min_length > 1:
        # 找到人数最多和最少的组
        max_group_index = group_lengths.index(max_length)
        min_group_index = group_lengths.index(min_length)
        max_group = groups[max_group_index]
        min_group = groups[min_group_index]

        # 创建一个包含最小组所有班级的集合
        min_group_classes = {member['班级'] for member in min_group}

        # 需要移动的成员索引
        member_to_move_index = None

        # 从最大组中找到可以移动的成员
        for index, member in enumerate(max_group):
            if member['班级'] not in min_group_classes:
                member_to_move_index = index
                # 不能直接将member作为参数传递给list对象的remove方法，要从人多的组移出一位选手建议使用索引。
                # 因为member是一个pandas行对象（Series对象），不能直接用于标准的Python比较操作。
                # 有概率会出现错误：The truth value of a Series is ambiguous.Use a.empty, a.bool(), a.item(), a.any() or a.all().
                # max_group.remove(member)
                break

        # 如果找到了合适的成员，则进行移动
        if member_to_move_index is not None:
            # 故使用索引来移出符合要求的元素
            member_to_move = max_group.pop(member_to_move_index)
            min_group.append(member_to_move)
    else:
        break

# 根据组内人数排序组别（人数多的组优先）
groups.sort(key=len, reverse=True)

# 创建一个新的 DataFrame 来保存分组结果
grouped_df = pd.DataFrame([item for group in groups for item in group])

# 添加组号
grouped_df['组号'] = [i + 1 for i, group in enumerate(groups) for _ in group]

# 保存为 Excel 文件
grouped_df.to_excel("参赛分组.xlsx", index=False)
