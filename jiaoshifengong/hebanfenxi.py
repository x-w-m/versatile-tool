#合班分析
import pandas as pd
from pprint import pprint


# --- 这是我们上一版的核心处理函数，保持不变 ---
def split_lists_traceable(input_data_dict):
    """
    对一个以字典形式存储的列表集合进行拆分，并提供可追溯的中文结果。
    (代码与上一版完全相同)
    """
    forbidden_pairs = set()
    for lst in input_data_dict.values():
        if len(lst) == 2:
            forbidden_pairs.add(tuple(sorted(lst)))

    results = {}
    for list_id, original_list in input_data_dict.items():
        list_len = len(original_list)

        if list_len == 2:
            results[list_id] = {'状态': '原始二元素列表', '原始列表': original_list}
            continue

        if list_len == 3:
            found_split = False
            for i in range(list_len):
                single_item = [original_list[i]]
                pair_item = [original_list[j] for j in range(list_len) if i != j]
                if tuple(sorted(pair_item)) not in forbidden_pairs:
                    results[list_id] = {'状态': '拆分成功', '原始列表': original_list,
                                        '拆分结果': [single_item, pair_item]}
                    found_split = True
                    break
            if not found_split:
                results[list_id] = {'状态': '无法拆分', '原始列表': original_list,
                                    '原因': '所有可能的拆分都会生成已存在的二元素列表'}
            continue

        if list_len == 4:
            found_split = False
            potential_splits = [
                ([original_list[0], original_list[1]], [original_list[2], original_list[3]]),
                ([original_list[0], original_list[2]], [original_list[1], original_list[3]]),
                ([original_list[0], original_list[3]], [original_list[1], original_list[2]])
            ]
            for pair1, pair2 in potential_splits:
                if tuple(sorted(pair1)) not in forbidden_pairs and tuple(sorted(pair2)) not in forbidden_pairs:
                    results[list_id] = {'状态': '拆分成功', '原始列表': original_list, '拆分结果': [pair1, pair2]}
                    found_split = True
                    break
            if not found_split:
                results[list_id] = {'状态': '无法拆分', '原始列表': original_list,
                                    '原因': '所有拆分方式都至少会生成一个已存在的二元素列表'}
            continue

        results[list_id] = {'状态': '无需处理', '原始列表': original_list,
                            '原因': f'列表长度为 {list_len}，不符合拆分规则 (2, 3, 或 4)。'}
    return results


# --- 新增：将结果保存到 Excel 的函数 ---
def save_results_to_excel(results_dict, filename="拆分结果.xlsx"):
    """
    将处理结果字典保存到 Excel 文件中。

    Args:
        results_dict (dict): 从 split_lists_traceable 函数获取的结果。
        filename (str): 要保存的 Excel 文件名。
    """
    # 1. 将结果字典转换为一个列表，每一项代表 Excel 的一行
    data_for_excel = []
    for key, details in results_dict.items():
        # 准备第三列的内容
        split_output = ""
        if details['状态'] == '拆分成功':
            # 将列表转换为字符串以便在单元格中显示
            split_output = str(details.get('拆分结果', ''))
        elif details['状态'] == '无法拆分' or details['状态'] == '无需处理':
            split_output = details.get('原因', '')
        elif details['状态'] == '原始二元素列表':
            split_output = '无需拆分'  # 或者直接写 '原始二元素列表'

        row = {
            '键': key,
            '原始列表': str(details['原始列表']),  # 同样转为字符串
            '拆分后的列表': split_output
        }
        data_for_excel.append(row)

    # 2. 使用 pandas 创建 DataFrame
    df = pd.DataFrame(data_for_excel)

    # 3. 将 DataFrame 保存到 Excel 文件
    # index=False 表示不将 DataFrame 的索引（0, 1, 2...）写入文件
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"\n结果已成功保存到文件: {filename}")


# --- 示例和测试 ---
if __name__ == "__main__":
    # input_data = {
    #     "阳芳艳": [2317, 2318], "孙艺松": [2319, 2320], "罗拾": [2314, 2316], "王祝梅": [2313, 2321],
    #     "廖海燕": [2322, 2325], "陈志昂": [2323, 2324], "孙鹏": [2307, 2309], "刘芳": [2326, 2327],
    #     "晏丽芬": [2306, 2311], "李立文": [2301, 2315], "黄梦云": [2302, 2308], "孙伟": [2304, 2305, 2312],
    #     "张勤": [2303, 2310], "邹宇谭": [2315, 2320], "阚康红": [2318, 2325], "邹陈娟": [2312, 2317],
    #     "陈瑛": [2311, 2323], "刘梦炎": [2307, 2319], "肖体金": [2306, 2308, 2313], "戴甲展": [2310, 2316],
    #     "刘红清": [2324, 2326, 2327, 2328], "谢湖元": [2309, 2322], "赵叶凌": [2314, 2321], "岳灵": [2306, 2314, 2316],
    #     "谭卓亚": [2318, 2319], "谢中兴": [2326, 2327, 2328], "易松柏": [2320, 2323, 2325],
    #     "雷湘峰": [2308, 2313, 2321], "陈准华": [2307, 2311, 2317, 2324], "蒋德奇": [2309, 2322],
    #     "周后芳": [2310, 2312, 2315], "钱金玉": [2310, 2313, 2314], "马红梅": [2326, 2327, 2328],
    #     "李春风": [2309, 2311, 2318], "马辉勇": [2319, 2320, 2321], "杨彩红": [2305, 2312, 2315],
    #     "阳春方": [2317, 2324, 2325], "王爱华": [2316, 2322, 2323], "文巧华": [2303, 2304], "邓丽菊": [2301, 2302, 2305]
    # }
    input_data = {
        "刘明彰": [2301], "谭冬": [2307, 2312], "刘辉杨": [2304], "钱慧": [2309, 2323], "朱凌雁": [2310, 2316],
        "肖立华": [2313, 2314], "陈桂芳": [2305, 2308], "易姬梅": [2318, 2325], "颜改河": [2315, 2317],
        "彭小雄": [2302, 2303], "刘晓玉": [2321, 2322], "肖慧晖": [2326, 2327, 2328], "陆茜茜": [2319, 2320, 2324],
        "李子晴": [2306, 2311], "周国慧": [2305, 2313], "肖俐": [2326, 2327, 2328], "刘玲": [2301, 2317],
        "陈美凤": [2321], "陈建隆": [2319, 2325], "陈鹏军": [2318], "邹颖": [2302, 2314], "欧阳爱娥": [2315, 2316],
        "阮巧媛": [2312, 2320], "谢赛君": [2309, 2323], "陈晓红": [2303, 2324], "罗玉南": [2306, 2310],
        "黄玲": [2307, 2311], "谭海": [2304, 2322], "范海花": [2308]
    }

    print("--- 正在处理输入数据... ---")
    # 1. 运行拆分算法
    final_results = split_lists_traceable(input_data)

    print("\n--- 算法处理结果预览 ---")
    pprint(final_results, sort_dicts=False)

    # 2. 将结果保存到 Excel
    excel_filename = "拆分结果.xlsx"
    save_results_to_excel(final_results, excel_filename)
