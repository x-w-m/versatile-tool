"""主程序"""
import os

from autotestnum import autonum_core, chaifen
from autotestnum import formatexcel
from autotestnum import paging

if __name__ == '__main__':
    # 指定要检查的目录
    directories = ["编排结果", "考生去向表"]

    # 遍历目录列表
    for directory in directories:
        # 检查目录是否存在
        if not os.path.exists(directory):
            # 如果目录不存在，则创建它
            os.makedirs(directory)

    # 考号编排
    autonum_core.auto_num()
    # 调整去向表格式
    print("正在进行格式调整......")
    formatexcel.excel_format("考生去向表/考生去向表.xlsx")
    # 拆分表格,并调整格式,设置年级
    print("正在进行表格拆分......")
    chaifen.split_excel("高三")
    formatexcel.excel_format("考生去向表/考生去向表_拆分.xlsx")
    formatexcel.excel_format("考生去向表/考室座次表_拆分.xlsx")
    # 设置打印格式
    print("正在设置打印格式......")
    # 设置去向表汇总表打印格式
    paging.page_format()
    # 设置拆分表格打印格式
    # 输入考试名称
    exam_name = input("请输入考试名称：")
    # exam_name = "2024年上学期高一期中"
    paging.page_format_cf("考生去向表/考生去向表_拆分.xlsx", f"{exam_name}考生去向表")
    paging.page_format_cf("考生去向表/考室座次表_拆分.xlsx", f"{exam_name}考室座次表")
    input("考号编排完成，按 Enter 键退出...")
