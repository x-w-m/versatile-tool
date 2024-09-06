# 获取题库和题目参数
import base64
import json
import re
import time

import ddddocr
from selenium import webdriver

from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By


def login(driver, username, password):
    # 打开目标网站
    driver.get(
        "http://hn.12348.gov.cn/ucenter/#/login?redirect=http%3A%2F%2Fhn.12348.gov.cn%2Ffxmain%2Fstudy%3Fbiz%3Dfxmain")  # 替换为你要登录的网站URL
    # 等待页面加载
    time.sleep(5)  # 你可以改为WebDriverWait以便更精确的控制等待时间
    # 定位用户名输入框并输入用户名
    username_field = driver.find_element(By.XPATH, "//input[@placeholder='请输入您的账号']")
    username_field.clear()
    username_field.send_keys(username)
    # 定位密码输入框并输入密码
    password_field = driver.find_element(By.XPATH, "//input[@placeholder='请输入您的密码']")
    password_field.clear()
    password_field.send_keys(password)
    # 定位验证码图片元素并获取其src属性
    captcha_element = driver.find_element(By.XPATH, "//img[@class='el-image__inner']")
    captcha_base64 = captcha_element.get_attribute("src").split(",")[1]  # 提取base64编码部分
    # 将base64编码解码为二进制数据
    captcha_bytes = base64.b64decode(captcha_base64)
    # 初始化识别器
    ocr = ddddocr.DdddOcr()
    # 识别验证码
    captcha_text = ocr.classification(captcha_bytes)
    # 定位用户名输入框并输入用户名
    username_field = driver.find_element(By.XPATH, "//input[@placeholder='请输入验证码']")
    username_field.clear()
    username_field.send_keys(captcha_text)
    # 提交表单
    driver.find_element(By.XPATH, "//button[@class='el-button el-button--default']").click()
    # 等待页面跳转
    time.sleep(5)
    return driver


# 阅读读本，获取题目信息及参数
def read_book(driver, answers):
    # 声明为全局变量，字符串修改在方法之外生效
    global tiku
    # 题目参数
    global timucanshu

    # 获取读本id
    caseid_element = driver.find_element(By.ID, "caseid")
    caseid = caseid_element.get_attribute("value")

    # 记录答题个数
    j = 0
    # 获取所有的 <li> 元素
    li_elements = driver.find_elements(By.XPATH, "//div[@id='tree']//li")
    # 正则表达式匹配模式，用于检测形如 (x/y) 的文本
    progress_pattern = re.compile(r"\(\d+/\d+\)")
    # 遍历所有的 <li> 元素
    for i_li in range(len(li_elements)):
        # 重新获取元素，避免元素失效
        li_elements = driver.find_elements(By.XPATH, "//div[@id='tree']//li")
        li = li_elements[i_li]
        # 获取 <a> 标签
        a_tag = li.find_element(By.TAG_NAME, "a")

        # 获取链接文本
        link_text = a_tag.text

        # 判断文本中是否包含 (x/y) 形式的进度标识
        if progress_pattern.search(link_text):
            print(f"文章标题: {link_text}")

            # 点击链接
            a_tag.click()
            time.sleep(3)

            # 获取章节id
            chatpId_element = driver.find_element(By.CLASS_NAME, "chaptId")
            chatpId = chatpId_element.get_attribute("value")

            # 获取题目列表
            question_element = driver.find_element(By.ID, "question")
            # 获取题目表单
            form = question_element.find_element(By.CLASS_NAME, "questions")

            # 找到所有题目,包含单选和多选，使用CSS选择器来选择符合其中之一条件的。
            questions = form.find_elements(By.CSS_SELECTOR, ".question,.question2")

            # 存储题目信息的列表，每个题目是一个字典，多个题目用一个列表上传
            answerDTOList = []
            # 遍历每个题目
            for i, question in enumerate(questions, 0):
                # 获取题目文本
                question_text = question.text
                print(f"题目内容: {question_text}")

                # 获取题目ID
                question_qid = question.get_attribute("qid")

                # 获取题目flag，1为单选，2为多选
                question_flag = question.get_attribute("flag")

                # 获取ctype
                ctype_element = driver.find_element(By.CLASS_NAME, "ctype")
                ctype = ctype_element.get_attribute("value")

                # 找到对应的选项容器
                options_container = form.find_element(By.XPATH,
                                                      f"//span[@qid='{question_qid}']/following-sibling::div[@class='neiinput'][1]")

                # 获取选项（包括单选和多选）
                options = options_container.find_elements(By.CSS_SELECTOR,
                                                          "input[type='radio'], input[type='checkbox']")

                # 获取第当前题目的答案
                answer = answers[j]
                # 如果是多选题，答案需要使用","分隔
                if question_flag == "2":
                    answerResult = ",".join(answer)
                else:
                    answerResult = answer

                # 将题目信息封装到字典中
                answer_dict = {
                    "questionId": question_qid,
                    "contentId": caseid,
                    "contentType": ctype,
                    "flag": question_flag,
                    "chapterId": chatpId,
                    "answerResult": answerResult
                }
                answerDTOList.append(answer_dict)

                # 拼接完整的题目
                question_t = question_text + "[" + answer + "]\n"
                answers_t = options_container.text
                timu = question_t + answers_t
                # 打印当前题目
                print(timu)
                tiku = tiku + timu + "\n\n"
                # 准备下一个题目
                j = j + 1

            timucanshu = timucanshu + json.dumps(answerDTOList) + "\n"
        else:
            # 如果不匹配进度标识，则跳过
            print(f"跳过文章: {link_text}")

    # 关闭当前页
    driver.close()


if __name__ == "__main__":
    # 学习列表
    study_list = ["2024年度《八五普法导读》", "2024 应知应会法律知识导读", "2024省教育厅普法读本"]
    # 答案列表
    answer_list = [
        ['A', 'C', 'B', 'B', 'B', 'A', 'C', 'ABC', 'A', 'A', 'A', 'A', 'A', 'A', 'C', 'B', 'C', 'A', 'C', 'A', 'C',
         'ABC',
         'C', 'ABC', 'ABC', 'C', 'AB', 'B', 'C', 'A', 'ABC', 'B', 'ABC', 'C', 'B', 'AB', 'A', 'AB', 'B', 'A', 'B',
         'ABC',
         'B', 'A'],
        ['A', 'C', 'B', 'C', 'A', 'ABC', 'B', 'ABC', 'A', 'C', 'B', 'A', 'ABC', 'ABC', 'B', 'C', 'ABC', 'C', 'A',
         'ABC'],
        ['D', 'A', 'ABC', 'ABC', 'D', 'D', 'AD', 'ABD', 'C', 'D', 'CD', 'AC', 'B', 'B', 'ABCD', 'ABC', 'B', 'D', 'ABCD',
         'ABCD', 'B', 'D', 'ABCD', 'ABCD', 'B', 'D', 'ABCD', 'ABC', 'D', 'C', 'ABCD', 'ABD', 'A', 'C', 'ABCD', 'AC',
         'D',
         'C', 'ABD', 'ABCD']
    ]

    # 保存题库，入口函数定义的变量是模块级的全局变量，其他方法定义的是局部变量
    tiku = ''
    # 存储题目参数
    timucanshu = ''
    driver_path = "driver/chromedriver.exe"
    service = Service(driver_path)
    # 配置Chrome选项
    chrome_options = Options()
    # chrome_options.add_argument("--headless")  # 如果你想要在无头模式下运行
    # chrome_options.add_argument("--disable-gpu")  # 如果你在无头模式下使用Windows
    # chrome_options.add_argument("--no-sandbox")  # 解决DevToolsActivePort文件不存在的报错
    # chrome_options.add_argument("--disable-dev-shm-usage")  # 解决内存不足的问题
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])

    driver = webdriver.Chrome(service=service, options=chrome_options)

    driver = login(driver, "17673058592", "000000")

    first_handle = driver.current_window_handle
    # 获取必修读本
    bixiu_es = driver.find_elements(By.CLASS_NAME, "knowledge_a")

    for i, answers in enumerate(answer_list, 0):
        # 获取对应法律读本链接元素
        study = driver.find_element(By.XPATH, f"//a[contains(., '{study_list[i]}')]")
        study.click()
        time.sleep(5)
        # 切换到新标签页
        driver.switch_to.window(driver.window_handles[-1])
        read_book(driver, answers)
        time.sleep(5)
        # 切回到原始窗口
        driver.switch_to.window(first_handle)

    # 将题库写入文件
    with open("题库2.txt", "wt", encoding="UTF-8") as f:
        f.write(tiku)

        # 将题库写入文件
    with open("答题参数.txt", "wt", encoding="UTF-8") as f:
        f.write(timucanshu)

    time.sleep(10)
    # 关闭浏览器
    driver.quit()
