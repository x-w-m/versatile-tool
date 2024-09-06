# 使用题目参数快速答题
import base64
import json
import re
import time

import ddddocr
import requests
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


def dati(session, line):
    # 设置请求头，特别是 Content-Type
    headers = {
        'Content-Type': 'application/json'
    }

    url = "http://hn.12348.gov.cn/fxmain/onlineanswer/ex"

    # 发送 POST 请求，携带字符串格式的 JSON 数据
    response = session.post(url, data=line, headers=headers)
    print(response.text)
    # result = json.loads(response.text)
    # if result["msg"] == "SUCCESS":
    #     return True
    # else:
    #     return False


if __name__ == "__main__":
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

    # 登录
    driver = login(driver, "1", "000000")

    # 登录成功后，获取 cookies
    cookies = driver.get_cookies()
    print(cookies)

    # 创建一个 requests.Session() 对象
    session = requests.Session()
    # 将 Selenium 获取的 cookies 添加到 requests 的会话中
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])

    with open("题目参数.txt", "r") as f:
        for line in f:
            time.sleep(0.6)
            # 每行自带一个换行符，可以使用strip()方法去除。
            print(line.strip())
            if line != "":
                dati(session, line.strip())

    print("学习完毕！，请刷新查看学分。")

    time.sleep(20)
    # 关闭浏览器
    driver.quit()
