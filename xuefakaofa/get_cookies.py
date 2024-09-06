# 手动登录并保存cookies
import pickle

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time

# 配置Chrome选项
chrome_options = Options()
# chrome_options.add_argument("--headless")  # 如果你想要在无头模式下运行
# chrome_options.add_argument("--disable-gpu")  # 如果你在无头模式下使用Windows
# chrome_options.add_argument("--no-sandbox")  # 解决DevToolsActivePort文件不存在的报错
# chrome_options.add_argument("--disable-dev-shm-usage")  # 解决内存不足的问题

driver = webdriver.Chrome(options=chrome_options)

try:
    # 打开目标网站
    driver.get(
        "http://hn.12348.gov.cn/ucenter/#/login?redirect=http%3A%2F%2Fhn.12348.gov.cn%2Ffxmain%2Fstudy%3Fbiz%3Dfxmain")  # 替换为你要登录的网站URL

    # 等待手动登录
    time.sleep(15)  # 你可以改为WebDriverWait以便更精确的控制等待时间

    # 保存Cookies到文件
    with open("cookies.pkl", "wb") as file:
        pickle.dump(driver.get_cookies(), file)

    # 等待页面跳转
    time.sleep(5)


finally:
    # 关闭浏览器
    driver.quit()
