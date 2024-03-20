# 解析从官网获取到的评分标准，保存为xlsx文件
from lxml import html
import pandas as pd

with open("男生评分标准.txt", "r", encoding="utf-8") as f:
    text = f.read()
    print(text)

html_string = text
# 解析HTML
tree = html.fromstring(html_string)

# 提取表格
table = tree.xpath('//tbody')[0]

# 读取所有行
rows = table.xpath('.//tr')

# 解析表头
headers = [th.text_content().strip() for th in rows[0].xpath('.//th')]

# 解析数据行
data = []
for row in rows[1:]:
    data.append([td.text_content().strip() for td in row.xpath('.//td')])

# 转换为DataFrame
df = pd.DataFrame(data, columns=headers)

# 保存到xlsx文件
df.to_excel('高三男生.xlsx', index=False)
