import pandas
from sqlalchemy import create_engine, URL

url_object = URL.create(
    drivername="mysql+mysqldb",
    username="lhyzadmin",
    password="LhyzzyhL@1",  # plain (unescaped) text
    host="192.168.161.129",
    port="3306",
    database="lhyzsms"
)
engine = create_engine(url_object)
# 建立连接
# conn = engine.connect()
# print(type(conn))
query = "SELECT * FROM cjfx_studentscore"
df = pandas.read_sql(query, engine)
print(len(df))
