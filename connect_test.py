import pyodbc
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import pandas as pd
import sqlalchemy as sa


import pyodbc
connection_string = """DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost;DATABASE=bloodtest;UID=sa;PWD=1234"""
try:
    coxn = pyodbc.connect(connection_string)
except pyodbc.OperationalError:
    connection_string = "DRIVER={ODBC Driver 17 for SQL Server};SERVER=10.30.47.9;DATABASE=bloodtest;UID=henry423;PWD=1234"
finally:
    coxn = pyodbc.connect(connection_string)
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url)

# with coxn.cursor() as cursor:
#     sql = """SELECT [plasma cell],[abnormal lympho],[megakaryocyte],[nRBC],[blast],[metamyelocyte],[eosinophil],[plasmacytoid],[promonocyte],[promyelocyte],[band neutropil],[basopil],[atypical lymphocyte],[hypersegmented neutrophil],[myelocyte],[segmented neutrophil],[lymphocyte],[monocyte]
#                 FROM [bloodtest].[dbo].[test_data] 
#                 JOIN [bloodtest].[dbo].[id]
#                 ON [id].[No] = [test_data].[test_id]
#                 WHERE [test_id] IN (SELECT [id].[No] WHERE [id].[ac]='%s')
#                 AND [smear_id]='%s' 
#                 AND [count]=%d;"""%('admin','B2019_0001',1)
#     cursor.execute(sql)
#     db = cursor.fetchone()
# print(db)
# n=200
# cal = [] 
# for i in range(0,len(db)):
#     a = db[i] / n *100
#     cal.append(a)
# db = [e[0] for e in db]
with engine.begin() as conn:
    filter_test = """SELECT [id].[院區],[id].[ac],[bloodinfo].[year],[blood_final].[smear_id],[blood_final].[count],[blood_final].[celltype],[blood_final].[matrix_value],[blood_final].[timestamp]
FROM [bloodtest].[dbo].[blood_final]
JOIN [bloodtest].[dbo].[id] ON [id].[No] = [blood_final].[test_id]
JOIN [bloodtest].[dbo].[bloodinfo] ON [bloodinfo].[smear_id] = [blood_final].[smear_id];"""
    datadict2 = pd.read_sql_query(sa.text(filter_test), conn)
    # del datadict2['smear_id']
print(len(datadict2))
#     query = "SELECT * FROM [bloodtest].[dbo].[bloodinfo_cbc] WHERE [smear_id]='B2019_0001';"
#     db = pd.read_sql_query(sa.text(query), conn).to_dict('records')[0]
#     del db['smear_id']
#     query_2 = "SELECT [gender] FROM [bloodtest].[dbo].[bloodinfo] WHERE [smear_id]='B2019_0001'"
#     db3 = pd.read_sql_query(sa.text(query_2), conn).value_counts()
#     # for key in db.ke
# lst = [col[0] for col in cursor.description]
# del lst[0]
# print(cal)
# print(db3)
# with coxn.cursor() as cursor:
#     # query = 'SELECT "year","smear_id" FROM [bloodtest].[dbo].[bloodinfo];'
#     query = "SELECT [celltype],[value] FROM [bloodtest].[dbo].[bloodinfo_ans2] WHERE [smear_id]='B2019_0001';"
#     cursor.execute(query)
#     db2 = dict(cursor.fetchall())
# print(db2)