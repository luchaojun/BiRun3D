import pyodbc #导入模块




connectionString = 'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'.format(SERVER='BIOS2',DATABASE = 'SWPRODUCE',USERNAME = 'sa',PASSWORD = 'P@ssw0rd')
print(connectionString)
conn = pyodbc.connect(connectionString)

cursor = conn.cursor()
cursor.execute("select * from  NBSNOCHKLOG WHERE mono='Y6C0386'")
rows = cursor.fetchall()
for row in rows:
    print(row)

# 关闭连接
cursor.close()
conn.close()