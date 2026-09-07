import pymysql
import os
from dotenv import load_dotenv

load_dotenv()
conn = pymysql.connect(
    host=os.getenv('DB_HOST', '31.97.224.19'),
    user=os.getenv('DB_USER', 'spcool'),
    password=os.getenv('DB_PASSWORD', 'Spcool@123'),
    database=os.getenv('DB_NAME', 'decofurn')
)
cur = conn.cursor()
cur.execute('select id, full_name, team, category, is_admin from employee')
rows = cur.fetchall()
with open('db_employees.txt', 'w') as f:
    for r in rows:
        f.write(str(r) + '\n')
conn.close()
print("Success")
