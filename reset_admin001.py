import pymysql
import os
from dotenv import load_dotenv

load_dotenv()
conn = pymysql.connect(
    host=os.getenv('DB_HOST', '31.97.224.19'),
    user=os.getenv('DB_USER', 'spcool'),
    password=os.getenv('DB_PASSWORD', 'Spcool@123'),
    port=int(os.getenv('DB_PORT', '3306')),
    database=os.getenv('DB_NAME', 'decofurn')
)
cur = conn.cursor()
try:
    cur.execute("UPDATE employee SET category = NULL, team = NULL WHERE id = 'ADMIN001'")
    conn.commit()
    print("Successfully reset ADMIN001's category and team mappings to NULL.")
except Exception as e:
    conn.rollback()
    print(f"Error resetting ADMIN001 mapping: {e}")
finally:
    conn.close()
