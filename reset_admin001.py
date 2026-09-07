import pymysql
import os
from dotenv import load_dotenv

load_dotenv()
conn = pymysql.connect(
    host=os.getenv('DB_HOST', 'localhost'),
    user=os.getenv('DB_USER', 'root'),
    password=os.getenv('DB_PASSWORD', 'Harsha@27'),
    database=os.getenv('DB_NAME', 'attendance_db')
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
