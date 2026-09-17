import os
import pymysql
from dotenv import load_dotenv

load_dotenv()

DB_HOST = os.getenv('DB_HOST', '31.97.224.19')
DB_USER = os.getenv('DB_USER', 'spcool')
DB_PASSWORD = os.getenv('DB_PASSWORD', 'Spcool@123')
DB_PORT = int(os.getenv('DB_PORT', '3306'))
DB_NAME = os.getenv('DB_NAME', 'decofurn')

def add_column():
    print(f"Connecting to MySQL database {DB_NAME} on {DB_HOST}:{DB_PORT}...")
    conn = pymysql.connect(
        host=DB_HOST,
        user=DB_USER,
        password=DB_PASSWORD,
        port=DB_PORT,
        database=DB_NAME,
        autocommit=True
    )
    cursor = conn.cursor()

    # Kill sleeping or blocked queries
    cursor.execute("SHOW PROCESSLIST")
    procs = cursor.fetchall()
    my_id = conn.thread_id()
    for p in procs:
        p_id = p[0]
        p_command = p[4]
        if p_id != my_id and p_command != 'Daemon':
            try:
                cursor.execute(f"KILL {p_id}")
                print(f"Killed process {p_id}")
            except Exception as e:
                print(f"Failed to kill process {p_id}: {e}")

    # Check if column project exists
    cursor.execute("SHOW COLUMNS FROM employee LIKE 'project'")
    res = cursor.fetchall()
    if not res:
        print("Adding column 'project' to table 'employee'...")
        cursor.execute("ALTER TABLE `employee` ADD COLUMN `project` VARCHAR(50) DEFAULT 'Smart City' NULL")
        print("Column 'project' added successfully!")
    else:
        print("Column 'project' already exists in 'employee' table.")

    conn.close()

if __name__ == '__main__':
    add_column()
