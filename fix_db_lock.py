import os
import pymysql
from dotenv import load_dotenv

load_dotenv()

DB_HOST = os.getenv('DB_HOST', '31.97.224.19')
DB_USER = os.getenv('DB_USER', 'spcool')
DB_PASSWORD = os.getenv('DB_PASSWORD', 'Spcool@123')
DB_PORT = int(os.getenv('DB_PORT', '3306'))
DB_NAME = os.getenv('DB_NAME', 'decofurn')

def fix():
    print(f"Connecting to MySQL server {DB_HOST}:{DB_PORT} / {DB_NAME}...")
    try:
        conn = pymysql.connect(
            host=DB_HOST,
            user=DB_USER,
            password=DB_PASSWORD,
            port=DB_PORT,
            database=DB_NAME,
            autocommit=True
        )
        cursor = conn.cursor()

        # Check existing processlist for any lock/DDL statements
        try:
            cursor.execute("SHOW PROCESSLIST")
            processes = cursor.fetchall()
            print("Current MySQL processes:")
            for p in processes:
                print(p)
        except Exception as e:
            print("Processlist check notice:", e)

        create_table_sql = """
        CREATE TABLE IF NOT EXISTS `location_ping` (
            `id` INT AUTO_INCREMENT PRIMARY KEY,
            `employee_id` VARCHAR(50) NOT NULL,
            `latitude` FLOAT NOT NULL,
            `longitude` FLOAT NOT NULL,
            `accuracy` FLOAT NULL,
            `speed` FLOAT NULL,
            `battery_level` FLOAT NULL,
            `timestamp` DATETIME DEFAULT CURRENT_TIMESTAMP,
            `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (`employee_id`) REFERENCES `employee`(`id`) ON DELETE CASCADE,
            INDEX `idx_loc_ping_time` (`timestamp`),
            INDEX `idx_loc_ping_emp` (`employee_id`)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        
        cursor.execute(create_table_sql)
        print("Successfully created/verified 'location_ping' table in MySQL!")
        conn.close()
    except Exception as e:
        print("Error during table check/creation:", e)

if __name__ == '__main__':
    fix()
