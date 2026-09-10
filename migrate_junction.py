import os
import pymysql
from dotenv import load_dotenv

load_dotenv()

DB_HOST = os.getenv('DB_HOST', '31.97.224.19')
DB_USER = os.getenv('DB_USER', 'spcool')
DB_PASSWORD = os.getenv('DB_PASSWORD', 'Spcool@123')
DB_PORT = int(os.getenv('DB_PORT', '3306'))
DB_NAME = os.getenv('DB_NAME', 'decofurn')

def migrate():
    print(f"Connecting to {DB_USER}@{DB_HOST}:{DB_PORT}/{DB_NAME}")
    try:
        conn = pymysql.connect(
            host=DB_HOST,
            user=DB_USER,
            password=DB_PASSWORD,
            database=DB_NAME
        )
        cursor = conn.cursor()
        
        # Add visit_type if not exists
        try:
            cursor.execute("ALTER TABLE junction_visit ADD COLUMN visit_type VARCHAR(50) DEFAULT 'Regular Visit'")
            print("Added visit_type column.")
        except Exception as e:
            if "Duplicate column name" in str(e):
                print("visit_type column already exists.")
            else:
                raise
                
        # Add remark if not exists
        try:
            cursor.execute("ALTER TABLE junction_visit ADD COLUMN remark TEXT")
            print("Added remark column.")
        except Exception as e:
            if "Duplicate column name" in str(e):
                print("remark column already exists.")
            else:
                raise
                
        conn.commit()
        conn.close()
        print("Migration complete!")
    except Exception as e:
        print(f"Error connecting or migrating: {e}")

if __name__ == '__main__':
    migrate()
