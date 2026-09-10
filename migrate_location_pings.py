import os
import pymysql
from dotenv import load_dotenv
from app import app, db

load_dotenv()

DB_HOST = os.getenv('DB_HOST', '31.97.224.19')
DB_USER = os.getenv('DB_USER', 'spcool')
DB_PASSWORD = os.getenv('DB_PASSWORD', 'Spcool@123')
DB_PORT = int(os.getenv('DB_PORT', '3306'))
DB_NAME = os.getenv('DB_NAME', 'decofurn')

def migrate():
    print(f"Migrating JunctionVisit & LocationPing schema in database {DB_NAME}...")
    try:
        conn = pymysql.connect(
            host=DB_HOST,
            user=DB_USER,
            password=DB_PASSWORD,
            port=DB_PORT,
            database=DB_NAME
        )
        cursor = conn.cursor()

        # 1. Add time_spent_minutes column to junction_visit if missing
        try:
            cursor.execute("ALTER TABLE junction_visit ADD COLUMN time_spent_minutes FLOAT DEFAULT 0.0")
            print("Added time_spent_minutes column.")
        except Exception as e:
            if "Duplicate column name" in str(e):
                print("time_spent_minutes column already exists.")
            else:
                print("Notice on time_spent_minutes:", e)

        # 2. Add travel_time_minutes column to junction_visit if missing
        try:
            cursor.execute("ALTER TABLE junction_visit ADD COLUMN travel_time_minutes FLOAT DEFAULT 0.0")
            print("Added travel_time_minutes column.")
        except Exception as e:
            if "Duplicate column name" in str(e):
                print("travel_time_minutes column already exists.")
            else:
                print("Notice on travel_time_minutes:", e)

        conn.commit()
        conn.close()

        # 3. Create missing tables (e.g. location_ping) via SQLAlchemy
        with app.app_context():
            db.create_all()
            print("All SQLAlchemy tables (including location_ping) created successfully!")

    except Exception as e:
        print("Migration notice / error:", e)

if __name__ == '__main__':
    migrate()
