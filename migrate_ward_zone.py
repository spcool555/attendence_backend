"""
Migration script: adds 'ward' and 'zone' columns to the location_list table.

Run once: python migrate_ward_zone.py
"""
import pymysql
from dotenv import load_dotenv
import os

load_dotenv()

DB_HOST = os.getenv('DB_HOST', 'localhost')
DB_USER = os.getenv('DB_USER', 'root')
DB_PASSWORD = os.getenv('DB_PASSWORD', 'Harsha@27')
DB_PORT = int(os.getenv('DB_PORT', '3306'))
DB_NAME = os.getenv('DB_NAME', 'attendance_db')

conn = pymysql.connect(host=DB_HOST, user=DB_USER, password=DB_PASSWORD, port=DB_PORT, database=DB_NAME)
cursor = conn.cursor()

def add_col_if_missing(table, col, col_def):
    cursor.execute(f"SHOW COLUMNS FROM {table} LIKE '{col}'")
    if not cursor.fetchone():
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_def}")
        print(f"Added column '{col}' to '{table}'")
    else:
        print(f"Column '{col}' already exists in '{table}'")

add_col_if_missing('location_list', 'ward', 'VARCHAR(100)')
add_col_if_missing('location_list', 'zone', 'VARCHAR(100)')

conn.commit()
cursor.close()
conn.close()

print("\nMigration complete: ward & zone columns added to location_list.")
