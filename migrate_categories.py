import pymysql
from dotenv import load_dotenv
import os
import datetime

load_dotenv()

DB_HOST = os.getenv('DB_HOST', '31.97.224.19')
DB_USER = os.getenv('DB_USER', 'spcool')
DB_PASSWORD = os.getenv('DB_PASSWORD', 'Spcool@123')
DB_PORT = int(os.getenv('DB_PORT', '3306'))
DB_NAME = os.getenv('DB_NAME', 'decofurn')

print("Connecting to DB...")
conn = pymysql.connect(host=DB_HOST, user=DB_USER, password=DB_PASSWORD, port=DB_PORT, database=DB_NAME)
cursor = conn.cursor()

# 1. Add category column if missing
cursor.execute("SHOW COLUMNS FROM employee LIKE 'category'")
if not cursor.fetchone():
    cursor.execute("ALTER TABLE employee ADD COLUMN category VARCHAR(50) DEFAULT NULL")
    print("Added 'category' column to 'employee' table.")
else:
    print("'category' column already exists in 'employee' table.")

conn.commit()

# 2. Add the three admin logins
# First, clean up old itms_admin if it exists
cursor.execute("DELETE FROM employee WHERE id = 'itms_admin'")
conn.commit()

admins_to_create = [
    {
        'id': 'smartcity_admin',
        'full_name': 'Smart City Admin',
        'email': 'smartcity@keltron.com',
        'phone': '0000000001',
        'password': 'SmartCity@Admin',
        'is_admin': 1,
        'category': 'Smart City'
    },
    {
        'id': 'iitms_admin',
        'full_name': 'IITMS Admin',
        'email': 'iitms@keltron.com',
        'phone': '0000000002',
        'password': 'IITMS@Admin',
        'is_admin': 1,
        'category': 'IITMS'
    },
    {
        'id': 'construction_admin',
        'full_name': 'Construction Admin',
        'email': 'construction@keltron.com',
        'phone': '0000000003',
        'password': 'Construction@Admin',
        'is_admin': 1,
        'category': 'Construction'
    }
]

for admin in admins_to_create:
    cursor.execute("SELECT id FROM employee WHERE id = %s", (admin['id'],))
    if not cursor.fetchone():
        # Check if email is already taken by some other user
        cursor.execute("SELECT id FROM employee WHERE email = %s", (admin['email'],))
        if cursor.fetchone():
            # generate unique email if somehow already taken
            admin['email'] = f"{admin['id']}_{int(datetime.datetime.now().timestamp())}@keltron.com"
            
        cursor.execute(
            """
            INSERT INTO employee (id, full_name, email, phone, password, is_admin, category, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, NOW())
            """,
            (admin['id'], admin['full_name'], admin['email'], admin['phone'], admin['password'], admin['is_admin'], admin['category'])
        )
        print(f"Created admin user: {admin['id']}")
    else:
        # Update password and category if they already exist
        cursor.execute(
            "UPDATE employee SET password = %s, category = %s, is_admin = 1 WHERE id = %s",
            (admin['password'], admin['category'], admin['id'])
        )
        print(f"Admin user {admin['id']} already exists. Updated credentials/category.")

conn.commit()
cursor.close()
conn.close()
print("Migration completed successfully!")
