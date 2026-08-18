"""
Migration script: adds 'designation' and 'team' columns to the employee table,
then populates values from the Excel data provided.

Run once: python migrate_team_columns.py
"""
import pymysql
from dotenv import load_dotenv
import os

load_dotenv()

DB_HOST = os.getenv('DB_HOST', '31.97.224.19')
DB_USER = os.getenv('DB_USER', 'spcool')
DB_PASSWORD = os.getenv('DB_PASSWORD', 'Spcool@123')
DB_PORT = int(os.getenv('DB_PORT', '3306'))
DB_NAME = os.getenv('DB_NAME', 'decofurn')

conn = pymysql.connect(host=DB_HOST, user=DB_USER, password=DB_PASSWORD, port=DB_PORT, database=DB_NAME)
cursor = conn.cursor()

# ---------- 1. Add columns if they don't exist ----------
def add_col_if_missing(col, col_def):
    cursor.execute(f"SHOW COLUMNS FROM employee LIKE '{col}'")
    if not cursor.fetchone():
        cursor.execute(f"ALTER TABLE employee ADD COLUMN {col} {col_def}")
        print(f"Added column: {col}")
    else:
        print(f"Column already exists: {col}")

add_col_if_missing('designation', 'VARCHAR(100)')
add_col_if_missing('team', "VARCHAR(20) COMMENT 'field | coc | ccc'")

conn.commit()

# ---------- 2. Employee mapping from Excel ----------
# Format: (employee_id, designation, team)
EMPLOYEE_MAP = [
    # === FIELD TEAM ===
    ('EMP024', 'Network Engineer',         'field'),
    ('EMP025', 'Sub Zone Leader',          'field'),
    ('EMP027', 'Sr. Help Desk Engineer',   'field'),
    ('EMP028', 'Helper',                   'field'),
    ('EMP029', 'Passive Engineer',         'field'),
    ('EMP030', 'Sub Zone Leader',          'field'),
    ('EMP031', 'Zone Leader',              'field'),
    ('EMP034', 'Passive Engineer',         'field'),
    ('EMP037', 'Helper',                   'field'),
    ('EMP038', 'Zone Leader',              'field'),
    ('EMP039', 'Field Maintenance Engineer', 'field'),
    ('EMP040', 'Office Executive',         'field'),
    ('EMP041', 'Zones Passive Head',       'field'),
    ('EMP046', 'Zones Passive Head',       'field'),
    ('EMP047', 'Sr. Engineer',             'field'),
    ('EMP050', 'Zone Leader',              'field'),
    ('EMP051', 'Sub-Zone Leader',          'field'),
    ('EMP053', 'Helper',                   'field'),
    ('EMP054', 'Field Engineer Head',      'field'),
    ('EMP055', 'CTO',                      'field'),
    ('EMP061', 'Sr. BMS Engineer',         'field'),
    ('EMP063', 'Driver Tata Ace',          'field'),
    ('EMP066', 'Civil Site Engineer',      'field'),
    ('EMP067', '',                         'field'),
    ('EMP070', 'Sub Zone Leader',          'field'),
    ('EMP071', 'Field Engineer',           'field'),
    ('EMP072', 'Field Engineer',           'field'),
    ('EMP074', 'Sub Zone Leader',          'field'),
    ('EMP076', 'Account Executive',        'field'),

    # === COC TEAM ===
    ('EMP011', 'Operations Manager',       'coc'),
    ('EMP013', 'Sr. BMS Engineer',         'coc'),
    ('EMP035', 'Cyber Security Head',      'coc'),
    ('EMP044', 'Accounts Manager',         'coc'),
    ('EMP049', 'Air Conditioner Refrigerator Head', 'coc'),
    ('EMP056', 'Office Boy',               'coc'),
    ('EMP057', 'Godown Incharge',          'coc'),
    ('EMP058', 'Administrative Officer',   'coc'),
    ('EMP059', 'L1 Engineer',              'coc'),
    ('EMP062', 'Network Engineer',         'coc'),
    ('EMP065', 'IT Infra Head',            'coc'),
    ('EMP068', 'Desktop Engineer',         'coc'),
    ('EMP073', 'Finance & Accounts (Smart City)', 'coc'),
    ('EMP075', 'Assistant Accountant',    'coc'),
    ('EMP077', 'DC Operations Engineer',   'coc'),
    ('EMP078', 'DC Operations Engineer',   'coc'),
    ('EMP080', 'Project Manager',          'coc'),
    ('EMP085', 'DC Operations Engineer',   'coc'),

    # === CCC TEAM ===
    ('EMP032', 'Software Engineer',        'ccc'),
    ('EMP033', 'Sr. BMS Engineer',         'ccc'),
    ('EMP036', 'Operations Manager Head',  'ccc'),
    ('EMP042', 'Software Engineer',        'ccc'),
    ('EMP043', 'Software Engineer',        'ccc'),
    ('EMP048', 'Software Engineer',        'ccc'),
    ('EMP052', 'Software Engineer',        'ccc'),
    ('EMP064', 'Operation Manager',        'ccc'),
    ('EMP069', 'L1 Engineer',              'ccc'),
    ('EMP079', 'DC Operations Engineer',   'ccc'),
    ('EMP081', '',                         'ccc'),
    ('EMP082', '',                         'ccc'),
    ('EMP083', '',                         'ccc'),
    ('EMP084', 'DC Operations Engineer',   'ccc'),
]

updated = 0
for emp_id, designation, team in EMPLOYEE_MAP:
    cursor.execute(
        "UPDATE employee SET designation=%s, team=%s WHERE id=%s",
        (designation, team, emp_id)
    )
    if cursor.rowcount:
        updated += 1

conn.commit()
cursor.close()
conn.close()

print(f"\nMigration complete. {updated} employee records updated.")
