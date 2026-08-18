import os
from dotenv import load_dotenv
import pymysql
from app import app, db

load_dotenv()

def create_database():
    """Create the MySQL database if it doesn't exist"""
    DB_HOST = os.getenv('DB_HOST', '31.97.224.19')
    DB_USER = os.getenv('DB_USER', 'spcool')
    DB_PASSWORD = os.getenv('DB_PASSWORD', 'Spcool@123')
    DB_PORT = int(os.getenv('DB_PORT', '3306'))
    DB_NAME = os.getenv('DB_NAME', 'decofurn')

    try:
        # Connect to MySQL server (without specifying database)
        connection = pymysql.connect(
            host=DB_HOST,
            user=DB_USER,
            password=DB_PASSWORD,
            port=DB_PORT
        )
        
        with connection.cursor() as cursor:
            # Create database if it doesn't exist
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{DB_NAME}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
            print(f"Database '{DB_NAME}' created successfully or already exists")
        
        connection.close()
         
        # Now create tables using SQLAlchemy
        with app.app_context():
            db.create_all()
            print("Tables created successfully")
            
            # Import Employee model to create default admin
            from app import Employee
            
            # Create default admin user if not exists
            admin = Employee.query.filter_by(is_admin=True).first()
            if not admin:
                default_admin = Employee(
                    id='ADMIN001',
                    full_name='System Administrator',
                    email='multisulotionsdecofurn@gmail.com',
                    phone='+919518791736',
                    password='AD#987',
                    is_admin=True
                )
                db.session.add(default_admin)
                db.session.commit()
                print("Default admin created: ID=ADMIN001, Password=admin123")
            else:
                print("Admin user already exists")
                
    except Exception as e:
        print(f"Error creating database: {e}")
        print("Please make sure MySQL server is running and credentials are correct")

if __name__ == '__main__':
    create_database()
