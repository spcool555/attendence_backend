"""
Migration script to add half-day support columns to Leave table
Run this script once to update existing database
"""
from app import app, db
from sqlalchemy import text

def migrate_leave_table():
    with app.app_context():
        try:
            # Check if columns already exist
            with db.engine.connect() as conn:
                # Add is_half_day column if it doesn't exist
                try:
                    conn.execute(text("""
                        ALTER TABLE leave 
                        ADD COLUMN is_half_day BOOLEAN DEFAULT FALSE
                    """))
                    conn.commit()
                    print("✓ Added is_half_day column")
                except Exception as e:
                    if "Duplicate column name" in str(e) or "already exists" in str(e):
                        print("✓ is_half_day column already exists")
                    else:
                        print(f"✗ Error adding is_half_day: {e}")
                
                # Add half_day_period column if it doesn't exist
                try:
                    conn.execute(text("""
                        ALTER TABLE leave 
                        ADD COLUMN half_day_period VARCHAR(20)
                    """))
                    conn.commit()
                    print("✓ Added half_day_period column")
                except Exception as e:
                    if "Duplicate column name" in str(e) or "already exists" in str(e):
                        print("✓ half_day_period column already exists")
                    else:
                        print(f"✗ Error adding half_day_period: {e}")
            
            print("\n✅ Migration completed successfully!")
            print("You can now restart your Flask application.")
            
        except Exception as e:
            print(f"\n❌ Migration failed: {e}")
            print("Please check your database connection and try again.")

if __name__ == '__main__':
    print("Starting Leave table migration...")
    print("Adding half-day support columns...\n")
    migrate_leave_table()
