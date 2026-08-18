from app import app, db
from sqlalchemy import text

with app.app_context():
    print("Executing DB migration...")
    try:
        db.session.execute(text("ALTER TABLE junction_visit ADD COLUMN visit_type VARCHAR(50) DEFAULT 'Regular Visit'"))
        print("Added visit_type column.")
    except Exception as e:
        print(f"Error adding visit_type (might already exist): {e}")

    try:
        db.session.execute(text("ALTER TABLE junction_visit ADD COLUMN remark TEXT"))
        print("Added remark column.")
    except Exception as e:
        print(f"Error adding remark (might already exist): {e}")

    db.session.commit()
    print("Migration finished!")
