from flask import Flask, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
from datetime import datetime, time, date, timedelta
import pytz
import os
from werkzeug.utils import secure_filename
import pandas as pd
from io import BytesIO
import uuid
import pymysql
from dotenv import load_dotenv
from urllib.parse import quote_plus
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
#from twilio.rest import Client
from sqlalchemy import or_, and_
from datetime import date
from calendar import monthrange

# ADD / UPDATE THIS (top of file ya helper section me)
SUPPORTED_LANGUAGES = [
    {'code': 'en', 'name': 'English'},
    {'code': 'hi', 'name': 'Hindi'},
    {'code': 'mr', 'name': 'Marathi'}
]

# TIMEZONE SETUP
UTC = pytz.UTC
IST = pytz.timezone("Asia/Kolkata")

def to_ist(dt):
    if dt is None:
        return None

    #IF DATETIME IS NOT NATIVE
    if dt.tzinfo is None:
        dt = UTC.localize(dt)

    return dt.astimezone(IST).isoformat()


def utc_to_ist_dt(dt):
    if dt is None:
        return None
    if dt.tzinfo is None:
        dt = UTC.localize(dt)
    return dt.astimezone(IST)

from openpyxl.drawing.image import Image as OpenPyxlImage
from openpyxl.utils import get_column_letter

# Load environment variables
load_dotenv()

# Install PyMySQL as MySQLdb
pymysql.install_as_MySQLdb()

def get_image_file_path(filename):
    if not filename or str(filename).strip() in ('—', 'None', ''):
        return None
    p1 = os.path.join(app.config.get('UPLOAD_FOLDER', 'uploads'), filename)
    if os.path.exists(p1):
        return p1
    base_dir = os.path.dirname(os.path.abspath(__file__))
    p2 = os.path.join(base_dir, 'uploads', filename)
    if os.path.exists(p2):
        return p2
    p3 = os.path.join(os.path.dirname(base_dir), 'uploads', filename)
    if os.path.exists(p3):
        return p3
    return None

def embed_photos_in_excel(writer, sheet_name, image_map, row_height=65, max_size=(75, 75), merge_ranges=None):
    if sheet_name not in writer.sheets:
        return
    ws = writer.sheets[sheet_name]
    
    header_col_map = {}
    for col in range(1, ws.max_column + 1):
        cell_val = ws.cell(row=1, column=col).value
        if cell_val:
            header_col_map[str(cell_val).strip()] = (col, get_column_letter(col))

    from PIL import Image as PILImage
    from openpyxl.styles import Alignment
    import io

    # Apply vertical cell merging for common columns in group ranges
    if merge_ranges:
        # Merge all columns except Col 8 (Visit Date), Col 9 (Started At), Col 14 (Before Photo), Col 16 (Before Location), and Col 18 (Before Remark)
        columns_to_merge = [1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 13, 15, 17, 19]
        for start_row, end_row in merge_ranges:
            for col_num in columns_to_merge:
                if col_num <= ws.max_column:
                    try:
                        ws.merge_cells(start_row=start_row, start_column=col_num, end_row=end_row, end_column=col_num)
                        # Set top-left cell alignment to center content vertically
                        top_cell = ws.cell(row=start_row, column=col_num)
                        top_cell.alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
                    except Exception as merge_err:
                        print(f"Error merging rows {start_row}-{end_row} in column {col_num}: {merge_err}")

    # Apply height and photo embedding
    for item in image_map:
        row_idx = item['row_idx']
        excel_row = row_idx + 2  # Row 1 is header
        ws.row_dimensions[excel_row].height = row_height
        
        for col_name, photo_filename in item.get('photos', {}).items():
            if not photo_filename or col_name not in header_col_map:
                continue
            
            file_path = get_image_file_path(photo_filename)
            if file_path:
                try:
                    col_num, col_letter = header_col_map[col_name]
                    ws.column_dimensions[col_letter].width = 16
                    
                    # Open with PIL, convert to RGB, compress and make thumbnail for file size reduction
                    with PILImage.open(file_path) as pil_img:
                        if pil_img.mode in ('RGBA', 'LA') or (pil_img.mode == 'P' and 'transparency' in pil_img.info):
                            pil_img = pil_img.convert('RGB')
                        # Make thumbnail at max 150x150 width/height
                        pil_img.thumbnail((150, 150), PILImage.Resampling.LANCZOS)
                        
                        img_byte_arr = io.BytesIO()
                        pil_img.save(img_byte_arr, format='JPEG', quality=75)
                        img_byte_arr.seek(0)
                        img = OpenPyxlImage(img_byte_arr)
                    
                    img.width, img.height = max_size
                    ws.add_image(img, f'{col_letter}{excel_row}')
                except Exception as err:
                    print(f"Error embedding image {photo_filename} into Excel: {err}")

    # Set cell alignment to wrap text for multi-line columns
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and '\n' in cell.value:
                cell.alignment = Alignment(wrap_text=True, vertical='center')

app = Flask(__name__)
#CORS(app)
CORS(app,resources={r"/api/": {"origins": ""}},
     supports_credentials=True
)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key-here')

# Database configuration
DB_HOST = os.getenv('DB_HOST', '31.97.224.19')
DB_USER = os.getenv('DB_USER', 'spcool')
DB_PASSWORD = quote_plus(os.getenv('DB_PASSWORD', 'Spcool@123'))
DB_NAME = os.getenv('DB_NAME', 'decofurn')
DB_PORT = os.getenv('DB_PORT', '3306')
app.config['SQLALCHEMY_DATABASE_URI'] = f'mysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create uploads directory if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

db = SQLAlchemy(app)
CORS(app, resources={r"/api/*": {"origins": "*"}})

# Database Models
class Employee(db.Model):
    id = db.Column(db.String(50), primary_key=True)
    full_name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    phone = db.Column(db.String(20), nullable=False)
    password = db.Column(db.String(100), nullable=False)
    is_admin = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))

    # ✅ ADD THIS
    shift_type = db.Column(db.String(20), default='general')

    weekly_off = db.Column(db.String(10))
    preferred_language = db.Column(db.String(5), default='en')

    # ✅ ADD-ON: team bifurcation (field | coc | ccc) + designation
    # Columns are added to the actual MySQL table by migrate_team_columns.py
    team = db.Column(db.String(20))
    designation = db.Column(db.String(100))
    category = db.Column(db.String(50))

    # Relationship with attendance records
    attendance_records = db.relationship('Attendance', backref='employee', lazy=True)

class Attendance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(50), db.ForeignKey('employee.id'), nullable=False)
    date = db.Column(db.Date, nullable=False)
    check_in_time = db.Column(db.DateTime)
    check_out_time = db.Column(db.DateTime)
    check_in_photo = db.Column(db.String(255))  # Filename only
    check_out_photo = db.Column(db.String(255))  # Filename only
    check_in_location = db.Column(db.String(100))  # GPS coordinates
    check_out_location = db.Column(db.String(100))  # GPS coordinates
    status = db.Column(db.String(20))  # 'present', 'late', 'half_day', 'leave'
    user_message = db.Column(db.Text)  # Optional user message
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))
    office_time = db.Column(db.Time)
    shift_type = db.Column(db.String(20))
    
class Leave(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(50), db.ForeignKey('employee.id'), nullable=False)
    leave_type = db.Column(db.String(20), nullable=False)  # 'sick', 'emergency', 'Compensatory_off', 'lwp'
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    reason = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(20), default='pending')  # 'pending', 'approved', 'rejected'
    admin_comment = db.Column(db.Text)
    is_half_day = db.Column(db.Boolean, default=False)
    half_day_period = db.Column(db.String(20))  # 'first_half', 'second_half'
    supporting_document = db.Column(db.String(255))  # Filename for uploaded document (image or PDF)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))
    updated_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC), onupdate=lambda: datetime.now(pytz.UTC))
    
    # Relationship
    employee = db.relationship('Employee', backref='leaves')
    
class Announcement(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    message = db.Column(db.Text, nullable=False)
    created_by = db.Column(db.String(50), db.ForeignKey('employee.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))

# Helper Functions
def allowed_file(filename):
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def allowed_document(filename):
    """Check if file is allowed for leave documents (images and PDFs)"""
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'pdf'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

SHIFT_TIMINGS = {
    'general': {
        'start': time(9, 30),
        'late': time(9, 45),
        'half_day': time(13, 30),
        'end': time(18, 0),
        'early_checkout': time(17, 30)
    },
    'morning': {
        'start': time(8, 0),
        'late': time(8, 15),
        'half_day': time(12, 15),
        'end': time(16, 30),
        'early_checkout': time(16, 0)
    },
    'evening': {
        'start': time(16, 0),
        'late': time(16, 15),
        'half_day': time(20, 0),
        'end': time(23, 59, 59),
        'early_checkout': time(23, 30)
    },
    'night': {
        'start': time(0, 0),
        'late': time(0, 15),
        'half_day': time(4, 0),
        'end': time(8, 0),
        'early_checkout': time(7, 30)
    }
}

def calculate_attendance_status(check_in_time, shift_type, check_out_time=None):
    if not check_in_time or shift_type not in SHIFT_TIMINGS:
        return 'absent'

    shift = SHIFT_TIMINGS[shift_type]

    # ALWAYS convert to IST datetime
    check_in_dt = utc_to_ist_dt(check_in_time)
    check_out_dt = utc_to_ist_dt(check_out_time) if check_out_time else None

    check_in = check_in_dt.time()
    check_out = check_out_dt.time() if check_out_dt else None

    # FIX: If checking out after midnight for the evening shift, max out time so it passes early_checkout checks
    if shift_type == 'evening' and check_out and check_out.hour < 12:
        check_out = time(23, 59, 59)
        
    # FIX: If checking in early for the night shift (e.g. 11:50 PM), floor it to 00:00 to pass late checks
    if shift_type == 'night' and check_in.hour >= 20:
        check_in = time(0, 0)

    # ✅ PRESENT / SECOND HALF ABSENT
    if check_in <= shift['late']:
        if check_out and check_out < shift['early_checkout']:
            return 'half_day_second_half'
        return 'present'

    # ✅ LATE
    elif check_in < shift['half_day']:
        return 'late'

    # ✅ FIRST HALF ABSENT
    else:
        return 'half_day_first_half'
def send_email(to_email, subject, message):
    try:
        smtp_email = os.getenv("SMTP_EMAIL")
        smtp_password = os.getenv("SMTP_PASSWORD")


        msg = MIMEMultipart()
        msg['From'] = smtp_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(smtp_email, smtp_password)
        server.send_message(msg)
        server.quit()

        print(f"✅ Email sent to {to_email}")
    except Exception as e:
        print("❌ Email error:", e)

def notify_admin_leave_request(employee, leave):
    admins = Employee.query.filter_by(is_admin=True).all()

    for admin in admins:
        msg = (
            f"New Leave Request\n\n"
            f"Employee: {employee.full_name} ({employee.id})\n"
            f"Leave Type: {leave.leave_type}\n"
            f"Dates: {leave.start_date} to {leave.end_date}\n"
            f"Reason: {leave.reason}"
        )
        send_email(
            admin.email,
            "New Leave Request Submitted",
            msg
        )

def notify_employee_leave_status(employee, leave):
    msg = (
        f"Your leave request ({leave.leave_type}) "
        f"from {leave.start_date} to {leave.end_date} "
        f"is {leave.status.upper()}."
    )

    if leave.admin_comment:
        msg += f"\nComment: {leave.admin_comment}"

   # send_sms(employee.phone, msg)
    send_email(employee.email, "Leave Status Update", msg)
    
# Authentication Routes
@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    employee_id = data.get('employee_id')
    password = data.get('password')

    if not employee_id or not password:
        return jsonify({'error': 'Employee ID and password are required'}), 400

    # Retrieve employee record without password filter
    employee = Employee.query.filter_by(id=employee_id).first()

    if not employee:
        return jsonify({'error': 'Invalid credentials'}), 401

    # Admin authentication: check stored employee password or ADMIN_PASSWORD env variable
    if employee.is_admin:
        admin_pass = os.getenv('ADMIN_PASSWORD')
        if employee.password == password or (admin_pass and password == admin_pass):
            return jsonify({
                'success': True,
                'employee': {
                    'id': employee.id,
                    'full_name': employee.full_name,
                    'email': employee.email,
                    'phone': employee.phone,
                    'is_admin': employee.is_admin,
                    'team': employee.team,
                    'designation': employee.designation,
                    'category': employee.category
                }
            })
        else:
            return jsonify({'error': 'Invalid admin credentials'}), 401

    # Regular employee authentication using stored password
    if employee.password == password:
        return jsonify({
            'success': True,
            'employee': {
                'id': employee.id,
                'full_name': employee.full_name,
                'email': employee.email,
                'phone': employee.phone,
                'is_admin': employee.is_admin,
                'team': employee.team,
                'designation': employee.designation,
                'category': employee.category
            }
        })
    else:
        return jsonify({'error': 'Invalid credentials'}), 401

@app.route('/api/attendance/check-in', methods=['POST','OPTIONS'])
def check_in():
    shift_type = request.form.get('shift_type') or 'general'
    employee_id = request.form.get('employee_id')
    location = request.form.get('location')
    user_message = request.form.get('user_message', '')

    if not employee_id or not location:
        return jsonify({'error': 'Employee ID and location are required'}), 400

    employee = Employee.query.get(employee_id)
    if not employee:
        return jsonify({'error': 'Employee not found'}), 404

    today = datetime.utcnow().date()

    today_records = Attendance.query.filter(
        Attendance.employee_id == employee_id,
        Attendance.date == today
    ).order_by(Attendance.id.asc()).all()

    active_shift = None
    for r in today_records:
        if r.check_in_time and not r.check_out_time:
            active_shift = r
            break

    if active_shift:
        return jsonify({
            'error': 'Please check out of your current shift before checking in again (Shift 1 must be completed first).'
        }), 400

    if len(today_records) >= 2:
        return jsonify({
            'error': 'Maximum 2 shifts allowed per day'
        }), 400

    # Do not allow re-using the same shift type on the 2nd check-in (same day).
    used_shift_types = [
        (r.shift_type or 'general') for r in today_records if r.check_in_time
    ]
    if shift_type in used_shift_types:
        return jsonify({
            'error': f'You already used the "{shift_type}" shift today. Please select a different shift.'
        }), 400

    # Photo handling
    photo_filename = None
    if 'photo' in request.files:
        file = request.files['photo']
        if file and allowed_file(file.filename):
            filename = secure_filename(
                f"{employee_id}{datetime.now().strftime('%Y%m%d%H%M%S')}_checkin.{file.filename.rsplit('.', 1)[1].lower()}"
            )
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            photo_filename = filename

    if not photo_filename:
        return jsonify({'error': 'Photo is required'}), 400

    check_in_time = datetime.utcnow()
    status = calculate_attendance_status(check_in_time, shift_type)

    # ✅ ALWAYS NEW ENTRY
    attendance = Attendance(
        employee_id=employee_id,
        date=today,
        check_in_time=check_in_time,
        check_in_photo=photo_filename,
        check_in_location=location,
        status=status,
        shift_type=shift_type,
        user_message=user_message
    )

    db.session.add(attendance)
    db.session.commit()

    return jsonify({
        'success': True,
        'message': 'Checked in successfully',
        'check_in_time': to_ist(check_in_time)
    })

@app.route('/api/attendance/check-out', methods=['POST'])
def check_out():
    """Check-out is allowed at any time after check-in; status rules (e.g. half-day) apply from timings."""
    employee_id = request.form.get('employee_id')
    location = request.form.get('location')
    
    if not employee_id or not location:
        return jsonify({'error': 'Employee ID and location are required'}), 400
    
    # Check if employee exists
    employee = Employee.query.get(employee_id)
    if not employee:
        return jsonify({'error': 'Employee not found'}), 404
    
    # Check if checked in today
    today = datetime.utcnow().date()
    
    # Find the latest shift where check-out is NOT done
    attendance = Attendance.query.filter_by(
    employee_id=employee_id,
    date=today,
    check_out_time=None
    ).order_by(Attendance.id.desc()).first()

    if not attendance:
        return jsonify({
        'error': 'No active shift found. Please check in first.'
    }), 400

    if attendance.check_out_time:
        return jsonify({
        'error': 'You have already checked out for this shift.'
    }), 400
    
    if not attendance or not attendance.check_in_time:
        return jsonify({'error': 'Must check in before checking out'}), 400
    
    if attendance.check_out_time:
        return jsonify({'error': 'Already checked out today'}), 400
    
    # Handle photo upload
    photo_filename = None
    if 'photo' in request.files:
        file = request.files['photo']
        if file and allowed_file(file.filename):
            filename = secure_filename(f"{employee_id}{datetime.now().strftime('%Y%m%d%H%M%S')}_checkout.{file.filename.rsplit('.', 1)[1].lower()}")
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            photo_filename = filename
    
    if not photo_filename:
        return jsonify({'error': 'Photo is required for check-out'}), 400
    
    check_out_time = datetime.utcnow()
    
    # Calculate total office time
    work_duration = check_out_time - attendance.check_in_time
    total_hours = work_duration.total_seconds() / 3600
    office_time_hours = int(total_hours)
    office_time_minutes = int((total_hours - office_time_hours) * 60)
    print(office_time_hours,office_time_minutes)
    office_time = time(office_time_hours, office_time_minutes)

    # Update attendance record
    attendance.check_out_time = check_out_time
    attendance.check_out_photo = photo_filename
    attendance.check_out_location = location
    attendance.office_time = office_time
    attendance.shift_type = attendance.shift_type
    
    # Recalculate status with check-out time
    attendance.status = calculate_attendance_status(attendance.check_in_time, attendance.shift_type, check_out_time)
    
    db.session.commit()
    
    return jsonify({
        'success': True,
        'message': 'Checked out successfully',
        'check_out_time': to_ist(check_out_time),
        'office_time': office_time.strftime('%H:%M'),
        'total_hours': f"{office_time_hours}h {office_time_minutes}m"
    })
from datetime import date
from calendar import monthrange

@app.route('/api/attendance/monthly/<employee_id>', methods=['GET'])
def get_monthly_attendance(employee_id):
    month = int(request.args.get('month', date.today().month))
    year = int(request.args.get('year', date.today().year))

    records = Attendance.query.filter(
        Attendance.employee_id == employee_id,
        db.extract('month', Attendance.date) == month,
        db.extract('year', Attendance.date) == year
    ).order_by(Attendance.date.asc()).all()

    data = []
    for r in records:
        data.append({
                'date': r.date.isoformat(),
                'status': r.status,
                'check_in': to_ist(r.check_in_time) if r.check_in_time else None,
                'check_out': to_ist(r.check_out_time) if r.check_out_time else None,
                'office_time': r.office_time.strftime('%H:%M') if r.office_time else None,
                'shift_type': r.shift_type if r.shift_type else 'general'
})


    return jsonify(data), 200

@app.route('/api/attendance/active/<employee_id>', methods=['GET'])
def get_active_shift(employee_id):
    today = datetime.utcnow().date()

    record = Attendance.query.filter(
        Attendance.employee_id == employee_id,
        Attendance.date == today,
        Attendance.check_out_time.is_(None)
    ).order_by(Attendance.id.desc()).first()

    if not record:
        return jsonify({"active": False})

    return jsonify({
        "active": True,
        "shift_type": record.shift_type,
        "check_in_time": to_ist(record.check_in_time),
    })
@app.route('/api/attendance/status/<employee_id>', methods=['GET'])
def get_attendance_status(employee_id):
    today = datetime.utcnow().date()
    today_records = Attendance.query.filter(
        Attendance.employee_id == employee_id,
        Attendance.date == today
    ).order_by(Attendance.id.asc()).all()

    shifts = []
    for idx, r in enumerate(today_records):
        shifts.append({
            'shift_number': idx + 1,
            'shift_type': r.shift_type or 'general',
            'check_in_time': to_ist(r.check_in_time) if r.check_in_time else None,
            'check_out_time': to_ist(r.check_out_time) if r.check_out_time else None,
            'office_time': r.office_time.strftime('%H:%M') if r.office_time else None,
            'status': r.status,
        })

    shift_count = len(today_records)
    used_shift_types = []
    for r in today_records:
        if r.check_in_time:
            used_shift_types.append(r.shift_type or 'general')
    completed_shifts = sum(
        1 for r in today_records if r.check_in_time and r.check_out_time
    )
    active = next(
        (r for r in today_records if r.check_in_time and not r.check_out_time),
        None
    )
    remaining_slots = max(0, 2 - shift_count)

    if not today_records:
        return jsonify({
            'checked_in': False,
            'checked_out': False,
            'status': 'not_checked_in',
            'check_in_time': None,
            'check_out_time': None,
            'office_time': None,
            'remaining': remaining_slots,
            'shifts_today': 0,
            'completed_shifts_today': 0,
            'active_shift_number': None,
            'first_shift_check_in_done': False,
            'can_check_in_again': True,
            'day_complete': False,
            'used_shift_types': [],
            'shifts': [],
        }), 200

    if active:
        active_num = today_records.index(active) + 1
        return jsonify({
            'checked_in': True,
            'checked_out': False,
            'check_in_time': to_ist(active.check_in_time),
            'check_out_time': None,
            'office_time': None,
            'status': active.status,
            'remaining': remaining_slots,
            'shifts_today': shift_count,
            'completed_shifts_today': completed_shifts,
            'active_shift_number': active_num,
            'active_shift_type': active.shift_type or 'general',
            'first_shift_check_in_done': True,
            'can_check_in_again': False,
            'day_complete': False,
            'used_shift_types': used_shift_types,
            'shifts': shifts,
        }), 200

    last = today_records[-1]
    day_fully_done = shift_count >= 2 and completed_shifts >= 2
    return jsonify({
        'checked_in': False,
        'checked_out': bool(last.check_out_time),
        'check_in_time': to_ist(last.check_in_time) if last.check_in_time else None,
        'check_out_time': to_ist(last.check_out_time) if last.check_out_time else None,
        'office_time': last.office_time.strftime('%H:%M') if last.office_time else None,
        'status': last.status,
        'remaining': remaining_slots,
        'shifts_today': shift_count,
        'completed_shifts_today': completed_shifts,
        'active_shift_number': None,
        'first_shift_check_in_done': shift_count >= 1,
        'can_check_in_again': shift_count < 2 and not day_fully_done,
        'day_complete': day_fully_done,
        'used_shift_types': used_shift_types,
        'shifts': shifts,
    }), 200

# Admin Routes
@app.route('/api/admin/employees', methods=['GET'])
def get_employees():
    category = request.args.get('category')
    if category:
        employees = Employee.query.filter_by(category=category).all()
    else:
        employees = Employee.query.all()
    return jsonify([{
        'id': emp.id,
        'full_name': emp.full_name,
        'email': emp.email,
        'phone': emp.phone,
        'is_admin': emp.is_admin,
        'team': emp.team,
        'designation': emp.designation,
        'category': emp.category,
        'created_at': emp.created_at.isoformat()
    } for emp in employees])

@app.route('/api/admin/employees', methods=['POST'])
def create_employee():
    data = request.get_json()
    
    required_fields = ['id', 'full_name', 'email', 'phone', 'password']
    for field in required_fields:
        if not data.get(field):
            return jsonify({'error': f'{field} is required'}), 400
    
    # Check if employee ID already exists
    if Employee.query.get(data['id']):
        return jsonify({'error': 'Employee ID already exists'}), 400
    
    # Check if email already exists
    if Employee.query.filter_by(email=data['email']).first():
        return jsonify({'error': 'Email already exists'}), 400
    
    # Map designation to team
    designation = data.get('designation', '')
    team_mapped = None
    if designation:
        desig_lower = designation.lower()
        if 'field' in desig_lower:
            team_mapped = 'field'
        elif 'coc' in desig_lower or 'cooc' in desig_lower:
            team_mapped = 'coc'
        elif 'ccc' in desig_lower:
            team_mapped = 'ccc'
        elif 'towing' in desig_lower:
            team_mapped = 'towing'
            
    employee = Employee(
        id=data['id'],
        full_name=data['full_name'],
        email=data['email'],
        phone=data['phone'],
        password=data['password'],  # Plain text as requested
        is_admin=data.get('is_admin', False),
        shift_type=data.get('shift_type', 'general'),
        weekly_off=data.get('weekly_off') or 'Sunday',
        category=data.get('category'),
        designation=designation,
        team=team_mapped
        )
    
    db.session.add(employee)
    db.session.commit()
    
    return jsonify({
        'success': True,
        'message': 'Employee created successfully',
        'employee': {
            'id': employee.id,
            'full_name': employee.full_name,
            'email': employee.email,
            'phone': employee.phone,
            'is_admin': employee.is_admin,
            'category': employee.category,
            'designation': employee.designation,
            'team': employee.team
        }
    })

@app.route('/api/admin/attendance', methods=['GET'])
def get_attendance_logs():
    employee_id = request.args.get('employee_id')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    category = request.args.get('category')
    
    query = db.session.query(Attendance, Employee).join(Employee)
    
    if category:
        query = query.filter(Employee.category == category)
        
    if employee_id:
        query = query.filter(Attendance.employee_id == employee_id)
    
    if start_date:
        query = query.filter(Attendance.date >= datetime.strptime(start_date, '%Y-%m-%d').date())
    
    if end_date:
        query = query.filter(Attendance.date <= datetime.strptime(end_date, '%Y-%m-%d').date())
    
    results = query.all()
    
    attendance_logs = []
    for attendance, employee in results:
        attendance_logs.append({
            'id': attendance.id,
            'employee_id': attendance.employee_id,
            'employee_name': employee.full_name,
            'date': attendance.date.isoformat(),
            'check_in_time': to_ist(attendance.check_in_time) if attendance.check_in_time else None,
            'check_out_time': to_ist(attendance.check_out_time) if attendance.check_out_time else None,
            'office_time': attendance.office_time.strftime('%H:%M') if attendance.office_time else None,
            'status': attendance.status,
            'shift_type': attendance.shift_type,
            'user_message': attendance.user_message,
            'check_in_location': attendance.check_in_location,
            'check_out_location': attendance.check_out_location,
            'check_in_photo': attendance.check_in_photo,
            'check_out_photo': attendance.check_out_photo
        })
    attendance_logs.reverse()
    return jsonify(attendance_logs)

@app.route('/api/admin/attendance/export', methods=['GET'])
def export_attendance():
    employee_id = request.args.get('employee_id')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    category = request.args.get('category')
    
    query = db.session.query(Attendance, Employee).join(Employee)
    
    if category:
        query = query.filter(Employee.category == category)
        
    if employee_id:
        query = query.filter(Attendance.employee_id == employee_id)
    
    if start_date:
        query = query.filter(Attendance.date >= datetime.strptime(start_date, '%Y-%m-%d').date())
    
    if end_date:
        query = query.filter(Attendance.date <= datetime.strptime(end_date, '%Y-%m-%d').date())
    
    results = query.all()
    
    # Create DataFrame
    data = []
    image_map = []
    for idx, (attendance, employee) in enumerate(results):
        data.append({
            'Employee ID': attendance.employee_id,
            'Employee Name': employee.full_name,
            'Category': employee.category or '',
            'Date': attendance.date.strftime('%Y-%m-%d'),
            'Check In Time': utc_to_ist_dt(attendance.check_in_time).strftime('%H:%M:%S') if attendance.check_in_time else '',
            'Check Out Time': utc_to_ist_dt(attendance.check_out_time).strftime('%H:%M:%S') if attendance.check_out_time else '',
            'Office Time': attendance.office_time.strftime('%H:%M') if attendance.office_time else '',
            'Status': attendance.status,
            'Shift Type': attendance.shift_type or '',
            'User Message': attendance.user_message or '',
            'Check In Photo': attendance.check_in_photo or '',
            'Check Out Photo': attendance.check_out_photo or '',
            'Check In Location': attendance.check_in_location or '',
            'Check Out Location': attendance.check_out_location or ''
        })
        image_map.append({
            'row_idx': idx,
            'photos': {
                'Check In Photo': attendance.check_in_photo,
                'Check Out Photo': attendance.check_out_photo
            }
        })
    
    df = pd.DataFrame(data)
    
    # Create Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Attendance')
        embed_photos_in_excel(writer, 'Attendance', image_map)
    
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'attendance_report_{datetime.now().strftime("%Y%m%d")}.xlsx'
    )

@app.route('/api/admin/stats', methods=['GET'])
def get_attendance_stats():
    today = date.today()
    category = request.args.get('category')
    
    # Get today's attendance stats
    today_stats_query = db.session.query(Attendance.status, db.func.count(Attendance.id)).join(Employee).filter(
        Attendance.date == today,
        Employee.is_admin == False
    )
    if category:
        today_stats_query = today_stats_query.filter(Employee.category == category)
        
    today_stats = today_stats_query.group_by(Attendance.status).all()
    
    stats = {
        'present': 0,
        'late': 0,
        'half_day_first_half': 0,
        'half_day_second_half': 0,
        'absent': 0
    }
    
    for status, count in today_stats:
        if status in stats:
            stats[status] = count
    
    # Calculate absent count (employees who didn't check in)
    total_employees_query = Employee.query.filter_by(is_admin=0)
    if category:
        total_employees_query = total_employees_query.filter_by(category=category)
    total_employees = total_employees_query.count()
    
    checked_in_today_query = Attendance.query.join(Employee).filter(
        Attendance.date == today,
        Attendance.check_in_time.isnot(None),
        Employee.is_admin == 0
    )
    if category:
        checked_in_today_query = checked_in_today_query.filter(Employee.category == category)
    checked_in_today = checked_in_today_query.count()
    
    stats['absent'] = total_employees - checked_in_today
    stats['total_employees'] = total_employees
    
    return jsonify(stats)

# Image serving endpoint
@app.route('/api/images/<filename>', methods=['GET'])
def serve_image(filename):
    """Serve uploaded images"""
    try:
        # Security check: ensure filename doesn't contain path traversal
        if '..' in filename or '/' in filename or '\\' in filename:
            return jsonify({'error': 'Invalid filename'}), 400
        
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # Check if file exists
        if not os.path.exists(file_path):
            return jsonify({'error': 'Image not found'}), 404
        
        return send_file(file_path)
    except Exception as e:
        return jsonify({'error': 'Failed to serve image'}), 500


# Admin Password Management Routes
@app.route('/api/admin/employee/<employee_id>', methods=['GET'])
def get_employee_by_id(employee_id):
    """Get a specific employee by ID for password management"""
    employee = Employee.query.filter_by(id=employee_id).first()
    
    if not employee:
        return jsonify({'error': 'Employee not found'}), 404
    
    return jsonify({
        'id': employee.id,
        'full_name': employee.full_name,
        'email': employee.email,
        'phone': employee.phone,
        'is_admin': employee.is_admin,
        'team': employee.team,
        'designation': employee.designation,
        'category': employee.category,
        'created_at': employee.created_at.isoformat()
    })

@app.route('/api/admin/employee/<employee_id>/password', methods=['PUT'])
def change_employee_password(employee_id):
    """Change password for a specific employee"""
    data = request.get_json()
    
    if not data or 'new_password' not in data:
        return jsonify({'error': 'New password is required'}), 400
    
    new_password = data['new_password'].strip()
    
    if len(new_password) < 6:
        return jsonify({'error': 'Password must be at least 6 characters long'}), 400
    
    employee = Employee.query.filter_by(id=employee_id).first()
    
    if not employee:
        return jsonify({'error': 'Employee not found'}), 404
    
    try:
        employee.password = new_password
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': f'Password updated successfully for {employee.full_name}'
        })
    
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': 'Failed to update password'}), 500
    
    #Language Route
    
@app.route('/api/user/language', methods=['PUT'])
def update_language():
    if request.method == 'OPTIONS':
        return jsonify({'success': True}), 200
    
    data = request.get_json()
    employee_id = data.get('employee_id')
    language = data.get('language')

    if language not in [l['code'] for l in SUPPORTED_LANGUAGES]:
        return jsonify({'error': 'Invalid language'}), 400

    employee = Employee.query.get(employee_id)
    if not employee:
        return jsonify({'error': 'Employee not found'}), 404

    employee.preferred_language = language
    db.session.commit()

    return jsonify({
        'success': True,
        'language': language
    })

# Leave Management Routes
@app.route('/api/leave/request', methods=['POST'])
def request_leave():
    """Employee submits a leave request"""
    # Check if request has JSON or form data
    if request.is_json:
        data = request.get_json()
        document_file = None
    else:
        data = request.form.to_dict()
        document_file = request.files.get('supporting_document')
    
    required_fields = ['employee_id', 'leave_type', 'start_date', 'end_date', 'reason']
    for field in required_fields:
        if not data.get(field):
            return jsonify({'error': f'{field} is required'}), 400
    
    # Validate employee exists
    employee = Employee.query.get(data['employee_id'])
    if not employee:
        return jsonify({'error': 'Employee not found'}), 404
    
    # Parse dates
    try:
        start_date = datetime.strptime(data['start_date'], '%Y-%m-%d').date()
        end_date = datetime.strptime(data['end_date'], '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'error': 'Invalid date format. Use YYYY-MM-DD'}), 400
    
    # Get half-day info
    is_half_day = data.get('is_half_day', 'false').lower() == 'true' if isinstance(data.get('is_half_day'), str) else bool(data.get('is_half_day', False))
    half_day_period = data.get('half_day_period', 'first_half')
    
    # Validate dates
    if start_date > end_date:
        return jsonify({'error': 'Start date must be before or equal to end date'}), 400
    
    # Allow current day and future dates
    if start_date < date.today():
        return jsonify({'error': 'Cannot request leave for past dates'}), 400
    
    # Validate half-day: must be single day
    if is_half_day and start_date != end_date:
        return jsonify({'error': 'Half-day leave must be for a single day only'}), 400
    
    # Check for overlapping leave requests
    overlapping = Leave.query.filter(
        Leave.employee_id == data['employee_id'],
        Leave.status != 'rejected',
        db.or_(
            db.and_(Leave.start_date <= start_date, Leave.end_date >= start_date),
            db.and_(Leave.start_date <= end_date, Leave.end_date >= end_date),
            db.and_(Leave.start_date >= start_date, Leave.end_date <= end_date)
        )
    ).first()
    
    if overlapping:
        return jsonify({'error': 'You already have a leave request for overlapping dates'}), 400
    
    # Handle document upload
    document_filename = None
    if document_file and document_file.filename:
        if allowed_document(document_file.filename):
            # Generate unique filename
            ext = document_file.filename.rsplit('.', 1)[1].lower()
            document_filename = f"leave_{data['employee_id']}{datetime.now().strftime('%Y%m%d%H%M%S')}.{ext}"
            document_path = os.path.join(app.config['UPLOAD_FOLDER'], document_filename)
            document_file.save(document_path)
        else:
            return jsonify({'error': 'Invalid document format. Only images (PNG, JPG, JPEG, GIF) and PDF files are allowed'}), 400
    
    # --- Auto-LWP logic ---
    # Check if sick + emergency leaves are exhausted
    current_year = date.today().year
    approved_leaves = Leave.query.filter(
        Leave.employee_id == data['employee_id'],
        Leave.status.in_(['approved', 'pending']),
        db.extract('year', Leave.start_date) == current_year
    ).all()
    
    sick_used = 0
    emergency_used = 0
    comp_used = 0
    for lv in approved_leaves:
        lv_half = getattr(lv, 'is_half_day', False)
        lv_days = 0.5 if lv_half else (lv.end_date - lv.start_date).days + 1
        if lv.leave_type == 'sick':
            sick_used += lv_days
        elif lv.leave_type == 'emergency':
            emergency_used += lv_days
        elif lv.leave_type == 'Compensatory_off':
            comp_used += lv_days
    
    requested_leave_type = data['leave_type']
    
    # If sick and emergency both exhausted, force LWP
    both_exhausted = (sick_used >= 8 and emergency_used >= 8)
    
    if both_exhausted:
        if requested_leave_type in ('sick', 'emergency'):
            requested_leave_type = 'lwp'
    else:
        # Enforce balances if not both exhausted
        requested_days = 0.5 if is_half_day else (end_date - start_date).days + 1
        
        if requested_leave_type == 'sick' and (sick_used + requested_days) > 8:
            return jsonify({'error': f'Insufficient Sick Leave balance. Remaining: {max(0, 8.0 - sick_used)}, Requested: {requested_days}'}), 400
            
        if requested_leave_type == 'emergency' and (emergency_used + requested_days) > 8:
            return jsonify({'error': f'Insufficient Emergency Leave balance. Remaining: {max(0, 8.0 - emergency_used)}, Requested: {requested_days}'}), 400
    
    # Create leave request
    leave = Leave(
        employee_id=data['employee_id'],
        leave_type=requested_leave_type,
        start_date=start_date,
        end_date=end_date,
        reason=data['reason'],
        status='pending',
        is_half_day=is_half_day,
        half_day_period=half_day_period if is_half_day else None,
        supporting_document=document_filename
    )
    
    db.session.add(leave)
    db.session.commit()
    notify_admin_leave_request(employee, leave)
    
    return jsonify({
        'success': True,
        'message': 'Leave request submitted successfully',
        'leave': {
            'id': leave.id,
            'leave_type': leave.leave_type,
            'start_date': leave.start_date.isoformat(),
            'end_date': leave.end_date.isoformat(),
            'status': leave.status,
            'supporting_document': document_filename
        }
    }), 201

@app.route('/api/leave/employee/<employee_id>', methods=['GET'])
def get_employee_leaves(employee_id):
    """Get all leave requests for a specific employee"""
    leaves = Leave.query.filter_by(employee_id=employee_id).order_by(Leave.created_at.desc()).all()
    
    result = []
    for leave in leaves:
        # Safe attribute access for backward compatibility
        is_half_day = getattr(leave, 'is_half_day', False)
        half_day_period = getattr(leave, 'half_day_period', None)
        
        result.append({
            'id': leave.id,
            'leave_type': leave.leave_type,
            'start_date': leave.start_date.isoformat(),
            'end_date': leave.end_date.isoformat(),
            'reason': leave.reason,
            'status': leave.status,
            'admin_comment': leave.admin_comment,
            'is_half_day': is_half_day,
            'half_day_period': half_day_period,
            'supporting_document': getattr(leave, 'supporting_document', None),
            'created_at': to_ist(leave.created_at),
            'days_count': 0.5 if is_half_day else (leave.end_date - leave.start_date).days + 1
        })
    
    return jsonify(result)

@app.route('/api/admin/leaves', methods=['GET'])
def get_all_leaves():
    """Admin gets all leave requests with optional filters"""
    status_filter = request.args.get('status')
    employee_id = request.args.get('employee_id')
    category = request.args.get('category')
    
    query = db.session.query(Leave, Employee).join(Employee)
    if category:
        query = query.filter(Employee.category == category)
    
    if status_filter:
        query = query.filter(Leave.status == status_filter)
    
    if employee_id:
        query = query.filter(Leave.employee_id == employee_id)
    
    results = query.order_by(Leave.created_at.desc()).all()
    
    current_year = date.today().year
    emp_stats_cache = {}

    def get_emp_stats(emp_id):
        if emp_id in emp_stats_cache:
            return emp_stats_cache[emp_id]

        all_leaves = Leave.query.filter(
            Leave.employee_id == emp_id,
            db.extract('year', Leave.start_date) == current_year
        ).all()

        sick_used = 0
        emergency_used = 0
        comp_used = 0
        lwp_used = 0
        approved_days = 0

        pending_count = 0
        pending_days = 0

        for l in all_leaves:
            is_hd = getattr(l, 'is_half_day', False)
            d_count = 0.5 if is_hd else (l.end_date - l.start_date).days + 1

            if l.status == 'approved':
                approved_days += d_count
                if l.leave_type == 'sick':
                    sick_used += d_count
                elif l.leave_type == 'emergency':
                    emergency_used += d_count
                elif l.leave_type == 'Compensatory_off':
                    comp_used += d_count
                elif l.leave_type == 'lwp':
                    lwp_used += d_count
            elif l.status == 'pending':
                pending_count += 1
                pending_days += d_count

        total_sick = 8
        total_emergency = 8
        total_comp = 8

        stats = {
            'approved_days_taken': approved_days,
            'pending_requests_count': pending_count,
            'pending_days_count': pending_days,
            'sick_used': sick_used,
            'emergency_used': emergency_used,
            'comp_used': comp_used,
            'lwp_used': lwp_used,
            'sick_remaining': max(0, total_sick - sick_used),
            'emergency_remaining': max(0, total_emergency - emergency_used),
            'comp_remaining': max(0, total_comp - comp_used),
            'total_remaining': max(0, total_sick - sick_used) + max(0, total_emergency - emergency_used) + max(0, total_comp - comp_used)
        }
        emp_stats_cache[emp_id] = stats
        return stats

    leaves = []
    for leave, employee in results:
        # Safe attribute access for backward compatibility
        is_half_day = getattr(leave, 'is_half_day', False)
        half_day_period = getattr(leave, 'half_day_period', None)
        
        leaves.append({
            'id': leave.id,
            'employee_id': leave.employee_id,
            'employee_name': employee.full_name,
            'employee_email': employee.email,
            'employee_phone': employee.phone,
            'leave_type': leave.leave_type,
            'start_date': leave.start_date.isoformat(),
            'end_date': leave.end_date.isoformat(),
            'reason': leave.reason,
            'status': leave.status,
            'admin_comment': leave.admin_comment,
            'is_half_day': is_half_day,
            'half_day_period': half_day_period,
            'supporting_document': getattr(leave, 'supporting_document', None),
            'created_at': to_ist(leave.created_at),
            'days_count': 0.5 if is_half_day else (leave.end_date - leave.start_date).days + 1,
            'employee_stats': get_emp_stats(leave.employee_id)
        })
    
    return jsonify(leaves)
@app.route('/api/admin/employees-by-status', methods=['GET'])
def get_employees_by_status():
    status = request.args.get('status')
    category = request.args.get('category')
    today = date.today()

    if status == 'absent':
        # Absent = jinhone check-in nahi kiya
        checked_ids_query = db.session.query(Attendance.employee_id).filter(
            Attendance.date == today
        )
        if category:
            checked_ids_query = checked_ids_query.join(Employee).filter(Employee.category == category)
            
        checked_ids = checked_ids_query
        query = db.session.query(Employee.id, Employee.full_name).filter(
            ~Employee.id.in_(checked_ids),
            Employee.is_admin == False
        )
        if category:
            query = query.filter(Employee.category == category)
    else:
        # Join Attendance + Employee
        query = db.session.query(Employee.id, Employee.full_name).join(Attendance).filter(
            Attendance.date == today,
            Employee.is_admin == False
        )
        if category:
            query = query.filter(Employee.category == category)
            
        # Status filter (distinct employees — multiple shifts today can duplicate rows)
        if status == 'present':
            query = query.filter(Attendance.status == 'present')
        elif status == 'late':
            query = query.filter(Attendance.status == 'late')
        elif status in ('half_day', 'half_day_first_half', 'half_day_second_half'):
            query = query.filter(Attendance.status.in_(['half_day_first_half', 'half_day_second_half']))

    result = query.distinct().all()

    return jsonify([
        {"employee_id": emp_id, "employee_name": name}
        for emp_id, name in result
    ])

@app.route('/api/admin/leave/<int:leave_id>/approve', methods=['PUT'])
def approve_leave(leave_id):
    """Admin approves a leave request"""
    data = request.get_json()
    
    leave = Leave.query.get(leave_id)
    if not leave:
        return jsonify({'error': 'Leave request not found'}), 404
    
    if leave.status != 'pending':
        return jsonify({'error': 'Leave request is already processed'}), 400
    
    leave.status = 'approved'
    leave.admin_comment = data.get('admin_comment', '')
    leave.updated_at = datetime.utcnow()
    
    db.session.commit()
    employee = Employee.query.get(leave.employee_id)
    notify_employee_leave_status(employee, leave)
    employee = Employee.query.get(leave.employee_id)
    msg = f"""
Your leave request has been APPROVED

Leave Type: {leave.leave_type}
Dates: {leave.start_date} to {leave.end_date}
Comment: {leave.admin_comment or 'N/A'}
"""

    send_email(employee.email, "Leave Approved", msg)
    
    return jsonify({
        'success': True,
        'message': 'Leave request approved successfully'
    })

@app.route('/api/admin/leave/<int:leave_id>/reject', methods=['PUT'])
def reject_leave(leave_id):
    """Admin rejects a leave request"""
    data = request.get_json()
    
    leave = Leave.query.get(leave_id)
    if not leave:
        return jsonify({'error': 'Leave request not found'}), 404
    
    if leave.status != 'pending':
        return jsonify({'error': 'Leave request is already processed'}), 400
    
    leave.status = 'rejected'
    leave.admin_comment = data.get('admin_comment', '')
    leave.updated_at = datetime.utcnow()
    
    db.session.commit()
    employee = Employee.query.get(leave.employee_id)
    notify_employee_leave_status(employee, leave)
    employee = Employee.query.get(leave.employee_id)

    msg = f"""
Your leave request has been REJECTED

Leave Type: {leave.leave_type}
Dates: {leave.start_date} to {leave.end_date}
Reason: {leave.admin_comment}
"""

    send_email(employee.email, "Leave Rejected", msg)
    
    return jsonify({
        'success': True,
        'message': 'Leave request rejected successfully'
    })

@app.route('/api/leave/<int:leave_id>/edit', methods=['PUT'])
def edit_leave(leave_id):
    """Employee edits a leave request (only if pending)"""
    # Check if request has JSON or form data
    if request.is_json:
        data = request.get_json()
        document_file = None
    else:
        data = request.form.to_dict()
        document_file = request.files.get('supporting_document')
    
    leave = Leave.query.get(leave_id)
    if not leave:
        return jsonify({'error': 'Leave request not found'}), 404
    
    # Only allow editing if status is pending
    if leave.status != 'pending':
        return jsonify({'error': 'Cannot edit leave request after it has been approved or rejected'}), 400
    
    # Update fields if provided
    if 'leave_type' in data:
        leave.leave_type = data['leave_type']
    
    if 'start_date' in data and 'end_date' in data:
        try:
            start_date = datetime.strptime(data['start_date'], '%Y-%m-%d').date()
            end_date = datetime.strptime(data['end_date'], '%Y-%m-%d').date()
            
            if start_date > end_date:
                return jsonify({'error': 'Start date must be before or equal to end date'}), 400
            
            if start_date < date.today():
                return jsonify({'error': 'Cannot request leave for past dates'}), 400
            
            leave.start_date = start_date
            leave.end_date = end_date
        except ValueError:
            return jsonify({'error': 'Invalid date format. Use YYYY-MM-DD'}), 400
    
    if 'reason' in data:
        leave.reason = data['reason']
    
    if 'is_half_day' in data:
        is_half_day = data.get('is_half_day', 'false').lower() == 'true' if isinstance(data.get('is_half_day'), str) else bool(data.get('is_half_day', False))
        leave.is_half_day = is_half_day
        
        if is_half_day and leave.start_date != leave.end_date:
            return jsonify({'error': 'Half-day leave must be for a single day only'}), 400
    
    if 'half_day_period' in data:
        leave.half_day_period = data['half_day_period'] if leave.is_half_day else None
    
    # Handle document upload
    if document_file and document_file.filename:
        if allowed_document(document_file.filename):
            # Delete old document if exists
            if leave.supporting_document:
                old_path = os.path.join(app.config['UPLOAD_FOLDER'], leave.supporting_document)
                if os.path.exists(old_path):
                    os.remove(old_path)
            
            # Save new document
            ext = document_file.filename.rsplit('.', 1)[1].lower()
            document_filename = f"leave_{leave.employee_id}{datetime.now().strftime('%Y%m%d%H%M%S')}.{ext}"
            document_path = os.path.join(app.config['UPLOAD_FOLDER'], document_filename)
            document_file.save(document_path)
            leave.supporting_document = document_filename
        else:
            return jsonify({'error': 'Invalid document format. Only images (PNG, JPG, JPEG, GIF) and PDF files are allowed'}), 400
    
    leave.updated_at = datetime.utcnow()
    db.session.commit()
    
    return jsonify({
        'success': True,
        'message': 'Leave request updated successfully',
        'leave': {
            'id': leave.id,
            'leave_type': leave.leave_type,
            'start_date': leave.start_date.isoformat(),
            'end_date': leave.end_date.isoformat(),
            'reason': leave.reason,
            'status': leave.status,
            'is_half_day': leave.is_half_day,
            'half_day_period': leave.half_day_period,
            'supporting_document': leave.supporting_document
        }
    })

@app.route('/api/leave/document/<filename>', methods=['GET'])
def get_leave_document(filename):
    """Get leave supporting document"""
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path)
        else:
            return jsonify({'error': 'Document not found'}), 404
    except Exception as e:
        return jsonify({'error': 'Failed to retrieve document'}), 500

@app.route('/api/leave/stats/<employee_id>', methods=['GET'])
def get_leave_stats(employee_id):
    """Get leave statistics for an employee"""
    current_year = date.today().year
    
    # Get approved and pending leaves for current year so balances reflect immediately
    active_leaves = Leave.query.filter(
        Leave.employee_id == employee_id,
        Leave.status.in_(['approved', 'pending']),
        db.extract('year', Leave.start_date) == current_year
    ).all()
    
    # Calculate days taken by leave type
    sick_days = 0
    emergency_days = 0
    Compensatory_off_days = 0
    lwp_days = 0
    
    for leave in active_leaves:
        # Safe attribute access for backward compatibility
        is_half_day = getattr(leave, 'is_half_day', False)
        days = 0.5 if is_half_day else (leave.end_date - leave.start_date).days + 1
        if leave.leave_type == 'sick':
            sick_days += days
        elif leave.leave_type == 'emergency':
            emergency_days += days
        elif leave.leave_type == 'Compensatory_off':
            Compensatory_off_days += days
        elif leave.leave_type == 'lwp':
            lwp_days += days
    
    # Define leave balances (8 each)
    total_sick = 8
    total_emergency = 8
    total_Compensatory_off = 8
    total_paid_leaves = total_sick + total_emergency + total_Compensatory_off  # 24
    
    total_used = sick_days + emergency_days + Compensatory_off_days + lwp_days
    
    # LWP activates when sick + emergency are both exhausted
    all_paid_exhausted = (sick_days >= total_sick and 
                          emergency_days >= total_emergency)
    
    return jsonify({
        'sick_leave': {
            'total': total_sick,
            'used': sick_days,
            'remaining': max(0, total_sick - sick_days)
        },
        'emergency_leave': {
            'total': total_emergency,
            'used': emergency_days,
            'remaining': max(0, total_emergency - emergency_days)
        },
        'Compensatory_leave': {
            'total': total_Compensatory_off,
            'used': Compensatory_off_days,
            'remaining': max(0, total_Compensatory_off - Compensatory_off_days)
        },
        'lwp': {
            'used': lwp_days,
            'active': all_paid_exhausted
        },
        'total_used': total_used,
        'all_paid_exhausted': all_paid_exhausted
    })
@app.route('/api/admin/announcements', methods=['POST'])
def create_announcement():
    data = request.get_json()

    message = data.get('message')
    admin_id = data.get('admin_id')

    if not message:
        return jsonify({'error': 'Message is required'}), 400

    announcement = Announcement(
        message=message,
        created_by=admin_id
    )

    db.session.add(announcement)
    db.session.commit()

    return jsonify({
        'success': True,
        'message': 'Announcement posted successfully'
    })
@app.route('/api/announcements', methods=['GET'])
def get_announcements():
    announcements = Announcement.query.order_by(Announcement.created_at.desc()).all()

    return jsonify([
        {
            'id': a.id, 
            'message': a.message,
            'created_by': a.created_by,
            'created_at': to_ist(a.created_at) if a.created_at else None
        }
        for a in announcements
    ])
@app.route('/api/admin/announcements/<int:id>', methods=['DELETE'])
def delete_announcement(id):
    announcement = Announcement.query.get(id)

    if not announcement:
        return jsonify({'error': 'Announcement not found'}), 404

    db.session.delete(announcement)
    db.session.commit()

    return jsonify({'success': True, 'message': 'Deleted successfully'})

# ============================================================
# ADD-ON: Junction Visit tracking (Field Team, COC Team & CCC Team) and
# per-team Admin dashboards (COC / CCC / Field). Nothing above
# this block was modified beyond adding the `team`/`designation`
# columns and echoing them back in a few existing responses.
# ============================================================

class JunctionList(db.Model):
    """Stores the list of predefined junction names uploaded via Excel (Field team - legacy)."""
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), unique=True, nullable=False)
    ward = db.Column(db.String(100), nullable=True)
    zone = db.Column(db.String(100), nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))


class LocationList(db.Model):
    """Stores predefined location names per team (field | coc | ccc) uploaded via Excel.
    This replaces/extends JunctionList with team-based filtering."""
    __tablename__ = 'location_list'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), nullable=False)
    team = db.Column(db.String(20), nullable=False)  # 'field' | 'coc' | 'ccc'
    ward = db.Column(db.String(100), nullable=True)
    zone = db.Column(db.String(100), nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))
    __table_args__ = (db.UniqueConstraint('name', 'team', name='uq_location_name_team'),)


class JunctionVisit(db.Model):
    """A single junction/site visit made by a roaming (Field/COC/CCC) employee
    during their shift. 'Before' photo + location are captured on arrival,
    'After' photo + location are captured when the visit is completed."""
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(50), db.ForeignKey('employee.id'), nullable=False)
    junction_name = db.Column(db.String(150), nullable=False)
    ward = db.Column(db.String(100), nullable=True)
    zone = db.Column(db.String(100), nullable=True)
    date = db.Column(db.Date, nullable=False)
    before_photo = db.Column(db.String(255))
    after_photo = db.Column(db.String(255))
    before_location = db.Column(db.String(100))
    after_location = db.Column(db.String(100))
    started_at = db.Column(db.DateTime)
    completed_at = db.Column(db.DateTime)
    status = db.Column(db.String(20), default='in_progress')  # 'in_progress' | 'completed'
    visit_type = db.Column(db.String(50), default='Regular Visit')
    remark = db.Column(db.Text)
    before_remark = db.Column(db.Text)
    asset_type = db.Column(db.String(100))
    fault_type = db.Column(db.String(100))
    time_spent_minutes = db.Column(db.Float, default=0.0)
    travel_time_minutes = db.Column(db.Float, default=0.0)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))

    employee = db.relationship('Employee', backref='junction_visits')


class LocationPing(db.Model):
    """Stores background/periodic GPS location pings of roaming/field employees during shift hours."""
    __tablename__ = 'location_ping'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(50), db.ForeignKey('employee.id'), nullable=False)
    latitude = db.Column(db.Float, nullable=False)
    longitude = db.Column(db.Float, nullable=False)
    accuracy = db.Column(db.Float, nullable=True)
    speed = db.Column(db.Float, nullable=True)
    battery_level = db.Column(db.Float, nullable=True)
    timestamp = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC), index=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))

    employee = db.relationship('Employee', backref=db.backref('location_pings', lazy='dynamic'))



class CallVisit(db.Model):
    """A call visit made to resolve an unresolved junction visit."""
    __tablename__ = 'call_visit'
    id = db.Column(db.Integer, primary_key=True)
    junction_visit_id = db.Column(db.Integer, db.ForeignKey('junction_visit.id'), nullable=False)
    before_photo = db.Column(db.String(255))
    after_photo = db.Column(db.String(255))
    before_location = db.Column(db.String(100))
    after_location = db.Column(db.String(100))
    started_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))
    completed_at = db.Column(db.DateTime)
    status = db.Column(db.String(20), default='in_progress')  # 'in_progress' | 'completed' | 'unresolved'
    remark = db.Column(db.Text)
    before_remark = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))

    junction_visit = db.relationship('JunctionVisit', backref=db.backref('call_visits', lazy=True))


class AssetFaultMapping(db.Model):
    """Stores predefined asset types and their associated fault types uploaded via Excel."""
    __tablename__ = 'asset_fault_mapping'
    id = db.Column(db.Integer, primary_key=True)
    asset_type = db.Column(db.String(100), nullable=False)
    fault_type = db.Column(db.String(100), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))
    __table_args__ = (db.UniqueConstraint('asset_type', 'fault_type', name='uq_asset_fault'),)


class JunctionRemark(db.Model):
    """Stores general remarks or issue messages posted by the field/roaming team
    regarding specific junctions or general issues."""
    __tablename__ = 'junction_remark'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(50), db.ForeignKey('employee.id'), nullable=False)
    junction_name = db.Column(db.String(150), nullable=False)
    remark = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))

    employee = db.relationship('Employee', backref='junction_remarks')


def _team_of(employee):
    return (employee.team or '').lower() if employee and employee.team else ''


# ---- Team label helpers ----
TEAM_LABELS = {
    'field': 'Junction',
    'coc':   'COC Location',
    'ccc':   'CCC Location',
    'towing': 'Junction',
}

ALLOWED_VISIT_TEAMS = ('field', 'coc', 'ccc', 'towing')


@app.route('/api/locations', methods=['GET'])
def get_locations():
    """Returns the list of predefined locations for a specific team.
    Query param: ?team=field|coc|ccc|towing
    Falls back to JunctionList (legacy) for field & towing teams if LocationList is empty."""
    team = (request.args.get('team') or 'field').lower()

    if team not in ALLOWED_VISIT_TEAMS:
        return jsonify({'error': 'Invalid team. Must be field, coc, ccc, or towing'}), 400

    if team in ('field', 'towing'):
        locations = JunctionList.query.order_by(JunctionList.name.asc()).all()
        return jsonify([{'id': loc.id, 'name': loc.name, 'ward': loc.ward or '', 'zone': loc.zone or ''} for loc in locations])

    locations = LocationList.query.filter_by(team=team).order_by(LocationList.name.asc()).all()
    return jsonify([{'id': loc.id, 'name': loc.name, 'ward': loc.ward or '', 'zone': loc.zone or ''} for loc in locations])


@app.route('/api/admin/locations/upload', methods=['POST'])
def upload_locations_excel():
    """Admin uploads an Excel file to register employees (and locations for COC/CCC).
    Form data: team=field|coc|ccc, file=<excel/csv>
    Field team: requires 'Employee Name' column only.
    COC/CCC team: requires 'Employee Name' and 'Location Name' columns."""
    team = (request.form.get('team') or request.args.get('team') or '').lower()

    if team not in ALLOWED_VISIT_TEAMS:
        return jsonify({'error': 'Invalid team. Must be field, coc, or ccc'}), 400

    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file or not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls') or file.filename.endswith('.csv')):
        return jsonify({'error': 'Invalid file format. Please upload an Excel (.xlsx/.xls) or CSV file.'}), 400

    try:
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        # Normalize column names (strip spaces)
        df.columns = [c.strip() for c in df.columns]

        # Find the Employee Name column
        emp_col = None
        for candidate in ['Employee Name', 'employee name', 'Employee name', 'employee Name', 'Name', 'name', 'EMPLOYEE NAME']:
            if candidate in df.columns:
                emp_col = candidate
                break

        if not emp_col:
            return jsonify({'error': 'Missing column. File must have an "Employee Name" column. Download the sample Excel for reference.'}), 400

        # For COC/CCC teams, Location Name is also required
        loc_col = None
        if team in ('coc', 'ccc'):
            for candidate in ['Location Name', 'location name', 'Location name', 'Location ID', 'location id', 'LOCATION NAME']:
                if candidate in df.columns:
                    loc_col = candidate
                    break
            if not loc_col:
                return jsonify({'error': f'For {team.upper()} team, file must have "Employee Name" and "Location Name" columns. Download the sample Excel for reference.'}), 400

        # Find Ward & Zone columns if present
        col_clean_map_loc = {str(c).strip().lower(): c for c in df.columns}
        ward_col_loc = None
        for w_cand in ['ward', 'ward name', 'ward_name', 'ward no', 'ward_no', 'wards']:
            if w_cand in col_clean_map_loc:
                ward_col_loc = col_clean_map_loc[w_cand]
                break
        
        zone_col_loc = None
        for z_cand in ['zone', 'zone name', 'zone_name', 'zone no', 'zone_no', 'zones']:
            if z_cand in col_clean_map_loc:
                zone_col_loc = col_clean_map_loc[z_cand]
                break

        added_employees = 0
        added_locations = 0
        skipped = 0

        for _, row in df.iterrows():
            emp_name = str(row.get(emp_col, '')).strip()
            if not emp_name:
                continue

            w_val = str(row.get(ward_col_loc, '')).strip() if ward_col_loc else None
            z_val = str(row.get(zone_col_loc, '')).strip() if zone_col_loc else None
            if w_val and w_val.lower() == 'nan': w_val = None
            if z_val and z_val.lower() == 'nan': z_val = None

            if team == 'field':
                # Field team: register employee name as location entry
                exists = LocationList.query.filter_by(name=emp_name, team=team).first()
                if not exists:
                    new_entry = LocationList(name=emp_name, team=team, ward=w_val, zone=z_val)
                    db.session.add(new_entry)
                    added_employees += 1
                else:
                    if w_val: exists.ward = w_val
                    if z_val: exists.zone = z_val
                    skipped += 1
            else:
                # COC/CCC: get location name and register both
                loc_name = str(row.get(loc_col, '')).strip()
                if not loc_name:
                    skipped += 1
                    continue

                # Create location if not exists
                loc = LocationList.query.filter_by(name=loc_name, team=team).first()
                if not loc:
                    loc = LocationList(name=loc_name, team=team, ward=w_val, zone=z_val)
                    db.session.add(loc)
                    added_locations += 1
                else:
                    if w_val: loc.ward = w_val
                    if z_val: loc.zone = z_val

                # Store employee-location mapping
                emp_loc_key = f"{emp_name} @ {loc_name}"
                emp_exists = LocationList.query.filter_by(name=emp_loc_key, team=f"{team}_emp").first()
                if not emp_exists:
                    emp_entry = LocationList(name=emp_loc_key, team=f"{team}_emp", ward=w_val, zone=z_val)
                    db.session.add(emp_entry)
                    added_employees += 1
                else:
                    if w_val: emp_exists.ward = w_val
                    if z_val: emp_exists.zone = z_val
                    skipped += 1

        db.session.commit()
        parts = []
        if added_employees:
            parts.append(f'{added_employees} employee(s)')
        if added_locations:
            parts.append(f'{added_locations} location(s)')
        if skipped:
            parts.append(f'{skipped} duplicate(s) skipped')

        msg = f'Upload successful: {", ".join(parts)}.' if parts else 'No new data to add.'
        return jsonify({'success': True, 'message': msg})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': f'Failed to process file: {str(e)}'}), 500


@app.route('/api/admin/employees/sample', methods=['GET'])
def download_employees_sample_excel():
    """Download a sample Excel file for Employee upload."""
    data = {
        'Employee ID': ['EMP001', 'EMP002', 'EMP003'],
        'Full Name': ['Snehal Patil', 'Amit Sharma', 'Priya Nair'],
        'Email': ['snehal@example.com', 'amit@example.com', 'priya@example.com'],
        'Phone': ['9876543210', '9876543211', '9876543212'],
        'Password': ['pass123', 'pass456', 'pass789'],
        'Team': ['Field Team', 'CoC', 'CCC'],
        'Category': ['Smart City', 'IITMS', 'Towing']
    }
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Employees')
    output.seek(0)
    
    filename = 'sample_employees_upload.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/admin/employees/upload', methods=['POST'])
def upload_employees_excel():
    """Upload an Excel file to register or update employees."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file or not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls') or file.filename.endswith('.csv')):
        return jsonify({'error': 'Invalid file format. Please upload an Excel (.xlsx/.xls) or CSV file.'}), 400

    try:
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file, dtype=str)
        else:
            df = pd.read_excel(file, dtype=str)

        # Normalize column names (strip spaces)
        df.columns = [c.strip() for c in df.columns]

        # Check for Team/Designation column
        desig_col = None
        for candidate in ['Team', 'team', 'TEAM', 'Designation', 'designation', 'DESIGNATION']:
            if candidate in df.columns:
                desig_col = candidate
                break
        if not desig_col:
            return jsonify({'error': 'Missing Designation/Team column. File must have a "Team" or "Designation" column.'}), 400

        # Check required columns
        required_cols = ['Employee ID', 'Full Name', 'Email', 'Phone', 'Password']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            return jsonify({'error': f'Missing columns: {", ".join(missing_cols)}'}), 400

        # Check for Category/Group column
        group_col = None
        for candidate in ['Category', 'category', 'Group', 'group', 'GROUP', 'CATEGORY']:
            if candidate in df.columns:
                group_col = candidate
                break
        if not group_col:
            return jsonify({'error': 'Missing Category/Group column. File must have a "Category" or "Group" column.'}), 400

        def clean_val(val):
            if pd.isna(val):
                return ''
            s = str(val).strip()
            if s.endswith('.0'):
                s = s[:-2]
            return s

        # Process spreadsheet rows to add or update employees

        added_count = 0
        updated_count = 0

        for index, row in df.iterrows():
            emp_id = clean_val(row['Employee ID'])
            if not emp_id or emp_id.lower() == 'nan':
                continue # Skip empty rows

            if emp_id.isdigit():
                emp_id = f"EMP{int(emp_id):03d}"
            else:
                emp_id = emp_id.upper()

            full_name = clean_val(row['Full Name'])
            email = clean_val(row['Email']).lower()
            phone = clean_val(row['Phone'])
            password = clean_val(row['Password'])
            designation = clean_val(row[desig_col])
            category_val = clean_val(row[group_col])

            if not full_name or full_name.lower() == 'nan':
                continue # Skip empty rows

            # Map designation to team
            team_mapped = None
            desig_lower = designation.lower()
            if 'field' in desig_lower:
                team_mapped = 'field'
            elif 'coc' in desig_lower or 'cooc' in desig_lower:
                team_mapped = 'coc'
            elif 'ccc' in desig_lower:
                team_mapped = 'ccc'
            elif 'towing' in desig_lower:
                team_mapped = 'towing'

            # Normalize category value
            category_val_lower = category_val.lower()
            if 'smart' in category_val_lower and 'city' in category_val_lower:
                category_mapped = 'Smart City'
            elif 'itms' in category_val_lower:
                category_mapped = 'IITMS'
            elif 'towing' in category_val_lower or 'construction' in category_val_lower:
                category_mapped = 'Towing'
            else:
                category_mapped = category_val if category_val else None

            # Check if email is used by another employee
            existing_email_emp = Employee.query.filter_by(email=email).first()
            if existing_email_emp and existing_email_emp.id.strip().upper() != emp_id:
                return jsonify({'error': f'Email {email} is already in use by employee {existing_email_emp.id}'}), 400

            employee = Employee.query.get(emp_id)
            if employee:
                employee.full_name = full_name
                employee.email = email
                employee.phone = phone
                employee.password = password
                employee.designation = designation
                employee.team = team_mapped
                employee.category = category_mapped
                if emp_id.upper().startswith('ADMIN'):
                    employee.is_admin = True
                updated_count += 1
            else:
                employee = Employee(
                    id=emp_id,
                    full_name=full_name,
                    email=email,
                    phone=phone,
                    password=password,
                    designation=designation,
                    team=team_mapped,
                    category=category_mapped,
                    is_admin=emp_id.upper().startswith('ADMIN'),
                    weekly_off='Sunday',
                    shift_type='general'
                )
                db.session.add(employee)
                added_count += 1

        # Seed default admin and category admins if they don't exist
        for adm in [
            {'id': 'ADMIN001', 'full_name': 'System Administrator', 'email': 'multisulotionsdecofurn@gmail.com', 'phone': '+919518791736', 'password': 'AD#987', 'category': None},
            {'id': 'smartcity_admin', 'full_name': 'Smart City Admin', 'email': 'smartcity@keltron.com', 'phone': '0000000001', 'password': 'SmartCity@Admin', 'category': 'Smart City'},
            {'id': 'iitms_admin', 'full_name': 'IITMS Admin', 'email': 'iitms@keltron.com', 'phone': '0000000002', 'password': 'IITMS@Admin', 'category': 'IITMS'},
            {'id': 'towing_admin', 'full_name': 'Towing Admin', 'email': 'towing@keltron.com', 'phone': '0000000003', 'password': 'Towing@Admin', 'category': 'Towing'}
        ]:
            if not Employee.query.get(adm['id']):
                new_adm = Employee(
                    id=adm['id'],
                    full_name=adm['full_name'],
                    email=adm['email'],
                    phone=adm['phone'],
                    password=os.getenv(f"{adm['id']}_PASSWORD", adm['password']),
                    is_admin=True,
                    category=adm['category'],
                    weekly_off='Sunday',
                    shift_type='general'
                )
                db.session.add(new_adm)
                added_count += 1

        db.session.commit()

        return jsonify({
            'success': True,
            'message': f'Uploaded successfully. Added {added_count} new/admin employees and updated {updated_count} existing employees.'
        })

    except Exception as e:
        db.session.rollback()
        return jsonify({'error': f'An error occurred during file parsing: {str(e)}'}), 500


@app.route('/api/admin/locations/sample', methods=['GET'])
def download_sample_excel():
    """Download a sample Excel file for the specified team.
    Query param: ?team=field|coc|ccc"""
    team = (request.args.get('team') or '').lower()
    if team not in ALLOWED_VISIT_TEAMS:
        return jsonify({'error': 'Invalid team'}), 400

    if team == 'field':
        data = {'Employee Name': ['John Doe', 'Jane Smith', 'Amit Kumar']}
    else:
        data = {
            'Employee Name': ['Rahul Sharma', 'Priya Patel', 'Suresh Gupta'],
            'Location Name': ['Location A', 'Location B', 'Location C'],
        }

    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Employees')
    output.seek(0)

    filename = f'sample_{team}_upload.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/admin/locations/<team>', methods=['GET'])
def get_all_locations_for_team(team):
    """Admin: get full location list for a team."""
    team = team.lower()
    if team not in ALLOWED_VISIT_TEAMS:
        return jsonify({'error': 'Invalid team'}), 400
    locations = LocationList.query.filter_by(team=team).order_by(LocationList.name.asc()).all()
    return jsonify([{'id': loc.id, 'name': loc.name, 'team': loc.team, 'ward': loc.ward or '', 'zone': loc.zone or ''} for loc in locations])


@app.route('/api/admin/locations/<int:location_id>', methods=['DELETE'])
def delete_location(location_id):
    """Admin: delete a specific location entry."""
    loc = LocationList.query.get(location_id)
    if not loc:
        return jsonify({'error': 'Location not found'}), 404
    db.session.delete(loc)
    db.session.commit()
    return jsonify({'success': True, 'message': 'Location deleted'})


@app.route('/api/admin/locations/add', methods=['POST'])
def add_single_location():
    """Admin: manually add a single location entry (used by AddJunctionModal).
    JSON body: { name, team, ward?, zone?, parent_location_id? }"""
    data = request.get_json()
    name = (data.get('name') or '').strip()
    team = (data.get('team') or '').lower()
    ward = (data.get('ward') or '').strip()
    zone = (data.get('zone') or '').strip()

    if not name:
        return jsonify({'error': 'Location name is required'}), 400
    if team not in ALLOWED_VISIT_TEAMS:
        return jsonify({'error': 'Invalid team. Must be field, coc, or ccc'}), 400

    existing = LocationList.query.filter_by(name=name, team=team).first()
    if existing:
        return jsonify({'error': f'Location "{name}" already exists for {team} team'}), 400

    loc = LocationList(name=name, team=team, ward=ward or None, zone=zone or None)
    db.session.add(loc)
    db.session.commit()

    return jsonify({
        'success': True,
        'message': f'Location "{name}" added successfully',
        'location': {'id': loc.id, 'name': loc.name, 'team': loc.team, 'ward': loc.ward or '', 'zone': loc.zone or ''}
    }), 201


@app.route('/api/admin/locations/<int:location_id>/assign', methods=['POST'])
def assign_employees_to_location(location_id):
    """Admin: save a list of employee IDs assigned to a location.
    JSON body: { employee_ids: [str, ...] }
    NOTE: This stores the assignment in session; for a full persistent
    implementation a many-to-many join table would be needed.
    Currently returns 200 with a confirmation so the frontend modal works."""
    loc = LocationList.query.get(location_id)
    if not loc:
        return jsonify({'error': 'Location not found'}), 404

    data = request.get_json()
    employee_ids = data.get('employee_ids', [])

    # Validate all employee IDs exist
    valid = []
    for eid in employee_ids:
        emp = Employee.query.get(str(eid))
        if emp:
            valid.append(eid)

    return jsonify({
        'success': True,
        'message': f'Assignment saved for {len(valid)} employee(s) to "{loc.name}"',
        'location_id': location_id,
        'assigned_employee_ids': valid
    })


@app.route('/api/junction/start', methods=['POST'])
def junction_start():
    """Employee arrives at a junction/location: log the 'before' photo + GPS location.
    Available for Field, COC, and CCC team members."""
    employee_id = request.form.get('employee_id')
    junction_name = request.form.get('junction_name')
    location = request.form.get('location')
    ward = request.form.get('ward')
    zone = request.form.get('zone')
    before_remark = request.form.get('before_remark')

    if not employee_id or not junction_name or not location:
        return jsonify({'error': 'employee_id, junction_name and location are required'}), 400

    employee = Employee.query.get(employee_id)
    if not employee:
        return jsonify({'error': 'Employee not found'}), 404

    # Allow all employees to log junction visits (team restriction removed)
    # Previously: if _team_of(employee) not in ALLOWED_VISIT_TEAMS:
    #                 return jsonify({'error': 'Location visits are only available for Field, COC, and CCC Team members'}), 403

    today = datetime.utcnow().date()

    checked_in_today = Attendance.query.filter(
        Attendance.employee_id == employee_id,
        Attendance.date == today,
        Attendance.check_in_time.isnot(None)
    ).first()
    if not checked_in_today:
        return jsonify({'error': 'Please check in for the day before logging a junction visit'}), 400

    # Auto-close any existing open visits ('in_progress') for this employee
    open_visits = JunctionVisit.query.filter_by(
        employee_id=employee_id,
        status='in_progress'
    ).all()
    for ov in open_visits:
        ov.status = 'unresolved'
        ov.completed_at = datetime.utcnow()
        ov.remark = 'Unresolved'

    # REMOVED: Restriction that blocks a second open visit so employee can visit other locations
    # if one call is not completed.
    
    visit_type = request.form.get('visit_type', 'Regular Visit')
    asset_type = request.form.get('asset_type')
    fault_type = request.form.get('fault_type')

    before_photo_filename = None
    if 'photo' in request.files:
        file = request.files['photo']
        if file and allowed_file(file.filename):
            filename = secure_filename(
                f"{employee_id}{datetime.now().strftime('%Y%m%d%H%M%S')}_junction_before.{file.filename.rsplit('.', 1)[1].lower()}"
            )
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            before_photo_filename = filename

    if not before_photo_filename:
        return jsonify({'error': 'Before photo is required'}), 400

    is_completed_immediately = (visit_type == 'Regular Visit' and asset_type == 'None')

    now_utc = datetime.utcnow()
    last_visit = JunctionVisit.query.filter_by(employee_id=employee_id, date=today).order_by(JunctionVisit.started_at.desc()).first()
    prev_finish = (last_visit.completed_at or last_visit.started_at) if last_visit else (checked_in_today.check_in_time if checked_in_today else None)

    calc_travel_mins = 0.0
    if prev_finish and now_utc > prev_finish:
        calc_travel_mins = round((now_utc - prev_finish).total_seconds() / 60.0, 1)

    visit = JunctionVisit(
        employee_id=employee_id,
        junction_name=junction_name.strip(),
        ward=ward.strip() if ward else None,
        zone=zone.strip() if zone else None,
        date=today,
        before_photo=before_photo_filename,
        before_location=location,
        started_at=now_utc,
        status='completed' if is_completed_immediately else 'in_progress',
        completed_at=now_utc if is_completed_immediately else None,
        after_photo=before_photo_filename if is_completed_immediately else None,
        after_location=location if is_completed_immediately else None,
        remark='Regular Visit (No Fault)' if is_completed_immediately else None,
        before_remark=before_remark.strip() if before_remark else None,
        visit_type=visit_type,
        asset_type=asset_type,
        fault_type=fault_type,
        travel_time_minutes=calc_travel_mins,
        time_spent_minutes=0.0 if is_completed_immediately else 0.0
    )
    db.session.add(visit)
    db.session.commit()


    return jsonify({
        'success': True,
        'message': 'Junction visit completed' if is_completed_immediately else 'Junction visit started',
        'visit': {
            'id': visit.id,
            'junction_name': visit.junction_name,
            'ward': visit.ward or '',
            'zone': visit.zone or '',
            'started_at': to_ist(visit.started_at),
            'before_photo': visit.before_photo,
            'before_location': visit.before_location,
            'status': visit.status,
            'visit_type': visit.visit_type,
            'asset_type': visit.asset_type or '',
            'fault_type': visit.fault_type or '',
            'remark': visit.remark
        }
    })


@app.route('/api/junction/<int:visit_id>/complete', methods=['POST'])
def junction_complete(visit_id):
    """Employee finishes a junction visit: log the 'after' photo + GPS location."""
    location = request.form.get('location')
    remark = request.form.get('remark')
    
    if not location:
        return jsonify({'error': 'Location is required'}), 400
    if not remark or not remark.strip():
        return jsonify({'error': 'Remark is compulsory'}), 400
    visit = JunctionVisit.query.get(visit_id)
    if not visit:
        return jsonify({'error': 'Junction visit not found'}), 404

    if visit.status == 'completed':
        return jsonify({'error': 'This junction visit is already completed'}), 400

    after_photo_filename = None
    if 'photo' in request.files:
        file = request.files['photo']
        if file and allowed_file(file.filename):
            filename = secure_filename(
                f"{visit.employee_id}{datetime.now().strftime('%Y%m%d%H%M%S')}_junction_after.{file.filename.rsplit('.', 1)[1].lower()}"
            )
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            after_photo_filename = filename

    if not after_photo_filename:
        return jsonify({'error': 'After photo is required'}), 400

    visit.after_photo = after_photo_filename
    visit.after_location = location
    visit.remark = remark.strip()
    if request.form.get('asset_type'):
        visit.asset_type = request.form.get('asset_type')
    if request.form.get('fault_type'):
        visit.fault_type = request.form.get('fault_type')
    visit.completed_at = datetime.utcnow()
    visit.status = 'completed'
    if visit.started_at:
        visit.time_spent_minutes = round((visit.completed_at - visit.started_at).total_seconds() / 60.0, 1)
    db.session.commit()


    return jsonify({
        'success': True,
        'message': 'Junction visit completed',
        'visit': {
            'id': visit.id,
            'after_photo': visit.after_photo,
            'after_location': visit.after_location,
            'completed_at': to_ist(visit.completed_at),
            'status': visit.status,
            'visit_type': visit.visit_type,
            'asset_type': visit.asset_type or '',
            'fault_type': visit.fault_type or '',
            'remark': visit.remark
        }
    })


@app.route('/api/junction/<int:visit_id>/call-visit/start', methods=['POST'])
def call_visit_start(visit_id):
    """Starts a call visit for an unresolved junction visit."""
    employee_id = request.form.get('employee_id')
    location = request.form.get('location')
    before_remark = request.form.get('before_remark')
    
    if not employee_id or not location:
        return jsonify({'error': 'employee_id and location are required'}), 400

    employee = Employee.query.get(employee_id)
    if not employee:
        return jsonify({'error': 'Employee not found'}), 404

    # Check if checked in today
    today = datetime.utcnow().date()
    checked_in_today = Attendance.query.filter(
        Attendance.employee_id == employee_id,
        Attendance.date == today,
        Attendance.check_in_time.isnot(None)
    ).first()
    if not checked_in_today:
        return jsonify({'error': 'Please check in for the day before logging a call visit'}), 400

    # Auto-close any in_progress Call Visits for this employee
    open_cvs = CallVisit.query.join(JunctionVisit).filter(
        JunctionVisit.employee_id == employee_id,
        CallVisit.status == 'in_progress'
    ).all()
    for ocv in open_cvs:
        ocv.status = 'unresolved'
        ocv.completed_at = datetime.utcnow()
        ocv.remark = 'Unresolved'

    # Also auto-close any regular in-progress junction visits for consistency
    open_visits = JunctionVisit.query.filter_by(
        employee_id=employee_id,
        status='in_progress'
    ).all()
    for ov in open_visits:
        ov.status = 'unresolved'
        ov.completed_at = datetime.utcnow()
        ov.remark = 'Unresolved'

    photo_filename = None
    if 'photo' in request.files:
        file = request.files['photo']
        if file and allowed_file(file.filename):
            filename = secure_filename(
                f"{employee_id}{datetime.now().strftime('%Y%m%d%H%M%S')}_call_before.{file.filename.rsplit('.', 1)[1].lower()}"
            )
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            photo_filename = filename

    if not photo_filename:
        return jsonify({'error': 'Before photo is required'}), 400

    new_cv = CallVisit(
        junction_visit_id=visit_id,
        before_photo=photo_filename,
        before_location=location,
        started_at=datetime.utcnow(),
        status='in_progress',
        before_remark=before_remark.strip() if before_remark else None
    )
    db.session.add(new_cv)
    db.session.commit()

    return jsonify({
        'success': True,
        'message': 'Call visit started',
        'call_visit': {
            'id': new_cv.id,
            'junction_visit_id': new_cv.junction_visit_id,
            'started_at': to_ist(new_cv.started_at),
            'before_photo': new_cv.before_photo,
            'before_location': new_cv.before_location,
            'before_remark': new_cv.before_remark or '',
            'status': new_cv.status
        }
    })


@app.route('/api/junction/<int:visit_id>/call-visit/complete', methods=['POST'])
def call_visit_complete(visit_id):
    """Completes the active call visit for this junction visit."""
    location = request.form.get('location')
    remark = request.form.get('remark')
    
    if not location:
        return jsonify({'error': 'Location is required'}), 400
    if not remark or not remark.strip():
        return jsonify({'error': 'Remark is compulsory'}), 400

    visit = JunctionVisit.query.get(visit_id)
    if not visit:
        return jsonify({'error': 'Junction visit not found'}), 404

    active_cv = CallVisit.query.filter_by(
        junction_visit_id=visit_id,
        status='in_progress'
    ).first()
    if not active_cv:
        return jsonify({'error': 'No active call visit found for this junction visit'}), 400

    after_photo_filename = None
    if 'photo' in request.files:
        file = request.files['photo']
        if file and allowed_file(file.filename):
            filename = secure_filename(
                f"{visit.employee_id}{datetime.now().strftime('%Y%m%d%H%M%S')}_call_after.{file.filename.rsplit('.', 1)[1].lower()}"
            )
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            after_photo_filename = filename

    if not after_photo_filename:
        return jsonify({'error': 'After photo is required'}), 400

    # Update active call visit
    active_cv.after_photo = after_photo_filename
    active_cv.after_location = location
    active_cv.remark = remark.strip()
    active_cv.completed_at = datetime.utcnow()
    active_cv.status = 'completed'

    # Update parent junction visit status and photos
    visit.after_photo = after_photo_filename
    visit.after_location = location
    visit.completed_at = active_cv.completed_at
    visit.remark = remark.strip()
    visit.status = 'completed'

    db.session.commit()

    return jsonify({
        'success': True,
        'message': 'Call visit completed and junction visit marked as complete',
        'call_visit': {
            'id': active_cv.id,
            'junction_visit_id': active_cv.junction_visit_id,
            'after_photo': active_cv.after_photo,
            'after_location': active_cv.after_location,
            'completed_at': to_ist(active_cv.completed_at),
            'status': active_cv.status,
            'remark': active_cv.remark
        }
    })


@app.route('/api/junction/today/<employee_id>', methods=['GET'])
def junction_today(employee_id):
    try:
        today = datetime.utcnow().date()

        visits = JunctionVisit.query.filter(
            JunctionVisit.employee_id == employee_id,
            or_(
                JunctionVisit.date == today,
                JunctionVisit.status == 'in_progress',
                JunctionVisit.status == 'unresolved'
            )
        ).order_by(JunctionVisit.id.desc()).all()

        serialized = []
        for v in visits:
            active_cv = CallVisit.query.filter_by(junction_visit_id=v.id, status='in_progress').first()
            serialized.append({
                'id': v.id,
                'junction_name': v.junction_name,
                'ward': v.ward or '',
                'zone': v.zone or '',
                'before_photo': v.before_photo,
                'after_photo': v.after_photo,
                'before_location': v.before_location,
                'after_location': v.after_location,
                'started_at': to_ist(v.started_at),
                'completed_at': to_ist(v.completed_at),
                'status': v.status,
                'visit_type': v.visit_type,
                'asset_type': v.asset_type or '',
                'fault_type': v.fault_type or '',
                'remark': v.remark,
                'before_remark': v.before_remark or '',
                'time_spent_minutes': getattr(v, 'time_spent_minutes', 0.0) or (round((v.completed_at - v.started_at).total_seconds() / 60.0, 1) if (v.started_at and v.completed_at) else 0.0),
                'travel_time_minutes': getattr(v, 'travel_time_minutes', 0.0) or 0.0,
                'active_call_visit': {
                    'id': active_cv.id,
                    'before_photo': active_cv.before_photo,
                    'before_location': active_cv.before_location,
                    'started_at': to_ist(active_cv.started_at),
                    'before_remark': active_cv.before_remark or '',
                    'status': active_cv.status
                } if active_cv else None
            })

        return jsonify(serialized)

    except Exception as e:
        print("JUNCTION TODAY ERROR:", e)
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/location/ping', methods=['POST'])
def receive_location_ping():
    """Receives location pings from mobile/web clients during active official shift hours."""
    try:
        data = request.get_json(silent=True) or request.form
        employee_id = data.get('employee_id')
        lat = data.get('latitude') or data.get('lat')
        lng = data.get('longitude') or data.get('lng')

        if not employee_id or lat is None or lng is None:
            return jsonify({'error': 'employee_id, latitude, longitude are required'}), 400

        employee = Employee.query.get(employee_id)
        if not employee:
            return jsonify({'error': 'Employee not found'}), 404

        today = datetime.utcnow().date()
        attendance = Attendance.query.filter(
            Attendance.employee_id == employee_id,
            Attendance.date == today,
            Attendance.check_in_time.isnot(None)
        ).first()

        if not attendance or attendance.check_out_time:
            return jsonify({'success': False, 'message': 'Not in active shift'}), 200

        ping = LocationPing(
            employee_id=employee_id,
            latitude=float(lat),
            longitude=float(lng),
            accuracy=float(data.get('accuracy')) if data.get('accuracy') else None,
            speed=float(data.get('speed')) if data.get('speed') else None,
            battery_level=float(data.get('battery_level')) if data.get('battery_level') else None,
            timestamp=datetime.utcnow()
        )
        db.session.add(ping)
        db.session.commit()

        return jsonify({'success': True, 'message': 'Location ping logged'})
    except Exception as e:
        print("LOCATION PING ERROR:", e)
        return jsonify({'error': str(e)}), 500


@app.route('/api/employee/<employee_id>/timeline/today', methods=['GET'])
def get_employee_today_timeline(employee_id):
    """Returns today's movement timeline including visits, travel durations, and location pings."""
    try:
        today = datetime.utcnow().date()
        employee = Employee.query.get(employee_id)
        if not employee:
            return jsonify({'error': 'Employee not found'}), 404

        attendance = Attendance.query.filter_by(employee_id=employee_id, date=today).first()
        check_in_ist = to_ist(attendance.check_in_time) if attendance and attendance.check_in_time else None
        check_out_ist = to_ist(attendance.check_out_time) if attendance and attendance.check_out_time else None

        visits = JunctionVisit.query.filter_by(employee_id=employee_id, date=today).order_by(JunctionVisit.started_at.asc()).all()

        visit_list = []
        prev_finish_time = attendance.check_in_time if attendance else None

        for v in visits:
            arrival_ist = to_ist(v.started_at) if v.started_at else None
            completion_ist = to_ist(v.completed_at) if v.completed_at else None

            time_spent_mins = getattr(v, 'time_spent_minutes', 0.0)
            if not time_spent_mins or time_spent_mins == 0:
                if v.started_at and v.completed_at:
                    time_spent_mins = round((v.completed_at - v.started_at).total_seconds() / 60.0, 1)
                elif v.started_at:
                    time_spent_mins = round((datetime.utcnow() - v.started_at).total_seconds() / 60.0, 1)

            travel_mins = getattr(v, 'travel_time_minutes', 0.0)
            if not travel_mins or travel_mins == 0:
                if prev_finish_time and v.started_at and v.started_at > prev_finish_time:
                    travel_mins = round((v.started_at - prev_finish_time).total_seconds() / 60.0, 1)

            if v.completed_at:
                prev_finish_time = v.completed_at
            elif v.started_at:
                prev_finish_time = v.started_at

            visit_list.append({
                'id': v.id,
                'junction_name': v.junction_name,
                'ward': v.ward or '',
                'zone': v.zone or '',
                'visit_type': v.visit_type or 'Regular Visit',
                'status': v.status,
                'arrival_time': arrival_ist,
                'completion_time': completion_ist,
                'time_spent_minutes': time_spent_mins,
                'travel_time_minutes': travel_mins,
                'remark': v.remark or '',
                'asset_type': v.asset_type or '',
                'fault_type': v.fault_type or ''
            })

        pings = LocationPing.query.filter(
            LocationPing.employee_id == employee_id,
            db.func.date(LocationPing.timestamp) == today
        ).order_by(LocationPing.timestamp.asc()).all()

        ping_list = [{
            'lat': p.latitude,
            'lng': p.longitude,
            'time': to_ist(p.timestamp) if p.timestamp else None,
            'accuracy': p.accuracy
        } for p in pings]

        return jsonify({
            'employee_id': employee_id,
            'employee_name': employee.full_name,
            'date': today.strftime('%Y-%m-%d'),
            'check_in_time': check_in_ist,
            'check_out_time': check_out_ist,
            'total_visits': len(visits),
            'visits': visit_list,
            'location_pings_count': len(pings),
            'pings': ping_list
        })
    except Exception as e:
        print("TIMELINE ERROR:", e)
        return jsonify({'error': str(e)}), 500



@app.route('/api/junctions', methods=['GET'])
def get_junctions():
    """Returns the list of predefined junctions for the dropdown."""
    junctions = JunctionList.query.order_by(JunctionList.name.asc()).all()
    return jsonify([{'id': j.id, 'name': j.name, 'ward': j.ward or '', 'zone': j.zone or ''} for j in junctions])

@app.route('/api/admin/asset-faults/upload', methods=['POST'])
def upload_asset_faults_excel():
    """Admin uploads an Excel file with 'Asset Type' and 'Fault Type' columns to populate the AssetFaultMapping table."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if not file or not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls') or file.filename.endswith('.csv')):
        return jsonify({'error': 'Invalid file format. Please upload an Excel or CSV file.'}), 400
    
    try:
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file, dtype=str, keep_default_na=False)
        else:
            df = pd.read_excel(file, dtype=str, keep_default_na=False)
            
        # Normalize columns
        df.columns = [str(c).strip() for c in df.columns]
        
        # Find 'Asset Type' column (support aliases like Asset, asset type, ASSET TYPE)
        asset_col = None
        for candidate in ['Asset Type', 'asset type', 'Asset', 'asset', 'ASSET', 'ASSET TYPE']:
            if candidate in df.columns:
                asset_col = candidate
                break
                
        # Find 'Fault Type' column (support aliases like Fault, fault type, FAULT TYPE)
        fault_col = None
        for candidate in ['Fault Type', 'fault type', 'Fault', 'fault', 'FAULT', 'FAULT TYPE']:
            if candidate in df.columns:
                fault_col = candidate
                break

        if not asset_col or not fault_col:
            return jsonify({'error': 'Missing "Asset Type" or "Fault Type" columns in the uploaded file.'}), 400
            
        def clean_val(val):
            if pd.isna(val):
                return ''
            s = str(val).strip()
            if s.endswith('.0'):
                s = s[:-2]
            return s

        # Clear existing mappings
        db.session.query(AssetFaultMapping).delete()
        
        added_count = 0
        inserted_set = set() # Avoid duplicates in the excel file itself
        
        for _, row in df.iterrows():
            asset_type = clean_val(row[asset_col])
            fault_type = clean_val(row[fault_col])
            if not asset_type or not fault_type or asset_type.lower() == 'nan' or fault_type.lower() == 'nan':
                continue
                
            pair = (asset_type, fault_type)
            if pair not in inserted_set:
                new_mapping = AssetFaultMapping(asset_type=asset_type, fault_type=fault_type)
                db.session.add(new_mapping)
                inserted_set.add(pair)
                added_count += 1
                 
        db.session.commit()
        return jsonify({'success': True, 'message': f'Successfully processed asset & fault mapping. Added {added_count} mapping entries.'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': f'Failed to process file: {str(e)}'}), 500


@app.route('/api/admin/asset-faults/sample', methods=['GET'])
def download_asset_faults_sample_excel():
    """Download a sample Excel file for Asset and Fault Mapping upload."""
    data = {
        'Asset Type': ['Fiber', 'Fiber', 'Fixed Camera', 'Fixed Camera', 'PTZ Camera', 'JB', 'None'],
        'Fault Type': ['Damage', 'Other (Add as required)', 'Lens', 'Power Issue', 'Power Issue', 'Rectifier', 'None']
    }
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Asset Fault Mapping')
    output.seek(0)
    
    filename = 'sample_asset_fault_mapping.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/junction/asset-faults', methods=['GET'])
def get_asset_fault_mappings():
    """Returns the grouped dictionary of asset type to list of fault types."""
    mappings = AssetFaultMapping.query.order_by(AssetFaultMapping.asset_type.asc(), AssetFaultMapping.id.asc()).all()
    if not mappings:
        # Fall back to default hardcoded mappings if database is empty
        DEFAULT_MAPPING = {
            'Fiber': ['Damage', 'Other (Add as required)'],
            'Fixed Camera': ['Cat 6 Cable', 'Lens', 'Power Issue', 'Other (Add as required)'],
            'PTZ Camera': ['Cat 6 Cable', 'Lens', 'Power Issue', 'Other (Add as required)'],
            'MS Camera': ['Cat 6 Cable', 'Lens', 'Power Issue', 'Other (Add as required)'],
            'Dome Camera': ['Cat 6 Cable', 'Lens', 'Power Issue', 'Other (Add as required)'],
            'Wifi AP': ['Cat 6 Cable', 'Power Issue', 'Other (Add as required)'],
            'JB': ['Rectifier', 'Power Issue', 'Other (Add as required)'],
            'Kiosk': ['Screen Broken', 'Power Issue', 'Other (Add as required)'],
            'VaMS': ['Pixel Issue', 'Power Issue', 'Other (Add as required)'],
            'PA System': ['Speaker Issue', 'Power Issue', 'Other (Add as required)'],
            'None': ['None']
        }
        return jsonify(DEFAULT_MAPPING)
        
    grouped = {}
    for m in mappings:
        a_type = m.asset_type
        f_type = m.fault_type
        if a_type not in grouped:
            grouped[a_type] = []
        if f_type not in grouped[a_type]:
            grouped[a_type].append(f_type)
            
    return jsonify(grouped)


@app.route('/api/admin/junctions/upload', methods=['POST'])
def upload_junctions_excel():
    """Admin uploads an Excel file with 'Junction Name', 'Ward', and 'Zone' columns to populate the JunctionList table."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if not file or not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls') or file.filename.endswith('.csv')):
        return jsonify({'error': 'Invalid file format. Please upload an Excel or CSV file.'}), 400
    
    try:
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file, dtype=str)
        else:
            xl = pd.ExcelFile(file)
            sheet_names = xl.sheet_names
            
            # Prioritize sheets containing "junction" or "location" key words
            prioritized_sheets = []
            other_sheets = []
            for sheet in sheet_names:
                sheet_lower = sheet.lower()
                if 'junction' in sheet_lower or 'location' in sheet_lower:
                    prioritized_sheets.append(sheet)
                else:
                    other_sheets.append(sheet)
            search_order = prioritized_sheets + other_sheets
            
            # Find candidate sheet
            selected_sheet = sheet_names[0]
            candidates = [
                'Junction Name', 'junction name', 'Junction', 'junction', 
                'Location', 'location', 'LOCATION', 'JUNCTION', 'Name', 'name', 'NAME', 
                'Location / Junction', 'Location/Junction', 'LOCATION / JUNCTION', 
                'LOCATION/JUNCTION', 'Junction/Location', 'Junction / Location'
            ]
            
            for sheet in search_order:
                temp_df = pd.read_excel(file, sheet_name=sheet, nrows=5, dtype=str)
                has_candidate = False
                cols = [str(c).strip() for c in temp_df.columns]
                if any(cand in cols for cand in candidates):
                    has_candidate = True
                else:
                    for r_idx in range(len(temp_df)):
                        row_vals = [str(val).strip() for val in temp_df.iloc[r_idx]]
                        if any(cand in row_vals for cand in candidates):
                            has_candidate = True
                            break
                if has_candidate:
                    selected_sheet = sheet
                    break
            
            df = pd.read_excel(file, sheet_name=selected_sheet, dtype=str)
            
        # Check if the parsed columns are Unnamed (e.g. headers are not on row 0)
        is_unnamed = any('Unnamed' in str(c) for c in df.columns)
        candidates = [
            'Junction Name', 'junction name', 'Junction', 'junction', 
            'Location', 'location', 'LOCATION', 'JUNCTION', 'Name', 'name', 'NAME', 
            'Location / Junction', 'Location/Junction', 'LOCATION / JUNCTION', 
            'LOCATION/JUNCTION', 'Junction/Location', 'Junction / Location'
        ]
        if is_unnamed:
            for r_idx in range(min(5, len(df))):
                row_vals = [str(val).strip() for val in df.iloc[r_idx]]
                found = False
                for candidate in candidates:
                    if candidate in row_vals:
                        found = True
                        break
                if found:
                    df.columns = row_vals
                    df = df.iloc[r_idx + 1:].reset_index(drop=True)
                    break
            
        # Normalize columns map for case-insensitive matching
        col_clean_map = {str(c).strip().lower(): c for c in df.columns}
        
        def match_column(candidates, substrings=None):
            for cand in candidates:
                if cand.lower() in col_clean_map:
                    return col_clean_map[cand.lower()]
            if substrings:
                for k, orig in col_clean_map.items():
                    for sub in substrings:
                        if sub.lower() in k:
                            return orig
            return None

        # Find Junction Name column
        junction_candidates = [
            'junction name', 'junction', 'location name', 'location', 
            'site name', 'site', 'junction_name', 'location_name', 'name', 
            'location / junction', 'location/junction', 'junction/location', 'junction / location'
        ]
        junction_col = match_column(junction_candidates, ['junction', 'location', 'site'])

        if not junction_col:
            return jsonify({'error': 'Missing "Junction Name" (or "Location", "Junction") column in the uploaded file.'}), 400
            
        # Find Ward column
        ward_candidates = [
            'ward', 'ward name', 'ward_name', 'ward no', 'ward no.', 'ward_no', 
            'ward number', 'wards', 'from ward', 'to ward', 'from_ward', 'to_ward', 
            'ward/zone', 'ward code', 'ward_code', 'ward_id', 'ward id'
        ]
        ward_col = match_column(ward_candidates, ['ward'])

        # Find Zone column
        zone_candidates = [
            'zone', 'zone name', 'zone_name', 'zone no', 'zone no.', 'zone_no', 
            'zone number', 'zones', 'from zone', 'to zone', 'from_zone', 'to_zone', 
            'zone code', 'zone_code', 'zone_id', 'zone id'
        ]
        zone_col = match_column(zone_candidates, ['zone'])
                
        def clean_val(val):
            if pd.isna(val):
                return ''
            s = str(val).strip()
            if s.endswith('.0'):
                s = s[:-2]
            return s

        added_count = 0
        updated_count = 0

        for _, row in df.iterrows():
            name = clean_val(row[junction_col])
            if not name or name.lower() == 'nan':
                continue
            
            ward_val = clean_val(row[ward_col]) if ward_col else None
            zone_val = clean_val(row[zone_col]) if zone_col else None
            
            if not ward_val or ward_val.lower() == 'nan':
                ward_val = None
            if not zone_val or zone_val.lower() == 'nan':
                zone_val = None
                
            exists = JunctionList.query.filter_by(name=name).first()
            if not exists:
                new_j = JunctionList(name=name, ward=ward_val, zone=zone_val)
                db.session.add(new_j)
                added_count += 1
            else:
                if ward_val is not None:
                    exists.ward = ward_val
                if zone_val is not None:
                    exists.zone = zone_val
                updated_count += 1

            loc_exists = LocationList.query.filter_by(name=name, team='field').first()
            if not loc_exists:
                new_loc = LocationList(name=name, team='field', ward=ward_val, zone=zone_val)
                db.session.add(new_loc)
            else:
                if ward_val is not None:
                    loc_exists.ward = ward_val
                if zone_val is not None:
                    loc_exists.zone = zone_val
                 
        db.session.commit()
        return jsonify({'success': True, 'message': f'Successfully processed junctions. Added {added_count} new, updated {updated_count}.'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': f'Failed to process file: {str(e)}'}), 500

@app.route('/api/admin/junctions/sample', methods=['GET'])
def download_junctions_sample_excel():
    """Download a sample Excel file for Junction upload."""
    data = {
        'Junction Name': ['Junction A', 'Junction B', 'Junction C'],
        'Ward': ['Ward 1', 'Ward 2', 'Ward 3'],
        'Zone': ['Zone X', 'Zone Y', 'Zone Z']
    }
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Junctions')
    output.seek(0)
    
    filename = 'sample_junctions_upload.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/api/admin/team/<team>/summary', methods=['GET'])
def admin_team_summary(team):
    """Top-line stats for one team's admin dashboard tab (COC / CCC / Field / Towing)."""
    team = team.lower()
    date_str = request.args.get('date')
    if date_str:
        try:
            today = datetime.strptime(date_str, '%Y-%m-%d').date()
        except ValueError:
            today = datetime.utcnow().date()
    else:
        today = datetime.utcnow().date()

    if team == 'towing':
        team_filter = or_(Employee.team == 'towing', Employee.category == 'Towing')
    else:
        team_filter = Employee.team == team

    category = request.args.get('category')

    if category:
        members = Employee.query.filter(team_filter, Employee.is_admin == False, Employee.category == category).all()
    else:
        members = Employee.query.filter(team_filter, Employee.is_admin == False).all()
    member_ids = [m.id for m in members]

    checked_in_today = 0
    if member_ids:
        recs = Attendance.query.filter(
            Attendance.employee_id.in_(member_ids),
            Attendance.date == today,
            Attendance.check_in_time.isnot(None)
        ).all()
        checked_in_today = len({r.employee_id for r in recs})

    junctions_today = 0
    active_junctions = 0
    if member_ids and team in ('field', 'coc', 'towing'):
        junctions_today = JunctionVisit.query.filter(
            JunctionVisit.employee_id.in_(member_ids),
            JunctionVisit.date == today
        ).count()
        active_junctions = JunctionVisit.query.filter(
            JunctionVisit.employee_id.in_(member_ids),
            JunctionVisit.status.in_(['in_progress', 'unresolved']),
            JunctionVisit.date <= today
        ).count()

    return jsonify({
        'team': team,
        'total_members': len(members),
        'checked_in_today': checked_in_today,
        'absent_today': len(members) - checked_in_today,
        'junctions_today': junctions_today,
        'active_junctions': active_junctions
    })


@app.route('/api/admin/team/<team>/employees', methods=['GET'])
def admin_team_employees(team):
    """Team roster with today's attendance status, for one team's admin dashboard tab."""
    team = team.lower()
    date_str = request.args.get('date')
    if date_str:
        try:
            today = datetime.strptime(date_str, '%Y-%m-%d').date()
        except ValueError:
            today = datetime.utcnow().date()
    else:
        today = datetime.utcnow().date()

    if team == 'towing':
        team_filter = or_(Employee.team == 'towing', Employee.category == 'Towing')
    else:
        team_filter = Employee.team == team

    category = request.args.get('category')

    if category:
        members = Employee.query.filter(team_filter, Employee.is_admin == False, Employee.category == category).all()
    else:
        members = Employee.query.filter(team_filter, Employee.is_admin == False).all()

    result = []
    for m in members:
        today_records = Attendance.query.filter_by(employee_id=m.id, date=today).order_by(Attendance.id.asc()).all()
        checked_in = any(r.check_in_time for r in today_records)
        latest = today_records[-1] if today_records else None

        result.append({
            'id': m.id,
            'full_name': m.full_name,
            'designation': m.designation,
            'team': m.team,
            'checked_in_today': checked_in,
            'status': latest.status if latest else 'absent',
            'check_in_time': to_ist(latest.check_in_time) if latest and latest.check_in_time else None,
            'check_out_time': to_ist(latest.check_out_time) if latest and latest.check_out_time else None
        })

    return jsonify(result)


@app.route('/api/admin/team/<team>/junctions', methods=['GET'])
def admin_team_junctions(team):
    """Junction visit log (before/after photos + locations) for one team's admin dashboard tab.
    Supports ?date=YYYY-MM-DD or ?start_date=YYYY-MM-DD&end_date=YYYY-MM-DD and ?employee_id=..."""
    team = team.lower()
    date_filter = request.args.get('date')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    employee_id = request.args.get('employee_id')

    if team == 'towing':
        team_filter = or_(Employee.team == 'towing', Employee.category == 'Towing')
    else:
        team_filter = Employee.team == team

    category = request.args.get('category')

    if employee_id:
        emp = Employee.query.get(employee_id)
        if emp and category and emp.category != category:
            member_ids = []
        else:
            member_ids = [employee_id]
    else:
        if category:
            member_ids = [m.id for m in Employee.query.filter(team_filter, Employee.is_admin == False, Employee.category == category).all()]
        else:
            member_ids = [m.id for m in Employee.query.filter(team_filter, Employee.is_admin == False).all()]

    if not member_ids:
        return jsonify([])

    query = db.session.query(JunctionVisit, Employee).join(Employee).filter(JunctionVisit.employee_id.in_(member_ids))

    if start_date and end_date:
        try:
            s_date = datetime.strptime(start_date, '%Y-%m-%d').date()
            e_date = datetime.strptime(end_date, '%Y-%m-%d').date()
            query = query.filter(
                or_(
                    and_(JunctionVisit.date >= s_date, JunctionVisit.date <= e_date),
                    and_(db.func.date(JunctionVisit.completed_at) >= s_date, db.func.date(JunctionVisit.completed_at) <= e_date)
                )
            )
        except ValueError:
            pass
    elif date_filter:
        try:
            target_date = datetime.strptime(date_filter, '%Y-%m-%d').date()
            query = query.filter(
                or_(
                    JunctionVisit.date == target_date,
                    db.func.date(JunctionVisit.completed_at) == target_date,
                    and_(
                        JunctionVisit.status.in_(['in_progress', 'unresolved']),
                        JunctionVisit.date <= target_date
                    )
                )
            )
        except ValueError:
            pass

    results = query.order_by(JunctionVisit.id.desc()).all()

    junction_names = list(set([v.junction_name for v, e in results]))
    counts_map = {}
    if junction_names:
        # Count main visits
        main_counts = db.session.query(
            JunctionVisit.junction_name, 
            db.func.count(JunctionVisit.id)
        ).filter(JunctionVisit.junction_name.in_(junction_names)).group_by(JunctionVisit.junction_name).all()
        
        # Count call visits
        call_counts = db.session.query(
            JunctionVisit.junction_name, 
            db.func.count(CallVisit.id)
        ).join(CallVisit).filter(JunctionVisit.junction_name.in_(junction_names)).group_by(JunctionVisit.junction_name).all()
        
        main_map = {name: count for name, count in main_counts}
        call_map = {name: count for name, count in call_counts}
        counts_map = {name: main_map.get(name, 0) + call_map.get(name, 0) for name in junction_names}

    return jsonify([{
        'id': v.id,
        'employee_id': v.employee_id,
        'employee_name': e.full_name,
        'junction_name': v.junction_name,
        'visit_count': counts_map.get(v.junction_name, 0),
        'ward': v.ward or '',
        'zone': v.zone or '',
        'date': v.date.isoformat(),
        'before_photo': v.before_photo,
        'after_photo': v.after_photo,
        'before_location': v.before_location,
        'after_location': v.after_location,
        'started_at': to_ist(v.started_at),
        'completed_at': to_ist(v.completed_at),
        'time_spent_minutes': getattr(v, 'time_spent_minutes', 0.0) or 0.0,
        'travel_time_minutes': getattr(v, 'travel_time_minutes', 0.0) or 0.0,
        'status': v.status,
        'visit_type': v.visit_type,
        'asset_type': v.asset_type or '',
        'fault_type': v.fault_type or '',
        'remark': 'Unresolved' if v.remark == 'Auto-closed (Left unresolved)' else (v.remark or ''),
        'before_remark': v.before_remark or '',
        'call_visits': [{
            'id': cv.id,
            'before_photo': cv.before_photo,
            'after_photo': cv.after_photo,
            'before_location': cv.before_location,
            'after_location': cv.after_location,
            'started_at': to_ist(cv.started_at),
            'completed_at': to_ist(cv.completed_at),
            'before_remark': cv.before_remark or '',
            'remark': cv.remark or '',
            'status': cv.status
        } for cv in sorted(v.call_visits, key=lambda x: x.id)]
    } for v, e in results])


@app.route('/api/admin/field-activity-tracker', methods=['GET'])
@app.route('/admin/field-activity-tracker', methods=['GET'])
def admin_field_activity_tracker():
    """Returns field team activity timeline grouped by employee & date for Field Activity Tracker UI."""
    team = request.args.get('team', 'field').lower()
    employee_id = request.args.get('employee_id')
    date_str = request.args.get('date')
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')
    category = request.args.get('category')

    if team == 'towing':
        team_filter = or_(Employee.team == 'towing', Employee.category == 'Towing')
    else:
        team_filter = Employee.team == team

    if employee_id:
        emps = Employee.query.filter(Employee.id == employee_id).all()
    elif category:
        emps = Employee.query.filter(team_filter, Employee.is_admin == False, Employee.category == category).all()
    else:
        emps = Employee.query.filter(team_filter, Employee.is_admin == False).all()

    if not emps:
        return jsonify([])

    emp_ids = [e.id for e in emps]
    emp_map = {e.id: e for e in emps}

    s_date = None
    e_date = None
    if start_date_str and end_date_str:
        try:
            s_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            e_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
        except ValueError:
            pass
    elif date_str:
        try:
            target = datetime.strptime(date_str, '%Y-%m-%d').date()
            s_date = target - timedelta(days=13)
            e_date = target
        except ValueError:
            pass

    if not s_date or not e_date:
        today_val = datetime.now(pytz.UTC).astimezone(IST).date()
        e_date = today_val
        s_date = today_val - timedelta(days=13)

    visits_q = JunctionVisit.query.filter(
        JunctionVisit.employee_id.in_(emp_ids),
        JunctionVisit.date >= s_date,
        JunctionVisit.date <= e_date
    ).order_by(JunctionVisit.date.desc(), JunctionVisit.started_at.asc()).all()

    attendances_q = Attendance.query.filter(
        Attendance.employee_id.in_(emp_ids),
        Attendance.date >= s_date,
        Attendance.date <= e_date,
        Attendance.check_in_time.isnot(None)
    ).all()

    pings_q = LocationPing.query.filter(
        LocationPing.employee_id.in_(emp_ids),
        db.func.date(LocationPing.timestamp) >= s_date,
        db.func.date(LocationPing.timestamp) <= e_date
    ).order_by(LocationPing.timestamp.asc()).all()

    emp_data = {}
    for emp_id in emp_ids:
        emp = emp_map[emp_id]
        emp_data[emp_id] = {
            'employee_id': emp.id,
            'full_name': emp.full_name,
            'team': emp.team or ('towing' if emp.category == 'Towing' else 'field'),
            'designation': emp.designation or ('Towing Team' if emp.category == 'Towing' else 'Field Team'),
            'daily_logs': {},
            'pings': {}
        }

    for ping in pings_q:
        if ping.employee_id in emp_data:
            p_date_str = ping.timestamp.astimezone(IST).date().isoformat()
            if p_date_str not in emp_data[ping.employee_id]['pings']:
                emp_data[ping.employee_id]['pings'][p_date_str] = []
            emp_data[ping.employee_id]['pings'][p_date_str].append(f"{ping.latitude},{ping.longitude}")

    for att in attendances_q:
        if att.employee_id in emp_data:
            d_str = att.date.isoformat()
            if d_str not in emp_data[att.employee_id]['daily_logs']:
                emp_data[att.employee_id]['daily_logs'][d_str] = {
                    'date': d_str,
                    'check_in': None,
                    'visits': []
                }
            emp_data[att.employee_id]['daily_logs'][d_str]['check_in'] = {
                'time': to_ist(att.check_in_time),
                'location': att.check_in_location
            }

    for v in visits_q:
        if v.employee_id in emp_data:
            d_str = v.date.isoformat()
            if d_str not in emp_data[v.employee_id]['daily_logs']:
                emp_data[v.employee_id]['daily_logs'][d_str] = {
                    'date': d_str,
                    'check_in': None,
                    'visits': []
                }
            emp_data[v.employee_id]['daily_logs'][d_str]['visits'].append(v)

    response_list = []
    for emp_id in sorted(emp_data.keys()):
        e_info = emp_data[emp_id]
        daily_logs_list = []
        total_visits_count = 0
        total_site_mins = 0.0
        total_gap_mins = 0.0
        all_stops_coords = []

        sorted_dates = sorted(e_info['daily_logs'].keys(), reverse=True)
        for d_str in sorted_dates:
            day_log = e_info['daily_logs'][d_str]
            visits = day_log['visits']
            check_in = day_log['check_in']
            day_pings = e_info['pings'].get(d_str, [])

            if not check_in and not visits and not day_pings:
                continue

            day_visits_count = len(visits)
            total_visits_count += day_visits_count

            day_site_mins = 0.0
            day_gap_mins = 0.0
            events = []

            if check_in:
                events.append({
                    'type': 'check_in',
                    'time': check_in['time'],
                    'location': check_in['location']
                })
                if check_in['location']:
                    all_stops_coords.append(check_in['location'])

            # Add background location pings as intermediate stops
            for ping_loc in day_pings:
                if ping_loc not in all_stops_coords:
                    all_stops_coords.append(ping_loc)

            for v in visits:
                site_m = getattr(v, 'time_spent_minutes', 0.0) or 0.0
                gap_m = getattr(v, 'travel_time_minutes', 0.0) or 0.0
                day_site_mins += site_m
                day_gap_mins += gap_m

                loc_coord = v.before_location or v.after_location
                if loc_coord and loc_coord not in all_stops_coords:
                    all_stops_coords.append(loc_coord)

                events.append({
                    'type': 'visit',
                    'id': v.id,
                    'junction_name': v.junction_name,
                    'visit_type': v.visit_type or 'Regular Visit',
                    'ward': v.ward or '',
                    'zone': v.zone or '',
                    'started_at': to_ist(v.started_at),
                    'completed_at': to_ist(v.completed_at),
                    'status': v.status,
                    'time_spent_minutes': round(site_m, 1),
                    'travel_time_minutes': round(gap_m, 1),
                    'before_location': v.before_location,
                    'after_location': v.after_location,
                    'remark': v.remark or '',
                    'before_remark': v.before_remark or ''
                })

            total_site_mins += day_site_mins
            total_gap_mins += day_gap_mins

            daily_logs_list.append({
                'date': d_str,
                'visits_count': day_visits_count,
                'site_mins': round(day_site_mins, 1),
                'gap_mins': round(day_gap_mins, 1),
                'events': events,
                'stops_count': len([ev for ev in events if ev.get('location') or ev.get('before_location')]) + len(day_pings)
            })

        response_list.append({
            'employee_id': e_info['employee_id'],
            'full_name': e_info['full_name'],
            'team': e_info['team'],
            'designation': e_info['designation'],
            'active_days': len(daily_logs_list),
            'junctions_visited': total_visits_count,
            'total_time_on_site': round(total_site_mins, 1),
            'total_gap_travel_time': round(total_gap_mins, 1),
            'stops_coords': all_stops_coords,
            'daily_logs': daily_logs_list
        })

    return jsonify(response_list)


@app.route('/api/admin/team/<team>/junctions/export', methods=['GET'])
def admin_team_junctions_export(team):
    """Export junction visit log to Excel for a specific team, date or date range, and employee."""
    team = team.lower()
    date_filter = request.args.get('date')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    employee_id = request.args.get('employee_id')
    
    if team == 'towing':
        team_filter = or_(Employee.team == 'towing', Employee.category == 'Towing')
    else:
        team_filter = Employee.team == team

    category = request.args.get('category')

    if employee_id:
        emp = Employee.query.get(employee_id)
        if emp and category and emp.category != category:
            member_ids = []
        else:
            member_ids = [employee_id]
    else:
        if category:
            member_ids = [m.id for m in Employee.query.filter(team_filter, Employee.is_admin == False, Employee.category == category).all()]
        else:
            member_ids = [m.id for m in Employee.query.filter(team_filter, Employee.is_admin == False).all()]

    if not member_ids:
        df = pd.DataFrame()
    else:
        query = db.session.query(JunctionVisit, Employee).join(Employee).filter(JunctionVisit.employee_id.in_(member_ids))

        if start_date and end_date:
            try:
                s_date = datetime.strptime(start_date, '%Y-%m-%d').date()
                e_date = datetime.strptime(end_date, '%Y-%m-%d').date()
                query = query.filter(
                    or_(
                        and_(JunctionVisit.date >= s_date, JunctionVisit.date <= e_date),
                        and_(db.func.date(JunctionVisit.completed_at) >= s_date, db.func.date(JunctionVisit.completed_at) <= e_date)
                    )
                )
            except ValueError:
                pass
        elif date_filter:
            try:
                target_date = datetime.strptime(date_filter, '%Y-%m-%d').date()
                query = query.filter(
                    or_(
                        JunctionVisit.date == target_date,
                        db.func.date(JunctionVisit.completed_at) == target_date,
                        and_(
                            JunctionVisit.status.in_(['in_progress', 'unresolved']),
                            JunctionVisit.date <= target_date
                        )
                    )
                )
            except ValueError:
                pass

        results = query.order_by(JunctionVisit.id.desc()).all()

        junction_names = list(set([v.junction_name for v, e in results]))
        counts_map = {}
        if junction_names:
            # Count main visits
            main_counts = db.session.query(
                JunctionVisit.junction_name, 
                db.func.count(JunctionVisit.id)
            ).filter(JunctionVisit.junction_name.in_(junction_names)).group_by(JunctionVisit.junction_name).all()
            
            # Count call visits
            call_counts = db.session.query(
                JunctionVisit.junction_name, 
                db.func.count(CallVisit.id)
            ).join(CallVisit).filter(JunctionVisit.junction_name.in_(junction_names)).group_by(JunctionVisit.junction_name).all()
            
            main_map = {name: count for name, count in main_counts}
            call_map = {name: count for name, count in call_counts}
            counts_map = {name: main_map.get(name, 0) + call_map.get(name, 0) for name in junction_names}

        data = []
        image_map = []
        merge_ranges = []
        row_counter = 0
        for (v, e) in results:
            started_ist = utc_to_ist_dt(v.started_at).strftime('%Y-%m-%d %H:%M:%S') if v.started_at else '—'
            
            # 1. Completed date should NOT be shown if status is unresolved (jab tak unresolved hai completed date nahi dikhega)
            completed_ist = '—'
            if v.status == 'completed' and v.completed_at:
                completed_ist = utc_to_ist_dt(v.completed_at).strftime('%Y-%m-%d %H:%M:%S')
                
            # 2. Before photo and before location should show the latest call visit's photo/location if present
            display_before_photo = v.before_photo
            display_before_location = v.before_location
            if v.call_visits:
                latest_cv = sorted(v.call_visits, key=lambda x: x.id)[-1]
                if latest_cv.before_photo:
                    display_before_photo = latest_cv.before_photo
                if latest_cv.before_location:
                    display_before_location = latest_cv.before_location

            # 3. Completion remark in 'Remark' column (no call visit dates or attempt lists here)
            main_remark = v.remark or '—'
            if v.remark == 'Auto-closed (Left unresolved)':
                main_remark = 'Unresolved'
                
            group_size = 1 + len(v.call_visits)
            start_excel_row = row_counter + 2
            end_excel_row = start_excel_row + group_size - 1
            if group_size > 1:
                merge_ranges.append((start_excel_row, end_excel_row))
                
            # Calculate duration spent at junction
            time_spent_str = '—'
            if v.completed_at and v.started_at:
                time_spent_str = f"{round((v.completed_at - v.started_at).total_seconds() / 60.0, 1)} mins"
            elif v.started_at:
                time_spent_str = 'In Progress'

            arr_time_str = utc_to_ist_dt(v.started_at).strftime('%I:%M:%S %p') if v.started_at else '—'
            comp_time_str = utc_to_ist_dt(v.completed_at).strftime('%I:%M:%S %p') if (v.status == 'completed' and v.completed_at) else '—'
            travel_time_str = f"{getattr(v, 'travel_time_minutes', 0.0) or 0.0} mins"

            # Append parent row
            main_visit_date = v.date.strftime('%d-%m-%Y') if v.date else '—'
            data.append({
                'Employee ID': v.employee_id,
                'Employee Name': e.full_name,
                'Junction Name': v.junction_name,
                'Number of Visits': counts_map.get(v.junction_name, 0),
                'Visit Type': v.visit_type or 'Regular Visit',
                'Ward': v.ward or '',
                'Zone': v.zone or '',
                'Visit Date': main_visit_date,
                'Arrival Time (IST)': arr_time_str,
                'Completion Time (IST)': comp_time_str,
                'Time Spent at Junction': time_spent_str,
                'Travel Time to Reach': travel_time_str,
                'Started At (IST)': started_ist,
                'Completed At (IST)': completed_ist,
                'Status': 'Completed' if v.status == 'completed' else ('Unresolved' if v.status == 'unresolved' else 'Open'),
                'Asset Type': v.asset_type or '—',
                'Fault Type': v.fault_type or '—',
                'Before Photo': v.before_photo or '—',
                'After Photo': v.after_photo or '—',
                'Before Location': v.before_location or '—',
                'After Location': v.after_location or '—',
                'Before Remark': v.before_remark or '—',
                'Remark': main_remark
            })
            
            image_map.append({
                'row_idx': row_counter,
                'photos': {
                    'Before Photo': v.before_photo,
                    'After Photo': v.after_photo
                }
            })
            row_counter += 1
            
            # Append call visits rows (each row shows its own call visit's before photo, before location, before remark)
            for cv in sorted(v.call_visits, key=lambda x: x.id):
                cv_started_date = utc_to_ist_dt(cv.started_at).strftime('%d-%m-%Y') if cv.started_at else '—'
                cv_started_time = utc_to_ist_dt(cv.started_at).strftime('%Y-%m-%d %H:%M:%S') if cv.started_at else '—'
                cv_arr_time = utc_to_ist_dt(cv.started_at).strftime('%I:%M:%S %p') if cv.started_at else '—'
                
                data.append({
                    'Employee ID': v.employee_id,
                    'Employee Name': e.full_name,
                    'Junction Name': v.junction_name,
                    'Number of Visits': counts_map.get(v.junction_name, 0),
                    'Visit Type': v.visit_type or 'Regular Visit',
                    'Ward': v.ward or '',
                    'Zone': v.zone or '',
                    'Visit Date': cv_started_date,
                    'Arrival Time (IST)': cv_arr_time,
                    'Completion Time (IST)': comp_time_str,
                    'Time Spent at Junction': time_spent_str,
                    'Travel Time to Reach': travel_time_str,
                    'Started At (IST)': cv_started_time,
                    'Completed At (IST)': completed_ist,
                    'Status': 'Completed' if v.status == 'completed' else ('Unresolved' if v.status == 'unresolved' else 'Open'),
                    'Asset Type': v.asset_type or '—',
                    'Fault Type': v.fault_type or '—',
                    'Before Photo': cv.before_photo or '—',
                    'After Photo': v.after_photo or '—',
                    'Before Location': cv.before_location or '—',
                    'After Location': v.after_location or '—',
                    'Before Remark': cv.before_remark or '—',
                    'Remark': main_remark
                })

                
                image_map.append({
                    'row_idx': row_counter,
                    'photos': {
                        'Before Photo': cv.before_photo
                    }
                })
                row_counter += 1
                
        df = pd.DataFrame(data)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Junction Visits')
        if not df.empty:
            embed_photos_in_excel(writer, 'Junction Visits', image_map, merge_ranges=merge_ranges)
    output.seek(0)

    filename = f'junction_visits_{team}_{datetime.utcnow().date().isoformat()}.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/junction/remark', methods=['POST'])
def add_junction_remark():
    """Allows field/roaming team to post remarks/messages regarding a junction or issue."""
    data = request.get_json()
    if not data:
        return jsonify({'error': 'JSON body is required'}), 400

    employee_id = data.get('employee_id')
    junction_name = data.get('junction_name')
    remark = data.get('remark')

    if not employee_id or not junction_name or not remark:
        return jsonify({'error': 'employee_id, junction_name, and remark are required'}), 400

    employee = Employee.query.get(employee_id)
    if not employee:
        return jsonify({'error': 'Employee not found'}), 404

    new_remark = JunctionRemark(
        employee_id=employee_id,
        junction_name=junction_name.strip(),
        remark=remark.strip(),
        created_at=datetime.utcnow()
    )
    db.session.add(new_remark)
    db.session.commit()

    return jsonify({
        'success': True,
        'message': 'Remark logged successfully',
        'remark': {
            'id': new_remark.id,
            'employee_id': new_remark.employee_id,
            'employee_name': employee.full_name,
            'junction_name': new_remark.junction_name,
            'remark': new_remark.remark,
            'created_at': to_ist(new_remark.created_at)
        }
    }), 201


@app.route('/api/junction/remarks', methods=['GET'])
def get_junction_remarks():
    """Returns list of general remarks/messages.
    Optional query params:
    - date: YYYY-MM-DD
    - start_date & end_date: YYYY-MM-DD
    - employee_id: filter by employee
    - team: filter by employee's team (field/coc/ccc)
    """
    date_filter = request.args.get('date')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    employee_id = request.args.get('employee_id')
    team = request.args.get('team')
    category = request.args.get('category')

    query = db.session.query(JunctionRemark, Employee).join(Employee)
    if category:
        query = query.filter(Employee.category == category)

    if start_date and end_date:
        try:
            s_date = datetime.strptime(start_date, '%Y-%m-%d').date()
            e_date = datetime.strptime(end_date, '%Y-%m-%d').date()
            query = query.filter(
                db.func.date(JunctionRemark.created_at) >= s_date,
                db.func.date(JunctionRemark.created_at) <= e_date
            )
        except ValueError:
            pass
    elif date_filter:
        try:
            target_date = datetime.strptime(date_filter, '%Y-%m-%d').date()
            query = query.filter(db.func.date(JunctionRemark.created_at) == target_date)
        except ValueError:
            pass

    if employee_id:
        query = query.filter(JunctionRemark.employee_id == employee_id)

    if team:
        query = query.filter(Employee.team == team.lower())

    results = query.order_by(JunctionRemark.id.desc()).all()

    return jsonify([{
        'id': r.id,
        'employee_id': r.employee_id,
        'employee_name': e.full_name,
        'employee_team': e.team or '',
        'junction_name': r.junction_name,
        'remark': r.remark,
        'created_at': to_ist(r.created_at)
    } for r, e in results])


@app.route('/api/admin/remarks/export', methods=['GET'])
def admin_remarks_export():
    """Export general remarks/messages to Excel with employee and date range filters."""
    date_filter = request.args.get('date')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    employee_id = request.args.get('employee_id')
    team = request.args.get('team')
    category = request.args.get('category')

    query = db.session.query(JunctionRemark, Employee).join(Employee)
    if category:
        query = query.filter(Employee.category == category)

    if start_date and end_date:
        try:
            s_date = datetime.strptime(start_date, '%Y-%m-%d').date()
            e_date = datetime.strptime(end_date, '%Y-%m-%d').date()
            query = query.filter(
                db.func.date(JunctionRemark.created_at) >= s_date,
                db.func.date(JunctionRemark.created_at) <= e_date
            )
        except ValueError:
            pass
    elif date_filter:
        try:
            target_date = datetime.strptime(date_filter, '%Y-%m-%d').date()
            query = query.filter(db.func.date(JunctionRemark.created_at) == target_date)
        except ValueError:
            pass

    if employee_id:
        query = query.filter(JunctionRemark.employee_id == employee_id)

    if team:
        query = query.filter(Employee.team == team.lower())

    results = query.order_by(JunctionRemark.id.desc()).all()

    data = []
    for r, e in results:
        created_ist = utc_to_ist_dt(r.created_at).strftime('%Y-%m-%d %H:%M:%S') if r.created_at else '—'
        data.append({
            'Date & Time (IST)': created_ist,
            'Employee ID': r.employee_id,
            'Employee Name': e.full_name,
            'Team': (e.team or '').upper(),
            'Junction Name': r.junction_name,
            'Remark / Message': r.remark
        })

    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Junction Remarks')
    output.seek(0)

    filename = f'junction_remarks_{team or "all"}_{datetime.utcnow().date().isoformat()}.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/admin/leaves/export', methods=['GET'])
def export_leaves():
    """Export leave requests report to Excel with employee and date range/month filters."""
    employee_id = request.args.get('employee_id')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    status_filter = request.args.get('status')
    
    query = db.session.query(Leave, Employee).join(Employee)
    
    if employee_id:
        query = query.filter(Leave.employee_id == employee_id)
        
    if status_filter:
        query = query.filter(Leave.status == status_filter)
        
    if start_date:
        try:
            s_date = datetime.strptime(start_date, '%Y-%m-%d').date()
            query = query.filter(Leave.start_date >= s_date)
        except ValueError:
            pass
            
    if end_date:
        try:
            e_date = datetime.strptime(end_date, '%Y-%m-%d').date()
            query = query.filter(Leave.end_date <= e_date)
        except ValueError:
            pass
            
    results = query.order_by(Leave.created_at.desc()).all()
    
    data = []
    image_map = []
    for idx, (leave, employee) in enumerate(results):
        is_half_day = getattr(leave, 'is_half_day', False)
        doc = getattr(leave, 'supporting_document', None)
        data.append({
            'Employee ID': leave.employee_id,
            'Employee Name': employee.full_name,
            'Leave Type': leave.leave_type,
            'Start Date': leave.start_date.isoformat(),
            'End Date': leave.end_date.isoformat(),
            'Half Day': 'Yes' if is_half_day else 'No',
            'Reason': leave.reason or '',
            'Status': leave.status,
            'Admin Comment': leave.admin_comment or '',
            'Supporting Document': doc or '',
            'Applied On': utc_to_ist_dt(leave.created_at).strftime('%Y-%m-%d %H:%M:%S') if leave.created_at else ''
        })
        if doc and str(doc).lower().endswith(('.jpg', '.jpeg', '.png', '.gif')):
            image_map.append({
                'row_idx': idx,
                'photos': {
                    'Supporting Document': doc
                }
            })
        
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Leave Requests')
        if not df.empty and image_map:
            embed_photos_in_excel(writer, 'Leave Requests', image_map)
    output.seek(0)
    
    filename = f'leave_report_{datetime.now().strftime("%Y%m%d")}.xlsx'
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )



def auto_migrate_db():
    """Automatically detects missing columns in MySQL tables based on SQLAlchemy models and adds them."""
    from sqlalchemy import inspect, text
    try:
        # Ensure location_ping table exists directly without metadata inspection conflict
        create_ping_sql = """
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
            INDEX `idx_loc_ping_time` (`timestamp`),
            INDEX `idx_loc_ping_emp` (`employee_id`)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

        """
        try:
            db.session.execute(text(create_ping_sql))
            db.session.commit()
        except Exception as p_err:
            db.session.rollback()

        inspector = inspect(db.engine)
        for model in db.Model.__subclasses__():
            if hasattr(model, '__table__'):
                table_name = model.__table__.name
                if table_name == 'location_ping':
                    continue
                if inspector.has_table(table_name):
                    existing_cols = {col['name'] for col in inspector.get_columns(table_name)}
                    for column in model.__table__.columns:
                        if column.name not in existing_cols:
                            col_type = str(column.type)
                            default_clause = ""
                            if column.default is not None and column.default.arg is not None and not callable(column.default.arg):
                                default_clause = f" DEFAULT '{column.default.arg}'"
                            nullable_clause = " NULL" if column.nullable else ""
                            sql = f"ALTER TABLE `{table_name}` ADD COLUMN `{column.name}` {col_type}{default_clause}{nullable_clause}"
                            try:
                                db.session.execute(text(sql))
                                db.session.commit()
                                print(f"Auto-migrated: Added column '{column.name}' to table '{table_name}'.")
                            except Exception as ex:
                                db.session.rollback()
                                print(f"Auto-migration failed for '{column.name}' on '{table_name}': {ex}")
    except Exception as e:
        print(f"Auto-migration check notice: {e}")


if __name__ == "__main__":
    with app.app_context():
        # Create tables & auto-migrate missing columns cleanly
        try:
            target_tables = [t for t in db.metadata.sorted_tables if t.name != 'location_ping']
            db.metadata.create_all(bind=db.engine, tables=target_tables)
        except Exception as err:
            print(f"db.create_all() notice: {err}")
        auto_migrate_db()


    # Create default admin if no admin exists
        if not Employee.query.filter_by(is_admin=True).first():
            default_admin = Employee(
                id='ADMIN001',
                full_name='System Administrator',
                email='multisulotionsdecofurn@gmail.com',
                phone='+919518791736',
                password='AD#456',
                is_admin=True
            )
            db.session.add(default_admin)
            db.session.commit()
            print("Default admin created successfully.")

        app.run(debug=True, host='0.0.0.0', port=5001)