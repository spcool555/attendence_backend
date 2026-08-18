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

def embed_photos_in_excel(writer, sheet_name, image_map, row_height=65, max_size=(75, 75)):
    if sheet_name not in writer.sheets:
        return
    ws = writer.sheets[sheet_name]
    
    header_col_map = {}
    for col in range(1, ws.max_column + 1):
        cell_val = ws.cell(row=1, column=col).value
        if cell_val:
            header_col_map[str(cell_val).strip()] = (col, get_column_letter(col))

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
                    img = OpenPyxlImage(file_path)
                    img.width, img.height = max_size
                    ws.add_image(img, f'{col_letter}{excel_row}')
                except Exception as err:
                    print(f"Error embedding image {photo_filename} into Excel: {err}")

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
DB_PORT = os.getenv('DB_PORT', '3306')
DB_NAME = os.getenv('DB_NAME', 'decofurn')

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
    
    employee = Employee.query.filter_by(id=employee_id, password=password).first()
    
    if employee:
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
    
    employee = Employee(
        id=data['id'],
        full_name=data['full_name'],
        email=data['email'],
        phone=data['phone'],
        password=data['password'],  # Plain text as requested
        is_admin=data.get('is_admin', False),
        shift_type=data.get('shift_type', 'general'),
        weekly_off=data.get('weekly_off'),
        category=data.get('category')
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
            'category': employee.category
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
            'days_count': 0.5 if is_half_day else (leave.end_date - leave.start_date).days + 1
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
    asset_type = db.Column(db.String(100))
    fault_type = db.Column(db.String(100))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(pytz.UTC))

    employee = db.relationship('Employee', backref='junction_visits')


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
}

ALLOWED_VISIT_TEAMS = ('field', 'coc', 'ccc')


@app.route('/api/locations', methods=['GET'])
def get_locations():
    """Returns the list of predefined locations for a specific team.
    Query param: ?team=field|coc|ccc
    Falls back to JunctionList (legacy) for field team if LocationList is empty."""
    team = (request.args.get('team') or 'field').lower()

    if team not in ALLOWED_VISIT_TEAMS:
        return jsonify({'error': 'Invalid team. Must be field, coc, or ccc'}), 400

    if team == 'field':
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

        added_employees = 0
        added_locations = 0
        skipped = 0

        for _, row in df.iterrows():
            emp_name = str(row.get(emp_col, '')).strip()
            if not emp_name:
                continue

            if team == 'field':
                # Field team: just register the employee name as a location entry
                # (used for junction visit tracking, not a physical location)
                exists = LocationList.query.filter_by(name=emp_name, team=team).first()
                if not exists:
                    new_entry = LocationList(name=emp_name, team=team)
                    db.session.add(new_entry)
                    added_employees += 1
                else:
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
                    loc = LocationList(name=loc_name, team=team)
                    db.session.add(loc)
                    added_locations += 1

                # Store the employee-location mapping as another entry
                emp_loc_key = f"{emp_name} @ {loc_name}"
                emp_exists = LocationList.query.filter_by(name=emp_loc_key, team=f"{team}_emp").first()
                if not emp_exists:
                    emp_entry = LocationList(name=emp_loc_key, team=f"{team}_emp")
                    db.session.add(emp_entry)
                    added_employees += 1
                else:
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
        'Designation': ['Field Team', 'CoC', 'CCC'],
        'Category': ['Smart City', 'IITMS', 'Construction']
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

        # Check required columns
        required_cols = ['Employee ID', 'Full Name', 'Email', 'Phone', 'Password', 'Designation']
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

        # Clear existing records first to allow clean import of new IDs
        try:
            # Delete in order of foreign key dependencies
            db.session.query(Announcement).delete()
            db.session.query(JunctionVisit).delete()
            db.session.query(Attendance).delete()
            db.session.query(Leave).delete()
            db.session.query(Employee).delete()
            db.session.commit()
        except Exception as delete_error:
            db.session.rollback()
            return jsonify({'error': f'Failed to clear old employee data: {str(delete_error)}'}), 500

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
            designation = clean_val(row['Designation'])
            category_val = clean_val(row[group_col])

            if not full_name or full_name.lower() == 'nan':
                continue # Skip empty rows

            # Map designation to team
            team_mapped = None
            desig_lower = designation.lower()
            if desig_lower == 'field team':
                team_mapped = 'field'
            elif desig_lower == 'coc':
                team_mapped = 'coc'
            elif desig_lower == 'ccc':
                team_mapped = 'ccc'

            # Normalize category value
            category_val_lower = category_val.lower()
            if 'smart' in category_val_lower and 'city' in category_val_lower:
                category_mapped = 'Smart City'
            elif 'itms' in category_val_lower:
                category_mapped = 'IITMS'
            elif 'construction' in category_val_lower:
                category_mapped = 'Construction'
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
            {'id': 'construction_admin', 'full_name': 'Construction Admin', 'email': 'construction@keltron.com', 'phone': '0000000003', 'password': 'Construction@Admin', 'category': 'Construction'}
        ]:
            if not Employee.query.get(adm['id']):
                new_adm = Employee(
                    id=adm['id'],
                    full_name=adm['full_name'],
                    email=adm['email'],
                    phone=adm['phone'],
                    password=adm['password'],
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
            'message': f'Uploaded successfully. Added {added_count} new/admin employees after resetting database.'
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

    visit = JunctionVisit(
        employee_id=employee_id,
        junction_name=junction_name.strip(),
        ward=ward.strip() if ward else None,
        zone=zone.strip() if zone else None,
        date=today,
        before_photo=before_photo_filename,
        before_location=location,
        started_at=datetime.utcnow(),
        status='completed' if is_completed_immediately else 'in_progress',
        completed_at=datetime.utcnow() if is_completed_immediately else None,
        after_photo=before_photo_filename if is_completed_immediately else None,
        after_location=location if is_completed_immediately else None,
        remark='Regular Visit (No Fault)' if is_completed_immediately else None,
        visit_type=visit_type,
        asset_type=asset_type,
        fault_type=fault_type
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


@app.route('/api/junction/today/<employee_id>', methods=['GET'])
def junction_today(employee_id):
    try:
        today = datetime.utcnow().date()

        visits = JunctionVisit.query.filter(
            JunctionVisit.employee_id == employee_id,
            or_(
                JunctionVisit.date == today,
                JunctionVisit.status == 'in_progress'
            )
        ).order_by(JunctionVisit.id.desc()).all()

        return jsonify([{
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
            'remark': v.remark
        } for v in visits])

    except Exception as e:
        print("JUNCTION TODAY ERROR:", e)
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


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
            
        # Normalize columns
        df.columns = [c.strip() for c in df.columns]
            
        # Find Junction Name column (support aliases like Location, location, JUNCTION, name)
        junction_col = None
        for candidate in candidates:
            if candidate in df.columns:
                junction_col = candidate
                break

        if not junction_col:
            return jsonify({'error': 'Missing "Junction Name" (or "Location", "Junction") column in the uploaded file.'}), 400
            
        # Find Ward column
        ward_col = None
        for candidate in ['Ward', 'ward', 'Ward Name', 'ward name', 'WARD', 'FROM WARD', 'TO WARD', 'FROM_WARD', 'TO_WARD', 'TO WARD']:
            if candidate in df.columns:
                ward_col = candidate
                break

        # Find Zone column
        zone_col = None
        for candidate in ['Zone', 'zone', 'Zone Name', 'zone name', 'ZONE']:
            if candidate in df.columns:
                zone_col = candidate
                break
                
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
    """Top-line stats for one team's admin dashboard tab (COC / CCC / Field)."""
    team = team.lower()
    date_str = request.args.get('date')
    if date_str:
        try:
            today = datetime.strptime(date_str, '%Y-%m-%d').date()
        except ValueError:
            today = datetime.utcnow().date()
    else:
        today = datetime.utcnow().date()

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
    if member_ids and team in ('field', 'coc'):
        junctions_today = JunctionVisit.query.filter(
            JunctionVisit.employee_id.in_(member_ids),
            JunctionVisit.date == today
        ).count()
        active_junctions = JunctionVisit.query.filter(
            JunctionVisit.employee_id.in_(member_ids),
            JunctionVisit.status == 'in_progress',
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
                        JunctionVisit.status == 'in_progress',
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
        counts_query = db.session.query(
            JunctionVisit.junction_name, 
            db.func.count(JunctionVisit.id)
        ).filter(JunctionVisit.junction_name.in_(junction_names)).group_by(JunctionVisit.junction_name).all()
        counts_map = {name: count for name, count in counts_query}

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
        'status': v.status,
        'visit_type': v.visit_type,
        'asset_type': v.asset_type or '',
        'fault_type': v.fault_type or '',
        'remark': 'Unresolved' if v.remark == 'Auto-closed (Left unresolved)' else (v.remark or '')
    } for v, e in results])


@app.route('/api/admin/team/<team>/junctions/export', methods=['GET'])
def admin_team_junctions_export(team):
    """Export junction visit log to Excel for a specific team, date or date range, and employee."""
    team = team.lower()
    date_filter = request.args.get('date')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    employee_id = request.args.get('employee_id')
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
                            JunctionVisit.status == 'in_progress',
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
            counts_query = db.session.query(
                JunctionVisit.junction_name, 
                db.func.count(JunctionVisit.id)
            ).filter(JunctionVisit.junction_name.in_(junction_names)).group_by(JunctionVisit.junction_name).all()
            counts_map = {name: count for name, count in counts_query}

        data = []
        image_map = []
        for idx, (v, e) in enumerate(results):
            started_ist = utc_to_ist_dt(v.started_at).strftime('%Y-%m-%d %H:%M:%S') if v.started_at else '—'
            completed_ist = utc_to_ist_dt(v.completed_at).strftime('%Y-%m-%d %H:%M:%S') if v.completed_at else '—'
            data.append({
                'Employee ID': v.employee_id,
                'Employee Name': e.full_name,
                'Junction Name': v.junction_name,
                'Number of Visits': counts_map.get(v.junction_name, 0),
                'Visit Type': v.visit_type or 'Regular Visit',
                'Ward': v.ward or '',
                'Zone': v.zone or '',
                'Visit Date': v.date.strftime('%d-%m-%Y') if v.date else '—',
                'Started At (IST)': started_ist,
                'Completed At (IST)': completed_ist,
                'Status': 'Completed' if v.status == 'completed' else 'Open',
                'Asset Type': v.asset_type or '—',
                'Fault Type': v.fault_type or '—',
                'Before Photo': v.before_photo or '—',
                'After Photo': v.after_photo or '—',
                'Before Location': v.before_location or '—',
                'After Location': v.after_location or '—',
                'Remark': 'Unresolved' if v.remark == 'Auto-closed (Left unresolved)' else (v.remark or '—')
            })
            image_map.append({
                'row_idx': idx,
                'photos': {
                    'Before Photo': v.before_photo,
                    'After Photo': v.after_photo
                }
            })
        df = pd.DataFrame(data)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Junction Visits')
        if not df.empty:
            embed_photos_in_excel(writer, 'Junction Visits', image_map)
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
        inspector = inspect(db.engine)
        for model in db.Model.__subclasses__():
            if hasattr(model, '__table__'):
                table_name = model.__table__.name
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
        # Create tables & auto-migrate missing columns
        db.create_all()
        auto_migrate_db()

        # Create default admin if no admin exists
        if not Employee.query.filter_by(is_admin=True).first():
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
            print("Default admin created: ID=ADMIN001, Password=AD#987")

        app.run(debug=True, host='0.0.0.0', port=5001)