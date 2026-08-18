# Attendance Tracker Backend

A Flask-based backend for the attendance tracking application with MySQL database.

## Features

- Employee authentication (plain text passwords as requested)
- Check-in/Check-out with photo and GPS location
- Admin dashboard with attendance logs
- Export attendance data to Excel
- Attendance status calculation (Present, Late, Half-day, Leave)

## Setup Instruction
### Prerequisites

1. Python 3.11+ (you're using conda environment `aicopilot`)
2. MySQL Server running on 31.97.224.19:3306
3. MySQL credentials: username=`spcool`, password= `Spcool@123`

### Installation

1. **Activate your conda environment:**
   ```bash
   conda activate aicopilot
   ```

2. **Install required packages:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Initialize the database:**
   ```bash
   python init_db.py
   ```
   This will:
   - Create the `decofurn` database
   - Create all required tables
   - Create a default admin user (ID: ADMIN001, Password: admin123)

4. **Run the application:**
   ```bash
   python app.py
   ```

The server will start on `http://localhost:5000`

## API Endpoints

### Authentication
- `POST /api/login` - Employee login

### Attendance
- `POST /api/attendance/check-in` - Check-in with photo and location
- `POST /api/attendance/check-out` - Check-out with photo and location
- `GET /api/attendance/status/<employee_id>` - Get today's attendance status

### Admin
- `GET /api/admin/employees` - Get all employees
- `POST /api/admin/employees` - Create new employee
- `GET /api/admin/attendance` - Get attendance logs (with filters)
- `GET /api/admin/attendance/export` - Export attendance to Excel
- `GET /api/admin/stats` - Get attendance statistics

## Database Schema

### Employee Table
- `id` (Primary Key) - Employee ID
- `full_name` - Full name
- `email` - Email address (unique)
- `phone` - Phone number
- `password` - Plain text password
- `is_admin` - Admin flag
- `created_at` - Creation timestamp

### Attendance Table
- `id` (Primary Key) - Auto-increment ID
- `employee_id` - Foreign key to Employee
- `date` - Attendance date
- `check_in_time` - Check-in timestamp
- `check_out_time` - Check-out timestamp
- `check_in_photo` - Check-in photo filename
- `check_out_photo` - Check-out photo filename
- `check_in_location` - GPS coordinates for check-in
- `check_out_location` - GPS coordinates for check-out
- `status` - Attendance status (present, late, half_day, leave)
- `user_message` - Optional user message
- `created_at` - Creation timestamp

## Business Rules

- **Office Hours:** 9:00 AM start time
- **Grace Period:** 15 minutes (till 9:15 AM) - marked as "Present"
- **Late:** After 9:15 AM but before 1:30 PM - marked as "Late"
- **Half Day:** After 1:30 PM - marked as "Half Day"
- **Leave:** No check-in for the day - marked as "Leave"

## File Upload

Photos are stored in the `uploads/` directory with the naming convention:
`{employee_id}_{timestamp}_{checkin/checkout}.{extension}`

## Default Admin User

- **Employee ID:** ADMIN001
- **Password:** admin123
- **Email:** admin@company.com

## Environment Variables

The application uses the following environment variables (defined in `.env`):

- `DB_HOST` - MySQL host (default: 31.97.224.19)
- `DB_USER` - MySQL username (default: spcool)
- `DB_PASSWORD` - MySQL password (default: Spcool@123)
- `DB_PORT` - MySQL port (default: 3306)
- `DB_NAME` - Database name (default: decofurn)
- `SECRET_KEY` - Flask secret key
- `FLASK_ENV` - Flask environment (development/production)

## Troubleshooting

1. **Database Connection Error:**
   - Ensure MySQL server is running
   - Verify credentials in `.env` file
   - Check if the database exists

2. **Import Error:**
   - Make sure all packages are installed: `pip install -r requirements.txt`
   - Verify you're in the correct conda environment

3. **File Upload Issues:**
   - Check if `uploads/` directory exists and has write permissions
   - Verify file size is under 16MB limit
