import requests
import json
from datetime import datetime

BASE_URL = "http://localhost:5000/api"

def test_login():
    """Test admin login"""
    print("=== Testing Admin Login ===")
    login_data = {
        "employee_id": "ADMIN001",
        "password": "admin123"
    }
    
    response = requests.post(f"{BASE_URL}/login", json=login_data)
    print(f"Status Code: {response.status_code}")
    print(f"Response: {response.json()}")
    return response.json().get('success', False)

def test_create_employee():
    """Test creating a new employee"""
    print("\n=== Testing Create Employee ===")
    employee_data = {
        "id": "EMP001",
        "full_name": "John Doe",
        "email": "john.doe@company.com",
        "phone": "9876543210",
        "password": "password123",
        "is_admin": False
    }
    
    response = requests.post(f"{BASE_URL}/admin/employees", json=employee_data)
    print(f"Status Code: {response.status_code}")
    print(f"Response: {response.json()}")
    return response.json().get('success', False)

def test_employee_login():
    """Test employee login"""
    print("\n=== Testing Employee Login ===")
    login_data = {
        "employee_id": "EMP001",
        "password": "password123"
    }
    
    response = requests.post(f"{BASE_URL}/login", json=login_data)
    print(f"Status Code: {response.status_code}")
    print(f"Response: {response.json()}")
    return response.json().get('success', False)

def test_get_employees():
    """Test getting all employees"""
    print("\n=== Testing Get All Employees ===")
    response = requests.get(f"{BASE_URL}/admin/employees")
    print(f"Status Code: {response.status_code}")
    print(f"Response: {response.json()}")

def test_attendance_status():
    """Test getting attendance status"""
    print("\n=== Testing Attendance Status ===")
    response = requests.get(f"{BASE_URL}/attendance/status/EMP001")
    print(f"Status Code: {response.status_code}")
    print(f"Response: {response.json()}")

def test_attendance_stats():
    """Test getting attendance statistics"""
    print("\n=== Testing Attendance Stats ===")
    response = requests.get(f"{BASE_URL}/admin/stats")
    print(f"Status Code: {response.status_code}")
    print(f"Response: {response.json()}")

def main():
    print("Starting API Tests...")
    print("=" * 50)
    
    try:
        # Test admin login
        admin_login_success = test_login()
        
        if admin_login_success:
            # Test creating employee
            create_employee_success = test_create_employee()
            
            if create_employee_success:
                # Test employee login
                test_employee_login()
            
            # Test other endpoints
            test_get_employees()
            test_attendance_status()
            test_attendance_stats()
        
        print("\n" + "=" * 50)
        print("API Tests Completed!")
        
    except requests.exceptions.ConnectionError:
        print("Error: Could not connect to the server. Make sure the Flask app is running on http://localhost:5000")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
