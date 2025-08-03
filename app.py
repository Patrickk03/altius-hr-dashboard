import streamlit as st
import pandas as pd
import json
import tempfile
import os
from datetime import datetime, timedelta
import openpyxl
import xlrd
import calendar
import io
import sqlite3

# Original functions from your Tkinter code (integrated here)
def time_to_str(time_val):
    """Convert time value to string, handling NaN, '--:--', or empty values."""
    if pd.isna(time_val) or str(time_val).strip() in ['--:--', '']:
        return None
    return str(time_val).strip()

def calculate_total_hours(in_time, out_time):
    """Calculate total hours between in_time and out_time, return as HH:MM."""
    if in_time is None or out_time is None:
        return "00:00"
    try:
        in_time_dt = datetime.strptime(str(in_time), '%H:%M')
        out_time_dt = datetime.strptime(str(out_time), '%H:%M')
        delta = out_time_dt - in_time_dt
        if delta.total_seconds() < 0:
            delta += timedelta(days=1)
        hours, minutes = divmod(delta.seconds, 3600)
        minutes = minutes // 60
        return f"{hours:02d}:{minutes:02d}"
    except (ValueError, TypeError):
        return "00:00"

def determine_status(total_hours, att_date):
    """Determine status based on total hours and whether the date is a Saturday or Sunday."""
    try:
        date_obj = pd.to_datetime(att_date, format='%Y-%d-%m', errors='coerce')
        if pd.isna(date_obj):
            return "Absent"
        is_sunday = date_obj.weekday() == 6
        is_saturday = date_obj.weekday() == 5
        if is_sunday:
            return "Full day"
        if total_hours == "00:00":
            return "Absent"
        hours, minutes = map(int, str(total_hours).split(':'))
        total_minutes = hours * 60 + minutes
        if is_saturday:
            if total_minutes >= 5 * 60:
                return "Full day"
            elif total_minutes >= 2.5 * 60:
                return "Half day"
            else:
                return "Absent"
        else:
            if total_minutes >= 8 * 60:
                return "Full day"
            elif total_minutes > 4 * 60:
                return "Half day"
            else:
                return "Absent"
    except (ValueError, TypeError):
        return "Absent"

def extract_month_year(df, file_path):
    """Extract month/year from DataFrame or file."""
    try:
        report_month_rows = df[df[8].astype(str).str.contains("Report Month", na=False)]
        for _, row in report_month_rows.iterrows():
            month_str = row[8].split(':')[-1].strip()
            date_obj = pd.to_datetime(month_str, format='%B-%Y', errors='coerce')
            if not pd.isna(date_obj):
                return date_obj.strftime('%m/%Y')
        if file_path.endswith('.xls'):
            book = xlrd.open_workbook(file_path)
            sheet = book.sheet_by_index(0)
            if sheet.nrows > 3:
                for col in range(sheet.ncols):
                    cell = sheet.cell_value(3, col)
                    cell = str(cell).strip()
                    if 'To' in cell:
                        date_str = cell.split(' To ')[0].strip()
                        date_obj = pd.to_datetime(date_str, errors='coerce')
                        if not pd.isna(date_obj):
                            return date_obj.strftime('%m/%Y')
        else:
            with pd.ExcelFile(file_path, engine='openpyxl') as xls:
                sheet = xls.parse(xls.sheet_names[0], header=None, nrows=5)
                row_4 = sheet.iloc[3].dropna().astype(str).str.strip()
                for cell in row_4:
                    if 'To' in cell:
                        date_str = cell.split(' To ')[0].strip()
                        date_obj = pd.to_datetime(date_str, errors='coerce')
                        if not pd.isna(date_obj):
                            return date_obj.strftime('%m/%Y')
        return "07/2025"  # Default to July 2025
    except Exception as e:
        print(f"Error extracting month/year from {file_path}: {e}")
        return "07/2025"

def find_column_indices(df, header_row, file_type):
    """Find column indices based on file type."""
    headers = df.iloc[header_row].astype(str).str.strip().str.lower()
    if file_type == "altius":
        col_mapping = {'Att. Date': None, 'InTime': None, 'OutTime': None}
        for idx, header in enumerate(headers):
            if header == 'att. date':
                col_mapping['Att. Date'] = idx
            elif header == 'intime':
                col_mapping['InTime'] = idx
            elif header == 'outtime':
                col_mapping['OutTime'] = idx
        if not all(col_mapping.values()):
            raise KeyError(f"Could not find all required columns in row {header_row + 1}")
    else:
        col_mapping = {'Date': None, 'IN': None, 'Out': None}
        for idx, header in enumerate(headers):
            if header == 'date':
                col_mapping['Date'] = idx
            elif header == 'in':
                col_mapping['IN'] = idx
            elif header == 'out':
                col_mapping['Out'] = idx
        if not all(col_mapping.values()):
            col_mapping = {'Date': 0, 'IN': 2, 'Out': 17}
        if len(df.iloc[header_row]) < max(col_mapping.values()) + 1:
            raise KeyError(f"Row {header_row + 1} does not have enough columns")
    return col_mapping

def calculate_salary(status, daily_salary):
    """Calculate daily salary based on status."""
    if status in ["Full day", "WFH"]:
        return daily_salary
    elif status == "Half day":
        return daily_salary / 2
    else:
        return 0

def get_latest_date(df, file_type, header_row):
    """Extract the latest date from the DataFrame."""
    try:
        col_mapping = find_column_indices(df, header_row, file_type)
        date_key = 'Att. Date' if file_type == "altius" else 'Date'
        dates = df[col_mapping[date_key]].dropna()
        valid_dates = pd.to_datetime(dates, dayfirst=True, errors='coerce').dropna()
        if not valid_dates.empty:
            return valid_dates.max()
        return None
    except Exception:
        return None

def process_excel_file(df, file_path, json_data, file_type, employees, start_date, end_date, days_in_month):
    """Process a single Excel file and update json_data with salary for 1st-to-24th period."""
    identifier_col = 3 if file_type == "altius" else 7
    identifier = "Employee Name :" if file_type == "altius" else "Name"
    name_col = 7 if file_type == "altius" else 9
    name_rows = df[df[identifier_col].astype(str).str.strip() == identifier].index
    for name_row in name_rows:
        emp_name = str(df.iloc[name_row, name_col]).strip()
        if not emp_name or emp_name == 'nan':
            continue
        emp_id = next((data['employee_id'] for name, data in employees.items() if name == emp_name), None)
        if not emp_id:
            continue
        if emp_id not in json_data["Employee ID"]:
            json_data["Employee ID"][emp_id] = {"name": emp_name, "date": {}, "total_salary": 0}
        header_row = name_row + 1
        try:
            col_mapping = find_column_indices(df, header_row, file_type)
        except KeyError:
            continue
        start_row = name_row + 2
        end_row = df.index[-1] + 1 if name_row == name_rows[-1] else name_rows[name_rows > name_row][0]
        date_key = 'Att. Date' if file_type == "altius" else 'Date'
        in_key = 'InTime' if file_type == "altius" else 'IN'
        out_key = 'OutTime' if file_type == "altius" else 'Out'
        daily_salary = employees.get(emp_name, {}).get("monthly_salary", 0) / days_in_month
        total_salary = 0
        for row_idx in range(start_row, end_row):
            row = df.iloc[row_idx]
            att_date_val = row[col_mapping[date_key]]
            if pd.isna(att_date_val):
                continue
            try:
                date_obj = pd.to_datetime(att_date_val, dayfirst=True)
                att_date = date_obj.strftime('%Y-%d-%m')
                date_obj = datetime.strptime(att_date, '%Y-%d-%m')
                if not (start_date <= date_obj <= end_date):
                    continue
                day_of_week = date_obj.strftime('%A')
            except Exception:
                continue
            if date_obj.weekday() == 6:
                total_hours = "00:00"
                in_time = None
                out_time = None
                status = "Full day"
                salary = calculate_salary(status, daily_salary)
            else:
                in_time = time_to_str(row[col_mapping[in_key]])
                out_time = time_to_str(row[col_mapping[out_key]])
                total_hours = calculate_total_hours(in_time, out_time)
                status = determine_status(total_hours, att_date)
                salary = calculate_salary(status, daily_salary)
            json_data["Employee ID"][emp_id]["date"][att_date] = {
                "In Time": in_time,
                "Out Time": out_time,
                "Total hours": total_hours,
                "Status": status,
                "Salary": salary,
                "Remark": "",
                "Day": day_of_week
            }
            total_salary += salary
        json_data["Employee ID"][emp_id]["total_salary"] += total_salary

def fill_missing_dates(json_data, start_date, end_date, employees, days_in_month):
    """Fill missing dates in the period with Absent status."""
    delta = end_date - start_date
    for emp_id, emp_data in json_data["Employee ID"].items():
        emp_name = emp_data["name"]
        daily_salary = employees.get(emp_name, {}).get("monthly_salary", 0) / days_in_month
        for i in range(delta.days + 1):
            date = start_date + timedelta(days=i)
            att_date = date.strftime('%Y-%d-%m')
            if att_date not in emp_data["date"]:
                emp_data["date"][att_date] = {
                    "In Time": None,
                    "Out Time": None,
                    "Total hours": "00:00",
                    "Status": "Absent",
                    "Salary": 0,
                    "Remark": "",
                    "Day": date.strftime('%A')
                }

# SQLite Functions
def init_db():
    conn = sqlite3.connect('hr_data.db')
    c = conn.cursor()
    # Employees table
    c.execute('''CREATE TABLE IF NOT EXISTS employees
                 (employee_id TEXT PRIMARY KEY, name TEXT UNIQUE, email TEXT, mobile TEXT, designation TEXT, 
                  bank_name TEXT, account_number TEXT, ifsc TEXT, monthly_salary REAL)''')
    # Attendance table
    c.execute('''CREATE TABLE IF NOT EXISTS attendance
                 (emp_id TEXT, att_date TEXT, in_time TEXT, out_time TEXT, total_hours TEXT, status TEXT, 
                  salary REAL, remark TEXT, day TEXT, PRIMARY KEY (emp_id, att_date))''')
    conn.commit()
    conn.close()

def load_employees():
    init_db()
    conn = sqlite3.connect('hr_data.db')
    df = pd.read_sql_query("SELECT * FROM employees", conn)
    conn.close()
    employees = {}
    for _, row in df.iterrows():
        employees[row['name']] = {
            "employee_id": row['employee_id'],
            "email": row['email'],
            "mobile": row['mobile'],
            "designation": row['designation'],
            "bank_name": row['bank_name'],
            "account_number": row['account_number'],
            "ifsc": row['ifsc'],
            "monthly_salary": row['monthly_salary']
        }
    return employees

def save_employees(employees):
    init_db()
    conn = sqlite3.connect('hr_data.db')
    c = conn.cursor()
    c.execute("DELETE FROM employees")
    for name, data in employees.items():
        c.execute('''INSERT OR REPLACE INTO employees 
                     (employee_id, name, email, mobile, designation, bank_name, account_number, ifsc, monthly_salary) 
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (data['employee_id'], name, data['email'], data['mobile'], data['designation'],
                   data['bank_name'], data['account_number'], data['ifsc'], data['monthly_salary']))
    conn.commit()
    conn.close()

def load_attendance():
    init_db()
    conn = sqlite3.connect('hr_data.db')
    df = pd.read_sql_query("SELECT * FROM attendance", conn)
    conn.close()
    attendance = {"Month/year": "07/2025", "Employee ID": {}}
    for _, row in df.iterrows():
        emp_id = row['emp_id']
        if emp_id not in attendance["Employee ID"]:
            attendance["Employee ID"][emp_id] = {"name": "", "date": {}, "total_salary": 0}
        attendance["Employee ID"][emp_id]["date"][row['att_date']] = {
            "In Time": row['in_time'],
            "Out Time": row['out_time'],
            "Total hours": row['total_hours'],
            "Status": row['status'],
            "Salary": row['salary'],
            "Remark": row['remark'],
            "Day": row['day']
        }
        attendance["Employee ID"][emp_id]["total_salary"] += row['salary']
    # Load names from employees
    employees = load_employees()
    for emp_id in attendance["Employee ID"]:
        attendance["Employee ID"][emp_id]["name"] = next((name for name, data in employees.items() if data['employee_id'] == emp_id), "")
    return attendance

def save_attendance(attendance):
    init_db()
    conn = sqlite3.connect('hr_data.db')
    c = conn.cursor()
    c.execute("DELETE FROM attendance")
    for emp_id, data in attendance["Employee ID"].items():
        for att_date, att in data["date"].items():
            c.execute('''INSERT OR REPLACE INTO attendance 
                         (emp_id, att_date, in_time, out_time, total_hours, status, salary, remark, day) 
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                      (emp_id, att_date, att["In Time"], att["Out Time"], att["Total hours"], att["Status"],
                       att["Salary"], att["Remark"], att["Day"]))
    conn.commit()
    conn.close()

# Streamlit app configuration for a more open, HR dashboard-like UI
st.set_page_config(page_title="Altius Investech HR Dashboard", layout="wide", page_icon=":office_worker:")

# Custom CSS for better looking UI: spacious, modern, HR-themed (blues, clean fonts, cards)
st.markdown("""
    <style>
    /* General styling */
    .stApp {
        background-color: #f0f4f8;
    }
    h1, h2, h3 {
        color: #2c3e50;
        font-family: 'Helvetica Neue', sans-serif;
    }
    .stButton > button {
        background-color: #3498db;
        color: white;
        border-radius: 5px;
        padding: 10px 20px;
        font-weight: bold;
    }
    .stButton > button:hover {
        background-color: #2980b9;
    }
    .stDataFrame {
        border: 1px solid #ddd;
        border-radius: 5px;
        background-color: white;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .stExpander {
        border: 1px solid #ddd;
        border-radius: 5px;
        background-color: white;
        margin-bottom: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .stFileUploader {
        background-color: white;
        border: 1px solid #ddd;
        border-radius: 5px;
        padding: 10px;
    }
    /* Card-like sections */
    .card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }
    /* Spacious layout */
    .block-container {
        padding: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# Sidebar for navigation
st.sidebar.title("Navigation")
st.sidebar.markdown("---")
selected_tab = st.sidebar.radio("Go to", ["Dashboard Overview", "File Upload", "Employee Management", "Attendance Search", "Reports"])

# Initialize session state for data persistence
if 'employees' not in st.session_state:
    st.session_state.employees = load_employees()
if 'attendance' not in st.session_state:
    st.session_state.attendance = load_attendance()

# Main title
st.title("Altius Investech HR Dashboard")
st.markdown("Manage employees, process attendance, and generate reports efficiently.", unsafe_allow_html=True)

if selected_tab == "Dashboard Overview":
    st.header("Dashboard Overview")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Employees", len(st.session_state.employees))
    with col2:
        st.metric("Processed Months", "July 2025")
    with col3:
        st.metric("Total Salary Processed", sum(data.get("total_salary", 0) for data in st.session_state.attendance.get("Employee ID", {}).values()))
    st.markdown("---")
    st.subheader("Attendance Summary")
    if st.session_state.attendance.get("Employee ID"):
        att_df = pd.DataFrame([{"Employee": data["name"], "Total Salary": data["total_salary"]} for data in st.session_state.attendance["Employee ID"].values()])
        st.bar_chart(att_df.set_index("Employee"))

elif selected_tab == "File Upload":
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.header("File Upload")
        st.write("Detected Month: 07/2025")
        st.write("Upload attendance files for July 1-24, 2025 (GC Office and Merlin Heights).")

        col1, col2 = st.columns(2)
        with col1:
            altius_current = st.file_uploader("GC Office (Current Month)", type=["xls", "xlsx"])
            altius_prev = st.file_uploader("GC Office (Previous Month)", type=["xls", "xlsx"])
        with col2:
            monthinout_current = st.file_uploader("Merlin Heights (Current Month)", type=["xls", "xlsx"])
            monthinout_prev = st.file_uploader("Merlin Heights (Previous Month)", type=["xls", "xlsx"])

        if st.button("Process Files", key="process_files"):
            with st.spinner("Processing files..."):
                start_date = datetime(2025, 7, 1)
                end_date = datetime(2025, 7, 24)
                days_in_month = calendar.monthrange(2025, 7)[1]  # 31
                json_data = {"Month/year": "07/2025", "Employee ID": {}}

                uploaded_files = {
                    "altius_current": altius_current,
                    "altius_prev": altius_prev,
                    "monthinout_current": monthinout_current,
                    "monthinout_prev": monthinout_prev
                }

                progress_bar = st.progress(0)
                total_files = sum(1 for f in uploaded_files.values() if f is not None)
                processed = 0

                for file_type, uploaded_file in uploaded_files.items():
                    if uploaded_file:
                        with tempfile.NamedTemporaryFile(delete=False) as tmp:
                            tmp.write(uploaded_file.getvalue())
                            file_path = tmp.name
                        try:
                            engine = 'xlrd' if file_path.endswith('.xls') else 'openpyxl'
                            df = pd.read_excel(file_path, engine=engine, header=None)
                            process_excel_file(df, file_path, json_data, file_type.split('_')[0], st.session_state.employees, start_date, end_date, days_in_month)
                            processed += 1
                            progress_bar.progress(processed / total_files)
                        except Exception as e:
                            st.error(f"Failed to process {uploaded_file.name}: {e}")
                        finally:
                            os.unlink(file_path)

                fill_missing_dates(json_data, start_date, end_date, st.session_state.employees, days_in_month)
                st.session_state.attendance = json_data
                save_attendance(json_data)
                st.success(f"Files processed for July 2025, period {start_date.strftime('%Y-%d-%m')} to {end_date.strftime('%Y-%d-%m')}")
        st.markdown('</div>', unsafe_allow_html=True)

elif selected_tab == "Employee Management":
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.header("Employee Management")
        emp_list = [{"ID": data["employee_id"], "Name": name, "Email": data["email"], "Mobile": data["mobile"],
                     "Designation": data["designation"], "Bank Name": data["bank_name"], "Account": data["account_number"],
                     "IFSC": data["ifsc"], "Monthly Salary": data["monthly_salary"]} for name, data in st.session_state.employees.items()]
        emp_df = pd.DataFrame(emp_list)
        st.dataframe(emp_df, use_container_width=True)

        with st.expander("Add Employee", expanded=False):
            with st.form("Add Employee"):
                name = st.text_input("Name")
                email = st.text_input("Email")
                mobile = st.text_input("Mobile")
                designation = st.text_input("Designation")
                bank_name = st.text_input("Bank Name")
                account_number = st.text_input("Account Number")
                ifsc = st.text_input("IFSC")
                monthly_salary = st.number_input("Monthly Salary", min_value=0.0)
                if st.form_submit_button("Save"):
                    if not name or name in st.session_state.employees:
                        st.error("Invalid or duplicate name")
                    else:
                        max_id = max((int(data['employee_id'].replace('EMP', '')) for data in st.session_state.employees.values()), default=0) + 1
                        employee_id = f"EMP{max_id:03d}"
                        st.session_state.employees[name] = {
                            "employee_id": employee_id, "email": email, "mobile": mobile, "designation": designation,
                            "bank_name": bank_name, "account_number": account_number, "ifsc": ifsc, "monthly_salary": monthly_salary
                        }
                        save_employees(st.session_state.employees)
                        st.success("Employee added!")
                        st.rerun()

        with st.expander("Modify or Delete Employee", expanded=False):
            selected_emp = st.selectbox("Select Employee to Modify", options=[f"{data['employee_id']} - {name}" for name, data in st.session_state.employees.items()])
            if selected_emp:
                emp_id = selected_emp.split(" - ")[0]
                old_name = next(name for name, data in st.session_state.employees.items() if data["employee_id"] == emp_id)
                with st.form("Modify Employee"):
                    new_name = st.text_input("Name", value=old_name)
                    email = st.text_input("Email", value=st.session_state.employees[old_name]["email"])
                    mobile = st.text_input("Mobile", value=st.session_state.employees[old_name]["mobile"])
                    designation = st.text_input("Designation", value=st.session_state.employees[old_name]["designation"])
                    bank_name = st.text_input("Bank Name", value=st.session_state.employees[old_name]["bank_name"])
                    account_number = st.text_input("Account Number", value=st.session_state.employees[old_name]["account_number"])
                    ifsc = st.text_input("IFSC", value=st.session_state.employees[old_name]["ifsc"])
                    monthly_salary = st.number_input("Monthly Salary", value=st.session_state.employees[old_name]["monthly_salary"])
                    if st.form_submit_button("Save"):
                        if not new_name:
                            st.error("Name cannot be empty")
                        elif new_name != old_name and new_name in st.session_state.employees:
                            st.error("Employee name already exists")
                        else:
                            updated_data = {
                                "employee_id": emp_id, "email": email, "mobile": mobile, "designation": designation,
                                "bank_name": bank_name, "account_number": account_number, "ifsc": ifsc, "monthly_salary": monthly_salary
                            }
                            if new_name != old_name:
                                st.session_state.employees[new_name] = updated_data
                                del st.session_state.employees[old_name]
                                if emp_id in st.session_state.attendance["Employee ID"]:
                                    st.session_state.attendance["Employee ID"][emp_id]["name"] = new_name
                            else:
                                st.session_state.employees[new_name].update(updated_data)
                            save_employees(st.session_state.employees)
                            save_attendance(st.session_state.attendance)
                            st.success("Employee modified!")
                            st.rerun()

            selected_del = st.selectbox("Select Employee to Delete", options=[f"{data['employee_id']} - {name}" for name, data in st.session_state.employees.items()])
            if selected_del and st.button("Delete Employee"):
                emp_id = selected_del.split(" - ")[0]
                name = selected_del.split(" - ")[1]
                del st.session_state.employees[name]
                if emp_id in st.session_state.attendance["Employee ID"]:
                    del st.session_state.attendance["Employee ID"][emp_id]
                save_employees(st.session_state.employees)
                save_attendance(st.session_state.attendance)
                st.success("Employee deleted!")
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

elif selected_tab == "Attendance Search":
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.header("Attendance Search")
        emp_id_name = st.selectbox("Employee ID", options=[f"{data['employee_id']} - {name}" for name, data in st.session_state.employees.items()])
        if st.button("Search"):
            if emp_id_name:
                emp_id = emp_id_name.split(" - ")[0]
                if emp_id in st.session_state.attendance["Employee ID"]:
                    att_data = st.session_state.attendance["Employee ID"][emp_id]["date"]
                    att_list = [{"Date": date, "Day": data["Day"], "In Time": data["In Time"], "Out Time": data["Out Time"],
                                 "Total Hours": data["Total hours"], "Status": data["Status"], "Salary": data["Salary"],
                                 "Remark": data["Remark"]} for date, data in att_data.items()]
                    att_df = pd.DataFrame(att_list)

                    def color_status(val):
                        color = {'Full day': 'background-color: lightgreen', 'Half day': 'background-color: lightyellow',
                                 'Absent': 'background-color: lightcoral', 'WFH': 'background-color: lightblue'}.get(val, '')
                        return color

                    styled_df = att_df.style.applymap(color_status, subset=['Status'])
                    st.dataframe(styled_df, use_container_width=True)
                    st.write(f"Total Salary: {st.session_state.attendance['Employee ID'][emp_id].get('total_salary', 0)}")

        if emp_id_name:
            with st.expander("Update Status for Selected Employee"):
                emp_id = emp_id_name.split(" - ")[0]
                emp_name = st.session_state.attendance["Employee ID"][emp_id]["name"] if emp_id in st.session_state.attendance["Employee ID"] else ""
                date_options = list(st.session_state.attendance["Employee ID"][emp_id]["date"].keys()) if emp_id in st.session_state.attendance["Employee ID"] else []
                selected_date = st.selectbox("Select Date to Update", options=date_options)
                if selected_date:
                    current_data = st.session_state.attendance["Employee ID"][emp_id]["date"][selected_date]
                    with st.form("Update Status"):
                        status = st.selectbox("Status", options=["Full day", "Half day", "Absent", "WFH"], index=["Full day", "Half day", "Absent", "WFH"].index(current_data["Status"]))
                        remark = st.text_input("Remark", value=current_data["Remark"])
                        if st.form_submit_button("Save"):
                            if not remark:
                                st.error("Remark is required")
                            else:
                                days_in_month = calendar.monthrange(2025, 7)[1]
                                daily_salary = st.session_state.employees.get(emp_name, {}).get("monthly_salary", 0) / days_in_month
                                new_salary = calculate_salary(status, daily_salary)
                                old_salary = current_data["Salary"]
                                st.session_state.attendance["Employee ID"][emp_id]["date"][selected_date].update({
                                    "Status": status, "Salary": new_salary, "Remark": remark
                                })
                                st.session_state.attendance["Employee ID"][emp_id]["total_salary"] += (new_salary - old_salary)
                                save_attendance(st.session_state.attendance)
                                st.success("Status updated!")
                                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

elif selected_tab == "Reports":
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.header("Reports")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Generate Attendance Excel"):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(["Employee ID", "Employee Name", "Date", "Day", "In Time", "Out Time", "Total Hours", "Status", "Salary", "Remark", "Total Salary"])
                for emp_id, data in st.session_state.attendance["Employee ID"].items():
                    total_salary = data.get("total_salary", 0)
                    for date, att in data["date"].items():
                        ws.append([emp_id, data["name"], date, att["Day"], att["In Time"], att["Out Time"],
                                   att["Total hours"], att["Status"], att["Salary"], att["Remark"], ""])
                    ws.append([emp_id, data["name"], "", "", "", "", "", "", "", "", total_salary])
                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                st.download_button("Download Attendance Report", buffer, file_name="attendance_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with col2:
            with st.form("Payment File Options"):
                st.subheader("Generate Payment File")
                trans_type = st.selectbox("Transaction Type", options=["NEFT", "RTGS"], index=0)
                debit_acc = st.text_input("Debit Account Number")
                use_current = st.checkbox("Use Current Date", value=True)
                trans_date = st.text_input("Transaction Date (DD/MM/YYYY)", value=datetime.now().strftime("%d/%m/%Y")) if not use_current else datetime.now().strftime("%d/%m/%Y")
                remark = st.text_input("Remark")
                if st.form_submit_button("Generate"):
                    if not debit_acc:
                        st.error("Debit account number required")
                    else:
                        date_str = datetime.now().strftime("%d/%m/%Y") if use_current else trans_date
                        try:
                            pd.to_datetime(date_str, format="%d/%m/%Y")
                        except:
                            st.error("Invalid date format (DD/MM/YYYY)")
                        else:
                            wb = openpyxl.Workbook()
                            ws = wb.active
                            headers = ["Beneficiary Name", "Beneficiary Account Number", "IFSC", "Transaction Type",
                                       "Debit Account Number", "Transaction Date", "Amount", "Currency",
                                       "Beneficiary Email ID", "Remarks", "Custom Header – 1", "Custom Header – 2",
                                       "Custom Header – 3", "Custom Header – 4", "Custom Header – 5"]
                            ws.append(headers)
                            ws.append(["Enter beneficiary name. MANDATORY", "Enter beneficiary account number. MANDATORY",
                                       "Enter beneficiary bank IFSC code.", "Enter payment type: IFT/NEFT/RTGS",
                                       "Enter debit account number.", "Enter transaction value date. DD/MM/YYYY",
                                       "Enter payment amount. MANDATORY", "Enter transaction currency. INR",
                                       "Enter beneficiary email id OPTIONAL", "Enter remarks OPTIONAL",
                                       "Credit Advice: Custom Info -1", "Credit Advice: Custom Info -2",
                                       "Credit Advice: Custom Info -3", "Credit Advice: Custom Info -4",
                                       "Credit Advice: Custom Info -5"])
                            for emp_id, data in st.session_state.attendance["Employee ID"].items():
                                emp_name = data["name"]
                                emp_data = next((e for n, e in st.session_state.employees.items() if e["employee_id"] == emp_id), {})
                                ws.append([emp_name, emp_data.get("account_number", ""), emp_data.get("ifsc", ""),
                                           trans_type, debit_acc, date_str, data.get("total_salary", 0),
                                           "INR", emp_data.get("email", ""), remark, emp_id, "", "", "", ""])
                            buffer = io.BytesIO()
                            wb.save(buffer)
                            buffer.seek(0)
                            filename = "BLKPAY_{}.xlsx".format(datetime.now().strftime("%Y%m%d"))
                            st.download_button("Download Payment File", buffer, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.markdown('</div>', unsafe_allow_html=True)