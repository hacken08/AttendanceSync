from datetime import datetime
from io import BytesIO
import json
import logging
import math
import os
import sys
import openpyxl
import psutil
import requests
from logger import logger
import pyodbc


config = {
    "db_path": r"C:\Program Files (x86)\ONtime\ACCESSDB\ontime_att.mdb",
    # "db_path": r"D:\auto_attendancer\data\ontime_att.mdb",
    "db_password": "sss",
}


def get_db_connection():
    try:
        logger.info(f"Connecting to Access database: {config["db_path"]}")
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            rf'DBQ={config["db_path"]};'
            rf'PWD={config["db_password"]};'
        )
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        logger.info("Database connection successful.")
        return conn, cursor
    except Exception as e:
        logger.error(f"Failed to connect to database: {e}")
        db_path = input("Could Not Find DB: Please Provide Database Path:-")
        config["db_path"] = os.path.normpath(db_path)
        return None, None



def do_something_useful():
    print("Replace this with a utility function")


def close_excel_if_open(file_path):
    """Close Excel if this workbook is open."""
    file_name = os.path.basename(file_path).lower()
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'excel' in proc.info['name'].lower():
            try:
                # Check if file is locked by this process
                open_files = proc.open_files()
                for f in open_files:
                    if file_name in f.path.lower():
                        print(f"Closing Excel process (PID {proc.pid}) holding {file_name}")
                        proc.terminate()
                        proc.wait(timeout=3)
                        return True
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                continue
    return False


def open_excel(file_path):
    """Open Excel file for the user."""
    os.startfile(file_path)




def load_excel(excile_path: str):
    """
    Load Excel from either local path or online URL.
    Returns workbook object.
    """
    try :
        if excile_path.startswith("http://") or excile_path.startswith("https://"):
            print("[INFO] Loading online Excel file:", excile_path)
            response = requests.get(excile_path)
            print("Content-Type:", response.headers.get("Content-Type"))
            
            response.raise_for_status()
            wb = openpyxl.load_workbook(filename=BytesIO(response.content))
            
        elif os.path.exists(excile_path):
            print("[INFO] Loading local Excel file:", excile_path)
            wb = openpyxl.load_workbook(filename=excile_path)

        else:
            raise FileNotFoundError(f"Excel source not found: {excile_path}")
        return wb
    except Exception as e:
        logging.error("Enable to open exile")
        return None
        



def round_to_half_hour(hours):
    return round(hours * 2) / 2  # rounds to nearest 0.5


def calculate_ot_ut(in_time, out_time, day, sunday_duty=False, standard_hours=10, grace_minutes=20):
    """
    Calculate the working hour undertime and overtime of employees.
    """
    # Total worked hours
    total_hours = (out_time - in_time).total_seconds() / 3600

    # Grace in hours
    grace = grace_minutes / 60

    # Calculate OT/UT
    if day.lower() == "sun" and sunday_duty == False:
        return math.floor(total_hours * 2) / 2 
    elif total_hours > standard_hours + grace:
        diff = total_hours - standard_hours
        value = math.floor(diff * 2) / 2   # round down to nearest 0.5h
        return value
    elif total_hours < standard_hours - grace:
        diff = standard_hours - total_hours
        value = math.floor(diff * 2) / 2   # round down to nearest 0.5h
        return -value
    else:
        return " "


from typing import Tuple

def get_empl_working_hour(emp_code) -> Tuple[int, bool]:
    try:
        base_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
        shift_hour_json = os.path.join(base_dir, "shift_hour.json")

        with open(shift_hour_json, "r", encoding="utf-8") as r:
            # logger.info("Opened shift_hour.json file")
            data = json.load(r)

            for empl in data:
                if empl["employee_code"] == emp_code:
                    return empl["working_hours"], empl["sunday_duty"]

            return 10, False

    except Exception as e:
        logger.error("Unable to find shift_hour.json")
        sys.exit()
        return 10, False



# --- Validate date format before query ---
def get_valid_date():
    while True:
        try:
            user_input = input("Enter Date (mm/dd/yyyy): ").strip()
            valid_date = datetime.strptime(user_input, "%m/%d/%Y")
            return valid_date.strftime("%m/%d/%Y") 
        except Exception:
            logger.error("❌ Invalid date format. Please use mm/dd/yyyy (e.g. 10/08/2025).")
            
            
# --- Validate date format before query ---
def get_valid_month():
    while True:
        try:
            user_input = input("Enter Date (mm/yyyy): ").strip()
            valid_date = datetime.strptime(user_input, "%m/%Y")
            return valid_date.strftime("%m/%Y") 
        except Exception:
            logger.error("❌ Invalid date format. Please use mm/yyyy (e.g. 10/2025).")



from datetime import datetime, timedelta

def update_attendance_status(current_status: str, new_status: str, att_date: datetime, in_time: datetime | None):
    """
    Returns updated (in_time, out_time, total_minutes) tuple based on transition.
    Automatically handles all valid state changes with minimal branching.
    """
    # --- Define all transitions in one map ---
    SHIFT_HOURS = 10
    shift_in = att_date.replace(hour=8, minute=0, second=0)
    shift_out = att_date.replace(hour=18, minute=0, second=0)
    zero_time = att_date.replace(hour=0, minute=0, second=0)

    # Rule table: (current, new) -> (in_time, out_time, total_minutes)
    transitions = {
        ("A", "P"): (shift_in, shift_out, SHIFT_HOURS * 60),
        ("P", "A"): (zero_time, zero_time, 0),
        ("M", "P"): (in_time, in_time + timedelta(hours=SHIFT_HOURS) if in_time else shift_out, SHIFT_HOURS * 60),
        ("M", "A"): (zero_time, zero_time, 0),
    }

    # --- Same status (no change) ---
    if current_status == new_status:
        print("ℹ️ No change needed — already same status.")
        return None  # indicates skip

    # --- If transition not defined ---
    if (current_status, new_status) not in transitions:
        print(f"⚠️ Unsupported transition: {current_status} → {new_status}")
        return None

    # --- Safe unpacking ---
    updated_in, updated_out, updated_total = transitions[(current_status, new_status)]
    return updated_in, updated_out, updated_total




def analysing_att_status(in_t: datetime, out_t: datetime):
    current_status = ""
    if (in_t.hour == 0 and out_t.hour == 0):
       current_status = "A"
    elif in_t == out_t or (in_t.hour != 0 and out_t.hour == 0):
        current_status = "M"
    else:
        current_status =  "P"

    
    return current_status