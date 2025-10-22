import logging
import math
# import time
import pandas as pd
import json 
from logger import logger 
from pathlib import Path
from datetime import datetime, timedelta, time
import os
from utils import calculate_ot_ut, get_db_connection, get_empl_working_hour, update_attendance_status, analysing_att_status



def fetch_attendance(date):
    # Requesting a db connection
    conn, cursor = get_db_connection()
    
    if conn is None or cursor is None:
        fetch_attendance()
        return
    
    query = f"""
        SELECT 
            a.In_time,
            a.Out_time,
            a.Emp_id,
            a.Att_month,
            e.employee_code,
            e.employee_fname
        FROM 
            FinalDay_Attendance AS a
        INNER JOIN 
            employees AS e
        ON 
            a.Emp_id = e.employee_id
        WHERE 
            a.Att_month = #{date}#;
    """

    try:
        logger.info("Executing attendance query...")
        cursor.execute(query)

        # Fetch column names
        columns = [column[0] for column in cursor.description]
        logger.info(f"Columns fetched: {columns}")

        # Fetch rows
        rows = cursor.fetchall()

        # Convert to list of dicts
        data = [dict(zip(columns, row)) for row in rows]

        # Convert to JSON
        attd_data = json.dumps(data, default=str, indent=4)
        logger.info("Attendance data converted to JSON.")
        
        # Write to file
        output_file = "attendance.json"
        with open(output_file, "w") as f:
            f.write(attd_data)

    except Exception as e:
        logger.error(f"Error while fetching attendance data: {e}")
        attd_data = None

    finally:
        cursor.close()
        conn.close()
        logger.info("Database connection closed.")

    return data


def fetching_report(date, grace_min=20):
    """
    Generate a report of late arrivals, early leavers, overtimers, and missing attendance.
    shift_start, shift_end are hours in 24-hour format (int or float)
    """
    late_arrival = []
    early_leave = []
    overtimers = []
    attendance_miss = []

    # Fetch attendance data
    attnd_data = fetch_attendance(date)

    if attnd_data == [] or len(attnd_data) == 0:
        logger.info("No Data found (Press Enter to close): ")
        input("")
    
    if not attnd_data:
        logging.warning("No attendance data found.")
        return

    for entry in attnd_data:
        emp_id = entry["employee_code"]
        emp_name = entry["employee_fname"]
        in_time = entry["In_time"]
        out_time = entry["Out_time"]
        att_date = entry["Att_month"]  # assumed full datetime
        
        # --- Finding users working hours from json ---
        shift_start = 8
        shift_hours, sunday_duty = get_empl_working_hour(emp_id)
        

        # --- Handle missing punches ---
        if not in_time or not out_time:
            logging.warning(f"Error encounter for {emp_id}: There is no record of arriving and leave for {emp_name}")
            continue

        # --- Convert strings to datetime safely ---
        try:
            if isinstance(in_time, str):
                in_time = datetime.fromisoformat(in_time)
            if isinstance(out_time, str):
                out_time = datetime.fromisoformat(out_time)
            if isinstance(att_date, str):
                att_date = datetime.fromisoformat(att_date)
        except Exception as e:
            logging.error(f"Time format error for {emp_id} - {emp_name}: {e}")
            continue

        
        shift_start_time = datetime.combine(att_date.date(), time(int(shift_start), 0))
        shift_end_time = datetime.combine(att_date.date(), time(int(shift_start), 0)) + timedelta(hours=shift_hours)
        working_hour = math.floor(((out_time - in_time).total_seconds() / 3600) * 2) / 2

        logger.info(f"<< Report analysing for {emp_id} | {emp_name} >>")
        
        if in_time.strftime("%H:%M") == "00:00" and out_time.strftime("%H:%M") == "00:00":
            logger.info(f"Employee with id {emp_id} is absent")
            continue # skipping those who are absent
        
        # --- Check for missing attendance ---
        if out_time.strftime("%H:%M") == "00:00" or in_time.strftime("%H:%M") == out_time.strftime("%H:%M"):
            attendance_miss.append({
                "Sr No.":len(attendance_miss) + 1,
                "Code": int(emp_id),
                "Employee": emp_name,
                "Shift Hours": shift_hours,
                "In time": in_time.strftime("%H:%M"),
                "Out time": out_time.strftime("%H:%M"),
                "Working Hour": working_hour,
                "Reason": "MIS"
            })
            logger.info(f"Missing Attendance: In Time = 09:05, Out Time = {out_time}")  
            continue
            

        # --- Check late arrival ---
        if in_time > shift_start_time + timedelta(minutes=grace_min):  # 5 min grace
            delay = (in_time - shift_start_time).seconds / 60
            late_arrival.append({
                "Sr No.":len(late_arrival) + 1,
                "Code": int(emp_id),
                "Employee": emp_name,
                "Shift Hours": shift_hours,
                "In time": in_time.strftime("%H:%M"),
                "Out time": out_time.strftime("%H:%M"),
                "Working Hour": working_hour,
                "Late By": timedelta(minutes=delay)
            })
            logger.info(f"Later arrival: In Time = 09:05, Out Time = {out_time}, Late by = {timedelta(minutes=delay)}")  
            

        # --- Check early leave ---
        if out_time < shift_end_time - timedelta(minutes=grace_min):  # 5 min grace
            early = (shift_end_time - out_time).seconds / 60
            early_leave.append({
                "Sr No.":len(early_leave) + 1,
                "Code": int(emp_id),
                "Employee": emp_name,
                "Shift Hours": shift_hours,
                "In time": in_time.strftime("%H:%M"),
                "Out time": out_time.strftime("%H:%M"),
                "Working Hour": working_hour,
                "Left Early": timedelta(minutes=early)
            })
            logger.info(f"Left Early: In Time = 09:05, Out Time = {out_time}, Left Early by = {timedelta(minutes=early)}")  
            

        # --- Check overtime or undertime ---
        ot_ut = calculate_ot_ut(in_time, out_time, att_date.strftime("%a"), sunday_duty, shift_hours, grace_min)
        if ot_ut not in (0, 0.0, " ") and ot_ut > 0:
            overtimers.append({
                "Sr No.":len(overtimers) + 1,
                "Code": int(emp_id),
                "Employee": emp_name,
                "Shift Hours": shift_hours,
                "In time": in_time.strftime("%H:%M"),
                "Out time": out_time.strftime("%H:%M"),
                "Working Hour": working_hour,
                "Overtime": ot_ut
            })
            logger.info(f"Over/under: In Time = 09:05, Out Time = {out_time}, Overtime/Undertime = {ot_ut} hrs")  
            

    # --- Combine report ---
    report = {
        "report_date":  att_date,
        "late_arrival": late_arrival,
        "left_earlie": early_leave,
        "overtimers": overtimers,
        "missing_attendance": attendance_miss
    }
    
    logger.info("Serializing attendance data structure to JSON and persisting to disk.")
    with open("attendance_report.json", "w") as f:
        to_write = json.dumps(report, default=str)
        f.write(to_write)
        

    return report

    

def fetch_month_attendance(month):
    # Requesting a db connection
    conn, cursor = get_db_connection()
    
    if conn is None or cursor is None:
        fetch_attendance()
        return
    
    query = f"""
    SELECT 
        a.In_time,
        a.Out_time,
        a.Emp_id,
        a.Att_month,
        e.employee_code,
        e.employee_fname
    FROM 
        FinalDay_Attendance AS a
    INNER JOIN 
        employees AS e
        ON a.Emp_id = e.employee_id
    WHERE 
        MONTH(a.Att_month) = MONTH(#{month}#)
        AND YEAR(a.Att_month) = YEAR(#{month}#);
    """


    try:
        logger.info("Executing attendance query...")
        cursor.execute(query)

        # Fetch column names
        columns = [column[0] for column in cursor.description]
        logger.info(f"Columns fetched: {columns}")

        # Fetch rows
        rows = cursor.fetchall()

        # Convert to list of dicts
        data = [dict(zip(columns, row)) for row in rows]

        # Convert to JSON
        attd_data = json.dumps(data, default=str, indent=4)
        logger.info("Attendance data converted to JSON.")
        
        # Write to file
        output_file = "attendance.json"
        with open(output_file, "w") as f:
            f.write(attd_data)

    except Exception as e:
        logger.error(f"Error while fetching attendance data: {e}")
        attd_data = None
        
    finally:
        cursor.close()
        conn.close()
        logger.info("Database connection closed.")

    return data





def update_employee(employee_code: int, att_date: datetime):
    """Update employee attendance intelligently (A/P/Miss transitions)."""
    try:
        conn, cursor = get_db_connection()

        # Fetch existing record
        cursor.execute(f"""
            SELECT 
                a.In_time, a.Out_time, a.Tot_Min,
                e.employee_fname, e.employee_code
            FROM FinalDay_Attendance AS a
            INNER JOIN employees AS e 
            ON a.Emp_id = e.employee_id
            WHERE e.employee_code = ? AND a.Att_month = ?
        """, (employee_code, att_date))
        record = cursor.fetchone()

        if not record:
            logger.info(f"No record found for Employee {employee_code} on {att_date.date()}")
            return
        
        emp_name = record[3]
        in_time, out_time = record[0], record[1]
        logger.info(f"Employee: {emp_name} ({employee_code})")
        logger.info(f"Current -> In: {record[0]} | Out: {record[1]} | OT: {calculate_ot_ut(in_time, out_time, in_time.strftime("%a"))} | Total(m): {record[2]}")

        # Ask what to update
        print("\nWhat would you like to update?")
        print("1) Attendance status (P/A)")
        print("2) OT/UT value (float: +ve=OT, -ve=UT, 0=None)")
        choice = input("Select option (1/2): ").strip()

        if choice == "1": 
            # === Identify current status ===
            current_status = analysing_att_status(in_time, out_time)
            
            # === Get desired update ===
            new_status = input("Enter new attendance status (A/P): ").strip().upper()
            if new_status not in ("A", "P"):
                logger.error("❌ Invalid input. Only 'A' or 'P' allowed.")
                return

            update_data = update_attendance_status(current_status, new_status, att_date, in_time)
            if update_data == None:
                return
            
            updated_in, updated_out, updated_tot = update_data
            
            # === Update DB ===
            cursor.execute("""
                UPDATE FinalDay_Attendance 
                SET In_time = ?, Out_time = ?, Tot_Min = ?
                WHERE Card_Number = ? AND
                Att_month = ?
            """, (
                updated_in.strftime("%Y-%m-%d %H:%M:%S"),
                updated_out.strftime("%Y-%m-%d %H:%M:%S"),
                updated_tot,
                employee_code,
                att_date
            ))
            conn.commit()
            logger.info("✅ Attendance status updated in success.")
            
        elif choice == "2":
            ot_ut_val = float(input("Enter OT/UT value in hour (positive=OT, negative=UT, 0=None): "))
            
            current_status = analysing_att_status(in_time, out_time)
            if current_status != "P":
                logger.error(f"employe - {emp_name} | {employee_code} is absent or missed attendance for {att_date}")
                return
            
            out_time += timedelta(hours=ot_ut_val)
            tot_min = int((out_time - in_time).total_seconds() / 60)
            
            cursor.execute("""
            UPDATE FinalDay_Attendance 
            SET Out_time = ?, 
                ot_minute = ?,
                Tot_Min = ?,
                shift_late_minute = 0,
                early_dep_minute = 0
            WHERE Card_Number = ?
            AND Att_month = ?
            """, (out_time.strftime("%Y-%m-%d %H:%M:%S"), max(0, ot_ut_val), tot_min, employee_code, att_date))
            conn.commit()
            logger.info("✅ OT/UT updated successfully.")
        
        else:
            logger.error(f"There is no option for {choice}")
            return 

        # === Fetch updated record for confirmation ===
        cursor.execute("""
            SELECT 
                e.employee_code, e.employee_fname, a.In_time, a.Out_time, a.ot_minute, 
                a.shift_late_minute, a.early_dep_minute, a.Tot_Min
            FROM FinalDay_Attendance AS a
            INNER JOIN employees AS e ON a.Emp_id = e.employee_id
            WHERE e.employee_code = ? AND a.Att_month = ?
        """, (employee_code, att_date))

        updated = cursor.fetchone()
        if updated:
            logger.info("\nUpdated Record:")
            logger.info(f"Employee: {updated[1]} ({updated[0]})")
            logger.info(f"In: {updated[2]} | Out: {updated[3]} | OT: {updated[4]} | Late: {updated[5]} | Early: {updated[6]} | Total: {updated[7]}")

        conn.close()

    except Exception as e:
        print(f"❌ Error updating record: {e}")


