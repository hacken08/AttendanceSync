from datetime import datetime
import pandas as pd
import json 
from logger import logger 
from fetcher import fetch_attendance, fetch_month_attendance, update_employee
from utils import get_valid_date, get_valid_month
from writer import write_to_excel
import sys

def get_excel_path():
    #  ====== User inputs ======
    excile_path = input("Give file path of excel to write: ")

    if excile_path.startswith("\"") and excile_path.endswith("\""):
        excile_path = excile_path[1:len(excile_path)-1]
        
    print("excile_path", excile_path)

if __name__ == "__main__":
    try:
        menu_option = int(input("""
1). Mark attendance for a day.
2). Mark attendance for a month.
3). Update employee data.
Select from Above Menu:  """))
        
        match menu_option:
            case 1:
                date = get_valid_date()
                attd_data = fetch_attendance(date)
                
                if attd_data == [] or len(attd_data) == 0:
                    logger.info("No Data found (Press Enter to close): ")
                    input("")
                else:
                    excile_path = get_excel_path()
                    write_to_excel(attd_data, excile_path)
                    input("Attendance marking is complete (Press Enter to close): ")
                    
            case 2:
                month = get_valid_month()
                attd_data = fetch_month_attendance(month)
                if attd_data == [] or len(attd_data) == 0:
                    logger.info("No Data found (Press Enter to close): ")
                    input("")
                else:
                    excile_path = get_excel_path()
                    write_to_excel(attd_data, excile_path)
                    input("Attendance marking is complete (Press Enter to close): ")
                    
            case 3:
                employee_code = int(input("Enter employee id: "))
                date = datetime.strptime("10/12/2025", "%m/%d/%Y")
                # employee_code = int(input("Enter employee code: "))
                # date = get_valid_date()
                update_employee(employee_code, date)
            
            case _:
                logger.error("Inavlid input try again\n")
                    
        
    except Exception as e:
        logger.error(e)
        input("Press Enter to close")