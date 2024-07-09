import datetime
from openpyxl import load_workbook

def get_current_time():
    return datetime.datetime.now()
 
def increment_time_by_one_minute(time_obj):
    return time_obj + datetime.timedelta(minutes=1)
 
def write_to_excel(filename):
    wb = load_workbook(filename)
    sheet = wb.active
    current_time = get_current_time()
    current_time_str = current_time.strftime("%#I:%M")
 
    incremented_time = increment_time_by_one_minute(current_time)
    incremented_time_str = incremented_time.strftime("%#I:%M")
    
    sheet["C104"] = current_time_str
    sheet["C105"] = incremented_time_str
 
    wb.save(filename)
    print(f"{current_time_str} and {incremented_time_str}")
 
if __name__ == "__main__":
    filename = "Bench_Configuration.xlsx"
    write_to_excel(filename)