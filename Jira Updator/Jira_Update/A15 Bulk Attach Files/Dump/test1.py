import os
import pandas as pd
from datetime import datetime

def format_date(date_string):
  try:
    date_object = datetime.strptime(date_string, '%Y-%m-%d %H:%M:%S')
    formatted_date = date_object.strftime('%a, %d %b %Y %H:%M:%S +0000')
    return formatted_date
  except ValueError:
    return "Invalid date format. Please use YYYY-MM-DD HH:MM:SS"

EXCEL_FILE = os.path.join("input", "input.xlsx")
if not os.path.isfile(EXCEL_FILE):
   raise SystemExit(f"Input Excel not found: {EXCEL_FILE}")
df = pd.read_excel(EXCEL_FILE)
for _, row in df.iterrows():
    ts_modifications_on = str(row.get('TS Modifications On', "")).strip()
    format_date(ts_modifications_on)

abc = format_date(ts_modifications_on)
print(abc)