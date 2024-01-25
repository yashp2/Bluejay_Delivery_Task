import pandas as pd
from datetime import timedelta

def analyze_excel_file(file_path):
# Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path)

# Assuming 'Time In' and 'Time Out' columns are in datetime format
    df['Time In'] = pd.to_datetime(df['Time'])
    df['Time Out'] = pd.to_datetime(df['Time Out'])
# Sort the DataFrame by 'Employee Name' and 'in Time'
    df = df.sort_values(by=['Employee Name', 'Time In'])
# Iterate through each employee's records
    for name, records in df.groupby('Employee Name'):
        consecutive_days = 1
        prev_time_out = None

        for index, row in records.iterrows():
            # Check for consecutive days
            if prev_time_out is not None and (row['Time In'] - prev_time_out).days == 1:
                consecutive_days += 1
            else:
                consecutive_days = 1

            # Checking for less than 10 hours between shifts but greater than 1 hour
            time_between_shifts = (row['Time In'] - prev_time_out) if prev_time_out else timedelta(days=0)
            if 1 < time_between_shifts.total_seconds() // 3600 < 10:
                print(f"{name} ({row['Position ID']}) has less than 10 hours between shifts on {row['Time In']}")

            # Checking for more than 14 hours in a single shift
            if (row['Time Out'] - row['Time In']).total_seconds() // 3600 > 14:
                print(f"{name} ({row['Position ID']}) has worked more than 14 hours on {row['Time In']}")

            # Checking for 7 consecutive days
            if consecutive_days == 7:
                print(f"{name} ({row['Position ID']}) has worked for 7 consecutive days starting from {row['Time In']}")

            prev_time_out = row['Time Out']

# Assuming the file is in the same directory as the script
file_path = 'Assignment_Timecard.xlsx'
analyze_excel_file(file_path)
