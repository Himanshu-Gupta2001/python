import pandas as pd

def analyze_excel_file(file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Assume the columns in the Excel file are as follows
    employee_name_col = 'Employee Name'
    position_col = 'Position ID'
    time_in_col = 'Time'
    time_out_col = 'Time Out'
    timecard_hours_col = 'Timecard Hours (as Time)'

    # Convert time columns to datetime format for easier manipulation
    df[time_in_col] = pd.to_datetime(df[time_in_col])
    df[time_out_col] = pd.to_datetime(df[time_out_col])

    # a) Employees who have worked for 7 consecutive days
    consecutive_days_threshold = 7
    consecutive_days_mask = df.groupby(employee_name_col)[time_in_col].diff().dt.days == 1
    consecutive_days_employees = df[consecutive_days_mask].groupby(employee_name_col).filter(lambda x: len(x) >= consecutive_days_threshold)

    # b) Employees with less than 10 hours between shifts but greater than 1 hour
    time_between_shifts_min = pd.Timedelta(hours=1)
    time_between_shifts_max = pd.Timedelta(hours=10)
    time_between_shifts_mask = (df[time_in_col].shift(-1) - df[time_out_col]).between(time_between_shifts_min, time_between_shifts_max)
    time_between_shifts_employees = df[time_between_shifts_mask]

    # c) Employees who have worked for more than 14 hours in a single shift
    max_single_shift_hours = 14
    single_shift_mask = (df[time_out_col] - df[time_in_col]).dt.total_seconds() / 3600 > max_single_shift_hours
    long_shift_employees = df[single_shift_mask]

    # Print the results
    print("Employees who have worked for 7 consecutive days:")
    print(consecutive_days_employees[[employee_name_col, position_col]])

    print("\nEmployees with less than 10 hours between shifts but greater than 1 hour:")
    print(time_between_shifts_employees[[employee_name_col, position_col]])

    print("\nEmployees who have worked for more than 14 hours in a single shift:")
    print(long_shift_employees[[employee_name_col, position_col]])

if __name__ == "__main__":
    # Provide the path to your Excel file
    excel_file_path = "C:/Users/gupta/Downloads/Assignment_Timecard.xlsx"
    analyze_excel_file(excel_file_path)
