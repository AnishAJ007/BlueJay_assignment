import pandas as pd
from datetime import timedelta

# Load the spreadsheet
file_path = 'C:/Users/anish/Downloads/Assignment_Timecard.xlsx'  # Replace with the actual path to your spreadsheet
df = pd.read_excel(file_path)

print("Column Names:", df.columns)

# Function to check consecutive days worked
def consecutive_days(employee_data):
    output = []
    consecutive_count = 0
    for i in range(1, len(employee_data)):
        if (employee_data['Pay Cycle End Date'][i] - employee_data['Pay Cycle End Date'][i - 1]).days == 1:
            consecutive_count += 1
        else:
            consecutive_count = 0

        if consecutive_count == 6:
            output.append(f"{employee_data['Employee Name'].iloc[i]} has worked for 7 consecutive days at position {employee_data['Position ID'].iloc[i]}")
    return output

# Function to check time between shifts
def time_between_shifts(employee_data):
    output = []
    for i in range(1, len(employee_data)):
        time_diff = employee_data['Time'][i] - employee_data['Time Out'][i - 1]
        if timedelta(hours=1) < time_diff < timedelta(hours=10):
            output.append(f"{employee_data['Employee Name'].iloc[i]} has less than 10 hours between shifts at position {employee_data['Position ID'].iloc[i]}")
    return output

# Function to check hours worked in a single shift
# Function to check hours worked in a single shift
# Function to check hours worked in a single shift
# Function to check hours worked in a single shift
def hours_worked_single_shift(employee_data):
    output = []
    for i in range(len(employee_data)):
        # Get the 'Timecard Hours (as Time)' value
        time_str = str(employee_data['Timecard Hours (as Time)'][i])

        # Check if the value is already numeric
        if '.' in time_str or ':' in time_str:
            # Convert the 'Timecard Hours (as Time)' value to numerical format
            hours, minutes = map(int, time_str.split(':'))
            hours_worked = hours + minutes / 60.0

            # Compare the numerical value with 14
            if hours_worked > 14:
                output.append(f"{employee_data['Employee Name'].iloc[i]} has worked for more than 14 hours in a single shift at position {employee_data['Position ID'].iloc[i]}")
        else:
            # Handle the case where the value is already numeric
            hours_worked = float(time_str)
            if hours_worked > 14:
                output.append(f"{employee_data['Employee Name'].iloc[i]} has worked for more than 14 hours in a single shift at position {employee_data['Position ID'].iloc[i]}")

    return output



# Write output to file and print to console
def write_output_to_file(output):
    with open('output.txt', 'a') as file:
        for line in output:
            file.write(line + '\n')
            print(line)

# Apply the functions to the dataframe
consecutive_output = consecutive_days(df)
time_between_output = time_between_shifts(df)
hours_worked_output = hours_worked_single_shift(df)

# Combine all outputs
all_output = consecutive_output + time_between_output + hours_worked_output

# Write to output.txt and print to console
write_output_to_file(all_output)
