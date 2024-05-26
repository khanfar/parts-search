import pandas as pd
import requests
import re
from tkinter import Tk, Label, Entry, Button, StringVar, messagebox
from tkinter import ttk
from openpyxl import load_workbook

# Define the URL of the CSV file
url = "https://huggingface.co/spaces/mhammad/Khanfar/raw/main/docs/your_data.csv"

# Fetch the CSV file
response = requests.get(url)
with open("your_data.csv", "wb") as f:
    f.write(response.content)

# Load the CSV file into a pandas DataFrame
df = pd.read_csv("your_data.csv")

def generate_report(start_date, end_date):
    # Convert date columns to datetime
    df['تاريخ الدخول'] = pd.to_datetime(df['تاريخ الدخول'], format='%d.%m.%Y', errors='coerce')

    # Filter data for the given period
    mask = (df['تاريخ الدخول'] >= start_date) & (df['تاريخ الدخول'] <= end_date)
    filtered_df = df.loc[mask]

    # Extract rows where 'تقرير نهائي' contains the '#' sign
    final_records = filtered_df[filtered_df['تقرير نهائي'].str.contains('#', na=False)]

    # Prepare lists to hold the final report texts, values, and additional columns
    reports_list = []
    values_list = []
    car_nums = []
    car_types = []
    entry_dates = []
    company_names = []
    total_value = 0

    # Iterate over the final records and extract the value after '#'
    for _, row in final_records.iterrows():
        final_report = row['تقرير نهائي']
        if '#' in final_report:
            # Split the text at the '#' sign and take the part after it
            parts_after_hash = final_report.split('#')[1]
            # Find the first numeric value after the '#'
            value_str_match = re.search(r'\d+', parts_after_hash)
            if value_str_match:
                value = int(value_str_match.group())
                total_value += value
                reports_list.append(final_report)
                values_list.append(value)
                car_nums.append(row['رقم المركبة'])
                car_types.append(row['نوع المركبه'])
                entry_dates.append(row['تاريخ الدخول'])
                company_names.append(row['اسم الشركه'])

    # Write the results to a text file
    with open("final_report.txt", "w", encoding='utf-8') as f:
        for report, value in zip(reports_list, values_list):
            f.write(f"{report}    -----------    {value} شيكل\n")
        f.write(f"\nTotal Value in شيكل: {total_value} شيكل\n")

    # Create a DataFrame for the Excel file
    excel_df = pd.DataFrame({
        'رقم المركبة': car_nums,
        'نوع المركبه': car_types,
        'تاريخ الدخول': entry_dates,
        'اسم الشركه': company_names,
        'Final Report': reports_list,
        'Value (شيكل)': values_list
    })

    # Append the total value as a new row
    total_row = pd.DataFrame({
        'رقم المركبة': [''],
        'نوع المركبه': [''],
        'تاريخ الدخول': [''],
        'اسم الشركه': [''],
        'Final Report': ['Total Value'],
        'Value (شيكل)': [total_value]
    })
    
    excel_df = pd.concat([excel_df, total_row], ignore_index=True)

    # Write the DataFrame to an Excel file
    excel_file_path = 'final_report.xlsx'
    excel_df.to_excel(excel_file_path, index=False, engine='openpyxl')

    # Load the workbook and set column widths
    workbook = load_workbook(excel_file_path)
    worksheet = workbook.active

    # Set specific column widths
    column_widths = {
        'A': 20,  # رقم المركبة
        'B': 20,  # نوع المركبه
        'C': 25,  # تاريخ الدخول
        'D': 30,  # اسم الشركه
        'E': 50,  # Final Report
        'F': 20   # Value (شيكل)
    }

    for col_letter, width in column_widths.items():
        worksheet.column_dimensions[col_letter].width = width

    workbook.save(excel_file_path)
    messagebox.showinfo("Success", "The final report has been generated in 'final_report.txt' and 'final_report.xlsx'")

def on_start_button_click():
    start_date_str = start_date_var.get()
    end_date_str = end_date_var.get()
    try:
        start_date = pd.to_datetime(start_date_str, format="%d.%m.%Y")
        end_date = pd.to_datetime(end_date_str, format="%d.%m.%Y")
        generate_report(start_date, end_date)
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter dates in the format dd.mm.yyyy")

# Create the GUI
root = Tk()
root.title("Final Report Generator")

# Start Date
Label(root, text="Enter Start Date (dd.mm.yyyy):").grid(row=0, column=0, padx=10, pady=10)
start_date_var = StringVar()
Entry(root, textvariable=start_date_var).grid(row=0, column=1, padx=10, pady=10)

# End Date
Label(root, text="Enter End Date (dd.mm.yyyy):").grid(row=1, column=0, padx=10, pady=10)
end_date_var = StringVar()
Entry(root, textvariable=end_date_var).grid(row=1, column=1, padx=10, pady=10)

# Start Button
Button(root, text="Start", command=on_start_button_click).grid(row=2, column=0, columnspan=2, pady=20)

# Run the GUI event loop
root.mainloop()
