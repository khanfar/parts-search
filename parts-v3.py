import pandas as pd
import requests
import re
from tkinter import Tk, Label, Entry, Button, StringVar, messagebox, filedialog
from openpyxl import load_workbook

df = None

def load_csv_from_url(url):
    response = requests.get(url)
    with open("your_data.csv", "wb") as f:
        f.write(response.content)
    return pd.read_csv("your_data.csv")

def load_csv_from_file(file_path):
    return pd.read_csv(file_path)

def generate_main_report(start_date, end_date, df):
    filtered_df, reports_list, values_list, car_nums, car_types, entry_dates, company_names, total_value = filter_and_extract_data(start_date, end_date, df)
    
    # Write the results to a text file
    with open("final_report.txt", "w", encoding='utf-8') as f:
        for report, value in zip(reports_list, values_list):
            f.write(f"{report}    -----------    {value} شيكل\n")
        f.write(f"\nTotal Value in شيكل: {total_value} شيكل\n")

    # Create a DataFrame for the Excel file
    excel_df = create_report_dataframe(car_nums, car_types, entry_dates, company_names, reports_list, values_list, total_value)

    # Write the DataFrame to an Excel file
    write_to_excel(excel_df, 'final_report.xlsx')

def generate_mini_report(start_date, end_date, df):
    filtered_df, reports_list, values_list, car_nums, car_types, entry_dates, company_names, total_value = filter_and_extract_data(start_date, end_date, df)

    # Prepare the mini report data
    mini_reports_list = [extract_all_after_hashes(report) for report in reports_list]

    # Create a DataFrame for the mini report
    mini_excel_df = create_report_dataframe(car_nums, car_types, entry_dates, company_names, mini_reports_list, values_list, total_value)

    # Write the mini DataFrame to an Excel file
    write_to_excel(mini_excel_df, 'mini_report.xlsx')

def filter_and_extract_data(start_date, end_date, df):
    df['تاريخ الدخول'] = pd.to_datetime(df['تاريخ الدخول'], format='%d.%m.%Y', errors='coerce')
    mask = (df['تاريخ الدخول'] >= start_date) & (df['تاريخ الدخول'] <= end_date)
    filtered_df = df.loc[mask]
    final_records = filtered_df[filtered_df['تقرير نهائي'].str.contains('#', na=False)]

    reports_list = []
    values_list = []
    car_nums = []
    car_types = []
    entry_dates = []
    company_names = []
    total_value = 0

    for _, row in final_records.iterrows():
        final_report = row['تقرير نهائي']
        if '#' in final_report:
            parts_after_hash = final_report.split('#')
            values_in_report = []
            for part in parts_after_hash[1:]:
                value_str_match = re.search(r'\d+', part)
                if value_str_match:
                    value = int(value_str_match.group())
                    values_in_report.append(value)
            
            report_total_value = sum(values_in_report)
            total_value += report_total_value
            reports_list.append(final_report)
            values_list.append(report_total_value)
            car_nums.append(row['رقم المركبة'])
            car_types.append(row['نوع المركبه'])
            entry_dates.append(row['تاريخ الدخول'].date())  # Store only the date part
            company_names.append(row['اسم الشركه'])

    return filtered_df, reports_list, values_list, car_nums, car_types, entry_dates, company_names, total_value

def extract_all_after_hashes(final_report):
    parts_after_hash = final_report.split('#')
    text_after_hashes = [part.strip() for part in parts_after_hash[1:]]
    return ' # '.join(text_after_hashes)

def create_report_dataframe(car_nums, car_types, entry_dates, company_names, reports_list, values_list, total_value):
    excel_df = pd.DataFrame({
        'رقم المركبة': car_nums,
        'نوع المركبه': car_types,
        'تاريخ الدخول': entry_dates,
        'اسم الشركه': company_names,
        'Final Report': reports_list,
        'Value (شيكل)': values_list
    })

    total_row = pd.DataFrame({
        'رقم المركبة': [''],
        'نوع المركبه': [''],
        'تاريخ الدخول': [''],
        'اسم الشركه': [''],
        'Final Report': ['Total Value'],
        'Value (شيكل)': [total_value]
    })
    
    return pd.concat([excel_df, total_row], ignore_index=True)

def write_to_excel(df, filename):
    df.to_excel(filename, index=False, engine='openpyxl')
    workbook = load_workbook(filename)
    worksheet = workbook.active

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

    # Set the date format for the entry dates column
    for cell in worksheet['C']:
        cell.number_format = 'd/m/yyyy'

    workbook.save(filename)
    messagebox.showinfo("Success", f"The report has been generated in '{filename}'")

def on_main_report_button_click():
    start_date_str = start_date_var.get()
    end_date_str = end_date_var.get()
    try:
        start_date = pd.to_datetime(start_date_str, format="%d.%m.%Y")
        end_date = pd.to_datetime(end_date_str, format="%d.%m.%Y")
        generate_main_report(start_date, end_date, df)
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter dates in the format dd.mm.yyyy")

def on_mini_report_button_click():
    start_date_str = start_date_var.get()
    end_date_str = end_date_var.get()
    try:
        start_date = pd.to_datetime(start_date_str, format="%d.%m.%Y")
        end_date = pd.to_datetime(end_date_str, format="%d.%m.%Y")
        generate_mini_report(start_date, end_date, df)
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter dates in the format dd.mm.yyyy")

def on_attach_file_click():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        df = load_csv_from_file(file_path)
        messagebox.showinfo("File Loaded", f"CSV file loaded successfully from {file_path}")

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
Button(root, text="Generate Main Report", command=on_main_report_button_click).grid(row=2, column=0, columnspan=2, pady=10)

# Mini Report Button
Button(root, text="Generate Mini Report", command=on_mini_report_button_click).grid(row=3, column=0, columnspan=2, pady=10)

# Attach File Button
Button(root, text="Attach CSV File", command=on_attach_file_click).grid(row=4, column=0, columnspan=2, pady=10)

# Try to Load default CSV file when GUI starts, handle cases with no internet connectivity
try:
    df = load_csv_from_url("https://huggingface.co/spaces/mhammad/Khanfar/raw/main/docs/your_data.csv")
except requests.ConnectionError:
    messagebox.showwarning("Connection Error", "Failed to fetch default CSV file. Please attach a CSV file manually.")

# Run the GUI event loop
root.mainloop()
