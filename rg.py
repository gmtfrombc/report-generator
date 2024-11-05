import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import xlsxwriter
from datetime import datetime

# Initialize the main application window
root = tk.Tk()
root.title("Service Line Report Generator")
root.geometry("400x300")

# Function to open file dialog to select CSV file
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        file_label.config(text=f"Selected File: {file_path}")
        selected_file.set(file_path)

# Function to process the CSV file and generate the report
def process_file(file_path, num_months, save_path):
    try:
        # Load the CSV data
        df = pd.read_csv(file_path)

        # Ensure CPT codes are strings for consistent comparison
        df['CPT'] = df['CPT'].astype(str)

        # Define mappings for service lines and CPT codes as strings
        service_lines = {
            'Medical Services': {
                'E/M Visits': ['99202', '99203', '99204', '99205', '99212', '99213', '99214', '99215'],
                'Prevention Visits': ['99381', '99382', '99383', '99384', '99385', '99386', '99387', '99392', '99393', '99394', '99395', '99396', '99397'],
                'Prevention Counseling': ['99401', '99402', '99403', '99404'],
                'Procedures': ['11200', '17000', '17003', '17004', '20605', '20610', '69210'],
                'EKG': ['93000'],
                'Portal Visits': ['99421', '99422', '99442', '99443'],
                'GCode Screenings': ['96127', 'G0442', 'G0444']
            },
            'Lab Services': {
                'Lab Draw': ['36415'],
                'Lab Tests': ['81003', '81025', '86308', '87426', '87804', '87880', '88175', '87624']
            },
            'MA Services': {
                'Injections': ['96372', 'J3301']
            },
            'Physical Therapy': {
                'PT Treatments': ['97012', '97110', '97112', '97116', '97140', '97162', '97164', '97530']
            },
            'Chiropractic': {
                'Chiropractic Treatments': ['98940', '98941', '98943']
            },
            'Health Coaching': {
                'Health Coach Sessions': ['G0447']
            }
        }

        provider_mappings = {
            'TOLSON, GRAEME MEREDITH': 'TOLSON',
            'TOPHAM, MAREN': 'TOPHAM',
            'ALLEN, JEFFERSON M': 'ALLEN',
            'SLOAN, JACOY': 'SLOAN',
            'RIGGS, BRANNICK BARTON': 'RIGGS',
            'HAWKINS, DAVID ELI': 'HAWKINS',
            'ANDERSON, ADAM EARL': 'ANDERSON',
            'BAKER, STEVE A': 'BAKER',
        }

        # Function to map CPT codes to service lines and visit types
        def map_service_line(cpt, clinician):
            clinician_service_line = None
            if clinician in ["HAWKINS", "ANDERSON", "BAKER"]:
                clinician_service_line = "Chiropractic" if clinician in ["ANDERSON", "BAKER"] else "Physical Therapy"
            
            for service, types in service_lines.items():
                for visit_type, codes in types.items():
                    if cpt in codes:
                        if clinician_service_line:
                            return clinician_service_line, "Chiropractic Treatments" if clinician_service_line == "Chiropractic" else visit_type
                        return service, visit_type
            return 'Other Services', 'Uncategorized'

        # Map CPT codes and clinician names to service lines and visit types
        df['Clinician'] = df['Clinician'].map(provider_mappings).fillna('Unknown')
        df['Service Line'], df['Visit Type'] = zip(*df.apply(lambda x: map_service_line(x['CPT'], x['Clinician']), axis=1))

        # Aggregate all "Lab Services" into a single row for each type
        df.loc[df['Service Line'] == 'Lab Services', 'Clinician'] = 'Lab Services'
        df.loc[df['Service Line'] == 'Health Coaching', 'Clinician'] = 'Health Coaching'

        # Extract Year-Month for grouping and filter by the last 'num_months'
        df['Year-Month'] = pd.to_datetime(df['DOS'], errors='coerce').dt.to_period('M')
        recent_periods = pd.date_range(end=datetime.now(), periods=num_months, freq='M').to_period('M')
        df = df[df['Year-Month'].isin(recent_periods)]

        # Pivot data
        pivot_df = df.pivot_table(
            index=['Service Line', 'Clinician', 'Visit Type'],
            columns='Year-Month',
            values='CPT',
            aggfunc='count',
            fill_value=0
        )
        pivot_df['Total'] = pivot_df.sum(axis=1)
        pivot_df.reset_index(inplace=True)

        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Service Line Report")
            bold_format = workbook.add_format({'bold': True})
            month_header_format = workbook.add_format({'bold': True, 'align': 'center'})
            service_line_format = workbook.add_format({'bold': True, 'align': 'left', 'bottom': 1})
            header = ["Visit Type"] + sorted([str(m) for m in pivot_df.columns[3:-1]]) + ["Total"]
            worksheet.write_row(0, 1, header, month_header_format)

            ordered_service_lines = ['Medical Services', 'Physical Therapy', 'Chiropractic', 'Health Coaching', 'MA Services', 'Lab Services']
            row = 1
            for service_line in ordered_service_lines:
                worksheet.write(row, 0, "")  # Blank row before service line title
                row += 1
                worksheet.write(row, 0, service_line, service_line_format)
                row += 1

                service_data = pivot_df[pivot_df['Service Line'] == service_line]
                current_clinician = ""
                for i, record in service_data.iterrows():
                    clinician, visit_type = record['Clinician'], record['Visit Type']
                    
                    if clinician != current_clinician:
                        current_clinician = clinician
                        worksheet.write(row, 0, clinician)
                    else:
                        worksheet.write(row, 0, "")
                    
                    row_data = [visit_type] + list(record[3:-1]) + [record['Total']]
                    worksheet.write_row(row, 1, row_data)
                    row += 1

                if i < len(service_data) - 1 and service_data.iloc[i + 1]['Clinician'] != clinician:
                    row += 1

            row += 1

        messagebox.showinfo("Success", f"Report generated successfully at {save_path}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Validate and run the report generation
def run_report():
    try:
        num_months = int(month_entry.get())
        if num_months <= 0:
            raise ValueError("Number of months must be positive.")
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="PMH Production Report.xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            process_file(selected_file.get(), num_months, save_path)
    except ValueError:
        messagebox.showerror("Input Error", "Please enter a valid number of months.")

# GUI elements
selected_file = tk.StringVar()
file_label = tk.Label(root, text="No file selected", wraplength=300)
file_label.pack(pady=10)

select_button = tk.Button(root, text="Select CSV File", command=select_file)
select_button.pack(pady=10)

month_label = tk.Label(root, text="Enter number of months:")
month_label.pack(pady=5)
month_entry = tk.Entry(root)
month_entry.pack(pady=5)

generate_button = tk.Button(root, text="Generate Report", command=run_report)
generate_button.pack(pady=10)

# Run the application
root.mainloop()
