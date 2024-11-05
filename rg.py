import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import xlsxwriter

# Initialize the main application window
root = tk.Tk()
root.title("Service Line Report Generator")
root.geometry("400x200")

# Function to open file dialog to select CSV file
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        file_label.config(text=f"Selected File: {file_path}")
        process_file(file_path)

# Function to process the CSV file and generate the report
import pandas as pd
import xlsxwriter

import pandas as pd
import xlsxwriter

def process_file(file_path):
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
        def map_service_line(cpt):
            for service, types in service_lines.items():
                for visit_type, codes in types.items():
                    if cpt in codes:
                        return service, visit_type
            return 'Other Services', 'Uncategorized'

        # Map CPT codes and clinician names to service lines and visit types
        df['Clinician'] = df['Clinician'].map(provider_mappings).fillna('Unknown')
        df['Service Line'], df['Visit Type'] = zip(*df['CPT'].apply(map_service_line))

        # Extract Year-Month for grouping
        df['Year-Month'] = pd.to_datetime(df['DOS'], errors='coerce').dt.to_period('M')

        # Pivot data to show month-by-month breakdown with totals
        pivot_df = df.pivot_table(
            index=['Service Line', 'Clinician', 'Visit Type'],
            columns='Year-Month',
            values='CPT',
            aggfunc='count',
            fill_value=0
        )
        pivot_df['Total'] = pivot_df.sum(axis=1)
        pivot_df.reset_index(inplace=True)

        # Define output file path and initialize Excel writer
        output_path = "Service_Line_Report.xlsx"
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Service Line Report")

            # Write headers
            months = sorted(pivot_df.columns[3:-1])  # Exclude 'Service Line', 'Clinician', 'Visit Type', 'Total'
            header = ["Clinician", "Service Line", "Visit Type"] + [str(m) for m in months] + ["Total"]
            worksheet.write_row(0, 0, header)

            # Start writing data from row 1
            row = 1
            current_service_line = ""
            
            # Iterate over rows in pivot_df to structure the report
            for _, record in pivot_df.iterrows():
                service_line, clinician, visit_type = record['Service Line'], record['Clinician'], record['Visit Type']
                
                # Add a new section header for each service line
                if service_line != current_service_line:
                    worksheet.write(row, 0, service_line)
                    worksheet.write_row(row, 2, [""] * len(months) + [""])  # Empty row for spacing
                    current_service_line = service_line
                    row += 1

                # Write clinician, visit type, monthly data, and total
                row_data = [clinician, "", visit_type] + list(record[3:-1]) + [record['Total']]
                worksheet.write_row(row, 0, row_data)
                row += 1

        messagebox.showinfo("Success", f"Report generated successfully at {output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI elements
file_label = tk.Label(root, text="No file selected", wraplength=300)
file_label.pack(pady=10)

select_button = tk.Button(root, text="Select CSV File", command=select_file)
select_button.pack(pady=10)

generate_button = tk.Button(root, text="Generate Report", command=lambda: process_file(file_label.cget("text").split(": ")[1]))
generate_button.pack(pady=10)

# Run the application
root.mainloop()
