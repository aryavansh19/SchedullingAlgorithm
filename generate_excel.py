import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import os

def create_solution_excel():
    input_file = '/Users/aryavansh/Documents/Projects/Scheduling/Scheduling ToDos.xlsx'
    output_file = '/Users/aryavansh/Documents/Projects/Scheduling/Assignment_Solution.xlsx'
    
    # Read original input if exists, else create dummy
    if os.path.exists(input_file):
        xls = pd.ExcelFile(input_file)
        if 'INPUT Interview Requests' in xls.sheet_names:
            df_input = pd.read_excel(xls, 'INPUT Interview Requests')
        else:
            df_input = pd.DataFrame(columns=['Company', 'Interviewer', 'Interviewer Email', 'Candidate', 'Candidate Email ', 'Candidate Phone', 'Scheduling method', 'Round', 'Added On'])
    else:
        df_input = pd.DataFrame(columns=['Company', 'Interviewer', 'Interviewer Email', 'Candidate', 'Candidate Email ', 'Candidate Phone', 'Scheduling method', 'Round', 'Added On'])

    # Add a few dummy rows for demonstration if it's empty
    if df_input.empty:
        df_input.loc[0] = ['Google', 'Sundar P.', 'sundar@google.com', 'John Doe', 'john@gmail.com', '9999999999', 'calendly.com/sundar', 'R1', '2023-10-01']
        df_input.loc[1] = ['Meta', 'Mark Z.', 'mark@meta.com', 'Jane Smith', 'jane@gmail.com', '8888888888', 'calendly.com/mark', 'R2', '2023-10-01']

    wb = Workbook()
    ws_input = wb.active
    ws_input.title = "INPUT Interview Requests"
    
    # Write INPUT sheet
    for r in dataframe_to_rows(df_input, index=False, header=True):
        ws_input.append(r)
    
    # Format Headers
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    for cell in ws_input[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # ---------------------------------------------------------
    # RAM'S SHEET
    # ---------------------------------------------------------
    ws_ram = wb.create_sheet(title="OUTPUT ToDos for Ram")
    ram_headers = [
        "Candidate Name", "Company", "Phone", "Email", "Calendly Link",
        "10:00 AM Initial Blast (WA+Email)",
        "01:00 PM Calendly Check 1",
        "01:30 PM Follow-up Nudge (WA)",
        "03:30 PM Calendly Check 2"
    ]
    ws_ram.append(ram_headers)
    for cell in ws_ram[1]:
        cell.font = header_font
        cell.fill = PatternFill(start_color="9BBB59", end_color="9BBB59", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    # Populate Ram's sheet with basic info
    for _, row in df_input.iterrows():
        ws_ram.append([
            row.get('Candidate', ''),
            row.get('Company', ''),
            row.get('Candidate Phone', ''),
            row.get('Candidate Email ', ''),
            row.get('Scheduling method', ''),
            "", "", "", ""
        ])

    # Data Validations for Ram
    dv_blast = DataValidation(type="list", formula1='"Done,Failed"', allow_blank=True)
    dv_check1 = DataValidation(type="list", formula1='"Scheduled,Pending"', allow_blank=True)
    dv_nudge = DataValidation(type="list", formula1='"Done,N/A"', allow_blank=True)
    dv_check2 = DataValidation(type="list", formula1='"Scheduled,Escalate to Call"', allow_blank=True)

    ws_ram.add_data_validation(dv_blast)
    ws_ram.add_data_validation(dv_check1)
    ws_ram.add_data_validation(dv_nudge)
    ws_ram.add_data_validation(dv_check2)

    max_row = len(df_input) + 1
    dv_blast.add(f"F2:F{max_row}")
    dv_check1.add(f"G2:G{max_row}")
    dv_nudge.add(f"H2:H{max_row}")
    dv_check2.add(f"I2:I{max_row}")

    # ---------------------------------------------------------
    # SHYAM'S SHEET
    # ---------------------------------------------------------
    ws_shyam = wb.create_sheet(title="OUTPUT ToDos for Shyam")
    shyam_headers = [
        "Candidate Name", "Company", "Phone", "Calendly Link",
        "04:00 PM Call Disposition",
        "Next Day 09:30 AM Call 2"
    ]
    ws_shyam.append(shyam_headers)
    for cell in ws_shyam[1]:
        cell.font = header_font
        cell.fill = PatternFill(start_color="F08080", end_color="F08080", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    # In a real dynamic sheet, we'd use FILTER or IF formulas. Here we just add the rows
    # and instruct Shyam to only look at ones escalated. We can add a column "Escalated?"
    ws_shyam.cell(row=1, column=1, value="Candidate Name")
    # Actually, we can insert an Excel formula in Shyam's sheet linking to Ram's sheet 
    # But for simplicity, we just dump the candidates and add the Dropdowns.
    for i, row in df_input.iterrows():
        ws_shyam.append([
            row.get('Candidate', ''),
            row.get('Company', ''),
            row.get('Candidate Phone', ''),
            row.get('Scheduling method', ''),
            "", ""
        ])

    # Data Validations for Shyam
    dv_call1 = DataValidation(type="list", formula1='"Scheduled on Call,Unanswered,Invalid Number,Will do later"', allow_blank=True)
    dv_call2 = DataValidation(type="list", formula1='"Scheduled on Call,Still Unanswered,Escalate to AM"', allow_blank=True)
    
    ws_shyam.add_data_validation(dv_call1)
    ws_shyam.add_data_validation(dv_call2)
    
    dv_call1.add(f"E2:E{max_row}")
    dv_call2.add(f"F2:F{max_row}")

    # Adjust column widths
    for ws in [ws_input, ws_ram, ws_shyam]:
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(output_file)
    print(f"Successfully created {output_file}")

if __name__ == '__main__':
    create_solution_excel()
