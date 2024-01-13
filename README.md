# SalesReportExcel

import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference
import string

def automate_excel(file_name: str) -> None:
    """
    Automates the creation of Excel reports from a monthly sales data file.

    Parameters:
    - file_name (str): Name of the Excel file containing monthly sales data.
    """
    # Read Excel file
    excel_file = pd.read_excel(file_name)
    
    # Create pivot table
    report_table = excel_file.pivot_table(index='Gender', columns='Product line', values='Total', aggfunc='sum').round(0)
    
    # Separate the month and extension from the file name
    month_and_extension = file_name.split('_')[1]
    
    # Save the pivot table to the Excel file
    report_table.to_excel(f'report_{month_and_extension}', sheet_name='Report', startrow=4)
    
    # Load workbook and select sheet
    wb = load_workbook(f'report_{month_and_extension}')
    sheet = wb['Report']
    
    # Cell references (original spreadsheet)
    min_column = wb.active.min_column
    max_column = wb.active.max_column
    min_row = wb.active.min_row
    max_row = wb.active.max_row
    
    # Add a bar chart
    barchart = BarChart()
    data = Reference(sheet, min_col=min_column+1, max_col=max_column, min_row=min_row, max_row=max_row) # including headers
    categories = Reference(sheet, min_col=min_column, max_col=min_column, min_row=min_row+1, max_row=max_row) # not including headers
    barchart.add_data(data, titles_from_data=True)
    barchart.set_categories(categories)
    sheet.add_chart(barchart, "B12")  # chart location
    barchart.title = 'Sales by Product Line'
    barchart.style = 2  # choose the chart style
    
    # Apply formulas
    alphabet = list(string.ascii_uppercase)
    excel_alphabet = alphabet[0:max_column]
    
    for i in excel_alphabet:
        if i != 'A':
            sheet[f'{i}{max_row+1}'] = f'=SUM({i}{min_row+1}:{i}{max_row})'
            sheet[f'{i}{max_row+1}'].style = 'Currency'
    
    sheet[f'{excel_alphabet[0]}{max_row+1}'] = 'Total'
    
    # Get the month name
    month_name = month_and_extension.split('.')[0]
    
    # Format the report
    sheet['A1'] = 'Sales Report'
    sheet['A2'] = month_name.title()
    sheet['A1'].font = Font('Arial', bold=True, size=20)
    sheet['A2'].font = Font('Arial', bold=True, size=10)
    
    # Save the report
    wb.save(f'report_{month_and_extension}')
    return

# Example of use for an entire year
automate_excel('/content/sales_2021.xlsx')

# Example of use for individual monthly reports
automate_excel('/content/sales_january.xlsx')
automate_excel('/content/sales_february.xlsx')
automate_excel('/content/sales_march.xlsx')

# Option: Concatenating monthly reports and creating a report for the year
excel_file_1 = pd.read_excel('sales_january.xlsx')
excel_file_2 = pd.read_excel('sales_february.xlsx')
excel_file_3 = pd.read_excel('sales_march.xlsx')

new_file = pd.concat([excel_file_1, excel_file_2, excel_file_3], ignore_index=True)
new_file.to_excel('sales_2021.xlsx')
automate_excel('sales_2021.xlsx')
