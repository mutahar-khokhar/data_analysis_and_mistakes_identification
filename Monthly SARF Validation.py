import tkinter as tk
from tkinter import filedialog
import pandas as pd
from openpyxl import Workbook
from datetime import datetime

def check_conditions(row):
    details = []

    # PPW Conditions
    if row['PPW_12_TWC'] > row['PPW_11_TWN']:
        details.append('PPW_12_TWC')
    if row['PPW_13_TWR'] > row['PPW_11_TWN']:
        details.append('PPW_13_TWR')
    if row['PPW_14_STM'] > row['PPW_13_TWR']:
        details.append('PPW_14_STM')
    if row['PPW_14_LTM'] > row['PPW_13_TWR']:
        details.append('PPW_14_LTM')
    if row['PPW_14_PM'] > row['PPW_13_TWR']:
        details.append('PPW_14_PM')
    if row['PPW_14_TM'] > row['PPW_13_TWR']:
        details.append('PPW_14_TM')
    if row['PPW_11_TWN'] - row['PPW_13_TWR'] != row['PPW_15_TW_B']:
        details.append('PPW_15_TW_B')

    # NU Conditions
    if row['NU_22_TWC'] > row['NU_21_TWN']:
        details.append('NU_22_TWC')
    if row['NU_23_TWR'] > row['NU_21_TWN']:
        details.append('NU_23_TWR')
    if row['NU_24_STM'] > row['NU_23_TWR']:
        details.append('NU_24_STM')
    if row['NU_24_LTM'] > row['NU_23_TWR']:
        details.append('NU_24_LTM')
    if row['NU_24_PM'] > row['NU_23_TWR']:
        details.append('NU_24_PM')
    if row['NU_24_TMP'] > row['NU_23_TWR']:
        details.append('NU_24_TMP')
    if row['NU_21_TWN'] - row['NU_23_TWR'] != row['NU_25_TW_B']:
        details.append('NU_25_TW_B')

    # PGW Conditions
    if row['PGW_32_TWC'] > row['PGW_31_TWN']:
        details.append('PGW_32_TWC')
    if row['PGW_33_TWR'] > row['PGW_31_TWN']:
        details.append('PGW_33_TWR')
    if row['PGW_34_ANC'] > row['PGW_33_TWR']:
        details.append('PGW_34_ANC')
    if row['PGW_31_TWN'] - row['PGW_33_TWR'] != row['PGW_35_TW_B']:
        details.append('PGW_35_TW_B')

    # PGU Conditions
    if row['PGU_42_TWC'] > row['PGU_41_TWN']:
        details.append('PGU_42_TWC')
    if row['PGU_43_TWR'] > row['PGU_41_TWN']:
        details.append('PGU_43_TWR')
    if row['PGU_44_CAC'] > row['PGU_43_TWR']:
        details.append('PGU_44_CAC')
    if row['PGU_44_PAC'] > row['PGU_43_TWR']:
        details.append('PGU_44_PAC')
    if row['PGU_44_TWS'] > row['PGU_43_TWR']:
        details.append('PGU_44_TWS')

    # PG3T Conditions
    if row['PG3T_52_TWC'] > row['PG3T_51_TWN']:
        details.append('PG3T_52_TWC')
    if row['PG3T_53_TWR'] > row['PG3T_51_TWN']:
        details.append('PG3T_53_TWR')
    if row['PG3T_51_TWN'] - row['PG3T_52_TWC'] != row['PG3T_55_TW_B']:
        details.append('PG3T_55_TW_B')

    # W24 Conditions
    if row['W24_62_TWC'] > row['W24_61_TWN']:
        details.append('W24_62_TWC')
    if row['W24_63_TWR'] > row['W24_61_TWN']:
        details.append('W24_63_TWR')
    if row['W24_64_STM'] > row['W24_63_TWR']:
        details.append('W24_64_STM')
    if row['W24_64_LTM'] > row['W24_63_TWR']:
        details.append('W24_64_LTM')
    if row['W24_64_PM'] > row['W24_63_TWR']:
        details.append('W24_64_PM')
    if row['W24_64_TMP'] > row['W24_63_TWR']:
        details.append('W24_64_TMP')
    if row['W24_61_TWN'] - row['W24_63_TWR'] != row['W24_65_TW_B']:
        details.append('W24_65_TW_B')

    # TMU Conditions
    if row['TMU_72_TWC'] > row['TMU_71_TWN']:
        details.append('TMU_72_TWC')
    if row['TMU_73_TWR'] > row['TMU_71_TWN']:
        details.append('TMU_73_TWR')
    if row['TMU_74_STM'] > row['TMU_73_TWR']:
        details.append('TMU_74_STM')
    if row['TMU_74_LTM'] > row['TMU_73_TWR']:
        details.append('TMU_74_LTM')
    if row['TMU_74_PM'] > row['TMU_73_TWR']:
        details.append('TMU_74_PM')
    if row['TMU_74_TMP'] > row['TMU_73_TWR']:
        details.append('TMU_74_TMP')
    if row['TMU_71_TWN'] - row['TMU_73_TWR'] != row['TMU_75_TW_B']:
        details.append('TMU_75_TW_B')

    return ', '.join(details)

# Function to perform data analysis and save results to Excel
def analyze_data():
    # Get the selected Excel file path
    file_path = filedialog.askopenfilename()

    try:
        # Read the Excel file into a Pandas DataFrame
        df = pd.read_excel(file_path)

        # Parse the 'from' and 'to' date inputs
        from_date = datetime.strptime(from_date_entry.get(), '%d/%m/%Y')
        to_date = datetime.strptime(to_date_entry.get(), '%d/%m/%Y')

        # Convert the 'Reporting_Month' column to datetime objects
        df['Reporting_Month'] = pd.to_datetime(df['Reporting_Month'], format='%d/%m/%Y')

        # Filter the DataFrame based on the date range
        filtered_df = df[(df['Reporting_Month'] >= from_date) & (df['Reporting_Month'] <= to_date)]

        # Initialize an Excel writer
        output_file = 'output.xlsx'
        writer = pd.ExcelWriter(output_file, engine='openpyxl')

        # Write the filtered data to a new Excel sheet
        filtered_df.to_excel(writer, sheet_name='Filtered Data', index=False)

        # Create a list to hold rows that meet the condition
        results = []

        # Loop through the filtered DataFrame and check conditions
        for index, row in filtered_df.iterrows():
            condition_details = check_conditions(row)

            if condition_details:
                results.append([row['LHW_Code'], row['Reporting_Month'].strftime('%d/%m/%Y'),
                                row['Dist_ID'], row['User_entry'], row['Doc'], condition_details])

        # Create a DataFrame from the list of results
        results_df = pd.DataFrame(results, columns=['LHW Code', 'Reporting Month', 'District ID', 'Entered by', 'Date of Entry', 'Conditions Not Met'])

        # Write the results to a new sheet in the Excel file
        results_df.to_excel(writer, sheet_name='Results', index=False)

        # Save and close the Excel file
        writer.book.save(output_file)
        writer.close()

        # Notify the user that the analysis is complete
        status_label.config(text='Analysis and save complete. Output file: ' + output_file)

    except Exception as e:
        status_label.config(text='Error: ' + str(e))

# Create the GUI window
root = tk.Tk()
root.title("Data Analysis App")

# Create and configure GUI elements
file_label = tk.Label(root, text="Select an Excel file:")
file_label.pack()

browse_button = tk.Button(root, text="Browse", command=analyze_data)
browse_button.pack()

from_date_label = tk.Label(root, text="From Date (dd/mm/yyyy):")
from_date_label.pack()

from_date_entry = tk.Entry(root)
from_date_entry.pack()

to_date_label = tk.Label(root, text="To Date (dd/mm/yyyy):")
to_date_label.pack()

to_date_entry = tk.Entry(root)
to_date_entry.pack()

analyze_button = tk.Button(root, text="Analyze Data", command=analyze_data)
analyze_button.pack()

status_label = tk.Label(root, text="")
status_label.pack()

root.mainloop()
