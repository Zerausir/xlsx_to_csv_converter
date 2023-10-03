import os
import pandas as pd
import csv
import warnings
from dotenv import load_dotenv

load_dotenv()

# Define the paths for the files
excel_dir_path = os.getenv('excel_dir_path')
csv_dir_path = f'{excel_dir_path}/csv'

# Make sure the CSV directory exists
if not os.path.exists(csv_dir_path):
    os.makedirs(csv_dir_path)

# List all the .xlsx files in the Excel directory
xlsx_files = [f for f in os.listdir(excel_dir_path) if f.endswith('.xlsx')]

# Suppress the specific warning
warnings.filterwarnings("ignore", category=UserWarning, message="Cannot parse header or footer so it will be ignored")

# Iterate through each .xlsx file and convert it to .csv
for xlsx_file in xlsx_files:
    # Build the full paths for the input and output files
    excel_file_path = os.path.join(excel_dir_path, xlsx_file)
    csv_file_name = os.path.splitext(xlsx_file)[0] + '.csv'
    csv_file_path = os.path.join(csv_dir_path, csv_file_name)

    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(excel_file_path)


    # Define a function to change the date format
    def change_date_format(datetime_obj):
        return datetime_obj.strftime('%d/%m/%Y %H:%M:%S.%f')


    # Apply the function to the 'Tiempo' column
    df['Tiempo'] = df['Tiempo'].apply(change_date_format)

    # Save the DataFrame as a CSV file with all fields enclosed in double quotes
    df.to_csv(csv_file_path, index=False, sep=',', decimal='.', encoding='ansi', quoting=csv.QUOTE_ALL)

    print(f'Excel file "{excel_file_path}" has been converted to CSV and saved as "{csv_file_path}"')
