#Import Module
from win32com import client
import sys
import os
import comtypes.client

#  python excel_to_pdf_convert_all.py . excel-output/

def main():
#%% Get console arguments
    input_folder_path = "."
    output_folder_path = "excel-output/"
    # input_folder_path = sys.argv[1]
    # output_folder_path = sys.argv[2]

    #%% Convert folder paths to Windows format
    input_folder_path = os.path.abspath(input_folder_path)
    output_folder_path = os.path.abspath(output_folder_path)

    #%% Get files in input folder
    input_file_paths = os.listdir(input_folder_path)

    #%% Convert each file
    for input_file_name in input_file_paths:

        # Skip if file does not contain a power point extension
        if not input_file_name.lower().endswith((".xls", ".xlsx")):
            continue
        
        # Open Microsoft Excel
        excel = client.Dispatch("Excel.Application")

        # Get current working directory
        cwd = os.getcwd()
        filename = input_file_name.split(".")
        
        # Read Excel File
        sheets = excel.Workbooks.Open(cwd + '\\' + filename[0])
        # sheets = excel.Workbooks.Open('x:\\any_to_pdf_python\\test.xlsx')
        # sheets = excel.Workbooks.Open('Excel File Path')
        work_sheets = sheets.Worksheets[0]
        
        # Convert into PDF File
        work_sheets.ExportAsFixedFormat(0, cwd + '\\' + 'excel-output' + '\\' + filename[0])
        # work_sheets.ExportAsFixedFormat(0, 'x:\\any_to_pdf_python\\excel-output\\test.pdf')
        # work_sheets.ExportAsFixedFormat(0, 'PDF File Path')

        sheets.Close()