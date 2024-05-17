import win32com.client


# must pip install pywin32

def export_to_pdf(macro_name):
    """Available macros are "AMExportPDF", "RMExportPDF", and "CSRExportPDF"."""
    # Path to Excel file
    excel_file_path = (r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\PyWorkbooks & "
                       r"Queries\export_statements.xlsm")

    # Create an instance of Excel application
    excel = win32com.client.Dispatch("Excel.Application")

    # Open the Excel file
    workbook = excel.Workbooks.Open(excel_file_path)

    # Run the VBA script
    excel.Application.Run(macro_name)

    # Close the workbook and Excel application
    workbook.Close(SaveChanges=False)
    excel.Quit()
