import win32com.client


# must pip install pywin32

def export_to_pdf():
    # Path to Excel file
    excel_file_path = r"C:\Users\asorensen\PycharmProjects\generate_comp_statements\export_statements.xlsm"

    # Create an instance of Excel application
    excel = win32com.client.Dispatch("Excel.Application")

    # Open the Excel file
    workbook = excel.Workbooks.Open(excel_file_path)

    # Run the VBA script
    excel.Application.Run("ExportPDF")

    # Close the workbook and Excel application
    workbook.Close(SaveChanges=False)
    excel.Quit()
