import win32com.client
from VARIABLES import vba_excel_file


# must pip install pywin32

def export_to_pdf(macro_name):
    """Available macros are "AMExportPDF", "RMExportPDF", and "CSRExportPDF"."""

    # Create an instance of Excel application
    excel = win32com.client.Dispatch("Excel.Application")

    # Open the Excel file
    workbook = excel.Workbooks.Open(vba_excel_file)

    # Run the VBA script
    excel.Application.Run(macro_name)

    # Close the workbook and Excel application
    workbook.Close(SaveChanges=False)
    excel.Quit()
