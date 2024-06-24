import win32com.client
from VARIABLES import vba_excel_file
import time


# must pip install pywin32

def export_to_pdf(macro_name):
    """Available macros are "AMExportPDF", "RMExportPDF", and "CSRExportPDF"."""

    try:
        # Create an instance of Excel application
        excel = win32com.client.Dispatch("Excel.Application")

        # Open the Excel file
        workbook = excel.Workbooks.Open(vba_excel_file)

        # Run the VBA script
        excel.Application.Run(macro_name)

        time.sleep(2)

        # Close the workbook and Excel application
        workbook.Close(SaveChanges=True)
    except Exception as e:
        print(f"An error occurred: {e}")

    # quitting excel is not necessary because the VBA script does it
    # finally:
    #     excel.Quit()
