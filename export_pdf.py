import win32com.client
from VARIABLES import vba_excel_file, atm_comp_file, comp_month
import time
import pythoncom

# must pip install pywin32

def export_to_pdf(macro_name):
    """Run a statement-export macro in a dedicated Excel instance and always close it."""
    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        # DispatchEx starts a separate Excel process so we don't attach to a user-open instance.
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # Open the Excel file
        workbook = excel.Workbooks.Open(vba_excel_file)

        # Run the VBA script
        excel.Application.Run(macro_name)

        time.sleep(2)

        # Close the workbook and Excel application
        workbook.Close(SaveChanges=True)
        workbook = None
    except Exception as e:
        print(f"An error occurred: {e}")
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        raise
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    # quitting excel is not necessary because the VBA script does it
    # finally:
    #     excel.Quit()
