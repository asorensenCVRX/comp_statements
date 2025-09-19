import win32com.client
from VARIABLES import vba_excel_file, atm_comp_file, comp_month
import time
import pythoncom


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

def export_atm_pdf():
    """
        Opens the workbook, writes `b6_value` to ATM!B6, refreshes all connections,
        runs VBA macro SaveAllAsPDF, sleeps, then saves and closes.

        Parameters
        ----------
        atm_comp_file : str
            Full path to the Excel workbook.
        b6_value : any
            Value to write into cell B6 on the ATM worksheet.
        sleep_seconds : int
            Seconds to sleep after running the macro before closing.
        """
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        print("Generating ATM Statements...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(atm_comp_file)

        # 1) Write value to ATM!B6
        ws = workbook.Worksheets("ATM")
        ws.Range("B6").Value = comp_month + " 2025"

        # 2) Refresh all connections and wait for completion
        workbook.RefreshAll()
        # For async queries (Power Query, data connections), wait until done:
        # (CalculateUntilAsyncQueriesDone exists in modern Excel builds)
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            # Fallback polling loop if the above isn't available
            # (short, non-busy wait ~5s max)
            for _ in range(50):
                # If any background query tables still refreshing, keep waiting
                any_refreshing = False
                for ws_i in workbook.Worksheets:
                    for qt in getattr(ws_i, "QueryTables", []):
                        if getattr(qt, "Refreshing", False):
                            any_refreshing = True
                            break
                    if any_refreshing:
                        break
                if not any_refreshing:
                    break
                time.sleep(0.1)

        # Optional: save the value/refresh before macro (if macro reads from file)
        workbook.Save()

        # 3) Run your VBA macro
        excel.Run("SaveAllAsPDF")

        # 4) Sleep
        time.sleep(3)

        # 5) Save & close
        workbook.Close(SaveChanges=True)
        workbook = None

    except Exception as e:
        print(f"An error occurred: {e}")
        # If something goes wrong, try to close without saving to avoid locks
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
