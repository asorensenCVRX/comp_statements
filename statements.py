from export_pdf import export_to_pdf
import pandas as pd
from openpyxl import load_workbook
from VARIABLES import comp_month, tm_comp_file, cs_comp_file, asd_comp_file


def export_to_excel(excel_file: str, tab: str, dataframe: pd.DataFrame):
    """Writes a pandas dataframe pulled from database.py to the specified Excel file and tab."""
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        dataframe.to_excel(writer, sheet_name=tab, index=False)


def tm_statement(payees, **kwargs: list):
    """Exports all comp details to COMP_STATEMENT.xlsx, which can then be used to generate an official statement.
    Optionally, you can pass in an email kwarg as a string to view info for a single rep, or as a list to view info
    for multiple reps. Set export=True to run a VBA script to generate a PDF statement."""
    email = kwargs.get('email', None)
    print("Generating TM Statements...")
    for tm in payees.tm_info:
        if email is None or payees.tm_info[tm]['EMAIL'] in email:
            # get the info for only the current loop rep
            name = tm
            eid = payees.tm_info[tm]['EMAIL']
            asd = payees.tm_info[tm]['RM_EMAIL']
            terr = payees.tm_info[tm]['TERR_NM']
            threshold = payees.tm_info[tm]['THRESHOLD']
            plan = payees.tm_info[tm]['PLAN']
            payout_df = payees.tblpayout[(payees.tblpayout['EID'] == eid) & (payees.tblpayout['ROLE'] == 'TM')]
            comp_detail_df = payees.tm_comp_detail[payees.tm_comp_detail['SALES_CREDIT_REP_EMAIL'] == eid]

            excel_file = tm_comp_file

            # export the rep's tblPayout info and comp detail info to different tabs on COMP_STATEMENT.xlsx
            export_to_excel(excel_file, 'payout', payout_df)
            export_to_excel(excel_file, 'detail', comp_detail_df)

            # export the rep's name, email, asd email, and territory name to the 'info' tab

            wb = load_workbook(excel_file)
            sheet_name = 'info'
            sheet = wb[sheet_name]
            sheet['B1'].value = name
            sheet['B2'].value = eid
            sheet['B3'].value = asd
            sheet['B4'].value = terr
            sheet['B5'].value = 'Territory Manager'
            sheet['B6'].value = comp_month
            sheet['B7'].value = threshold
            sheet['B8'].value = plan
            wb.save(excel_file)
            wb.close()
            # break

            # set kwarg export=True to generate a PDF statement from the Excel file
            export = kwargs.get('export', None)

            # run the specified VBA script to export as a PDF
            if export:
                export_to_pdf("AMExportPDF")
        else:
            continue


def asd_statements(payees, **kwargs: list):
    """Exports all comp details to COMP_STATEMENT_RM.xlsx, which can then be used to generate an official statement.
        Optionally, you can pass in an email kwarg to view info for a single rep. Set export=True to generate a PDF
        statement."""
    email = kwargs.get('email', None)
    print("Generating ASD Statements...")
    for asd in payees.asd_info:
        if email is None or payees.asd_info[asd]['EMAIL'] in email:
            # get the info for only the current loop rep
            name = asd
            fname = payees.asd_info[asd]['FNAME']
            lname = payees.asd_info[asd]['LNAME']
            eid = payees.asd_info[asd]['EMAIL']
            region = payees.asd_info[asd]['REGION']
            payout_df = payees.tblpayout[(payees.tblpayout['EID'] == eid) & (payees.tblpayout['ROLE'] == 'ASD')]
            comp_detail_df = payees.asd_comp_detail[payees.asd_comp_detail['SALES_CREDIT_ASD_EMAIL'] == eid]

            excel_file = asd_comp_file

            # export the rep's tblPayout info and comp detail info to different tabs on COMP_STATEMENT.xlsx
            export_to_excel(excel_file, 'payout', payout_df)
            export_to_excel(excel_file, 'detail', comp_detail_df)

            wb = load_workbook(excel_file)
            sheet_name = 'info'
            sheet = wb[sheet_name]
            sheet['B1'].value = name
            sheet['B2'].value = fname
            sheet['B3'].value = lname
            sheet['B4'].value = eid
            sheet['B5'].value = region
            sheet['B7'].value = comp_month
            wb.save(excel_file)
            wb.close()

            # set kwarg export=True to generate a PDF statement from the Excel file
            export = kwargs.get('export', None)
            if export:
                export_to_pdf("RMExportPDF")
        else:
            continue


def cs_statements(payees, **kwargs: list):
    """Exports all comp details to COMP_STATEMENT_CSR.xlsx, which can then be used to generate an official statement.
        Optionally, you can pass in an email kwarg to view info for a single rep. Set export=True to generate a PDF
        statement."""
    email = kwargs.get('email', None)
    print("Generating CS Statements...")
    for cs in payees.cs_info:
        if email is None or payees.cs_info[cs]['EMAIL'] in email:
            # get the info for only the current loop rep
            name = cs
            eid = payees.cs_info[cs]['EMAIL']
            rm = payees.cs_info[cs]['RM_EMAIL']
            terr = payees.cs_info[cs]['TERR_NM']
            base_bonus = payees.cs_info[cs]['BASE_BONUS']
            quota = payees.cs_info[cs]['PLAN']
            payout_df = payees.tblpayout[(payees.tblpayout['EID'] == eid) & (payees.tblpayout['ROLE'] == 'CS')]
            comp_detail_df = payees.cs_comp_detail[payees.cs_comp_detail['SALES_CREDIT_CS_EMAIL'] == eid]

            excel_file = cs_comp_file

            # export the rep's tblPayout info and comp detail info to different tabs on COMP_STATEMENT.xlsx
            export_to_excel(excel_file, 'payout', payout_df)
            export_to_excel(excel_file, 'detail', comp_detail_df)

            wb = load_workbook(excel_file)
            sheet_name = 'info'
            sheet = wb[sheet_name]
            sheet['B1'].value = name
            sheet['B2'].value = eid
            sheet['B3'].value = rm
            sheet['B4'].value = terr
            sheet['B6'].value = base_bonus
            sheet['B7'].value = comp_month
            sheet['B8'].value = quota
            wb.save(excel_file)
            wb.close()

            # set kwarg export=True to generate a PDF statement from the Excel file
            export = kwargs.get('export', None)
            if export:
                export_to_pdf("CSRExportPDF")
        else:
            continue
