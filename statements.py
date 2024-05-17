from export_pdf import export_to_pdf
import pandas as pd
from openpyxl import load_workbook

tms = ['ccraigo@cvrx.com', 'dduffy@cvrx.com', 'ecorson@cvrx.com', 'jbuxton@cvrx.com', 'jclemmons@cvrx.com',
       'jrussell@cvrx.com', 'jsantoli@cvrx.com', 'jwyatt@cvrx.com', 'tbarker@cvrx.com']


def export_to_excel(excel_file: str, tab: str, dataframe: pd.DataFrame):
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        dataframe.to_excel(writer, sheet_name=tab, index=False)


def am_statement(payees, **kwargs):
    """Exports all comp details to COMP_STATEMENT.xlsx, which can then be used to generate an official statement.
    Optionally, you can pass in an email kwarg to view info for a single rep. Set export=True to generate a PDF
    statement."""
    email = kwargs.get('email', None)
    for am in payees.am_info:
        # get the info for only the current loop rep
        name = None if email else am
        eid = email if email else payees.am_info[am]['EMAIL']
        rm = None if email else payees.am_info[am]['RM_EMAIL']
        terr = None if email else payees.am_info[am]['TERR_NM']
        payout_df = payees.tblpayout[payees.tblpayout['EID'] == eid]
        comp_detail_df = payees.am_comp_detail[payees.am_comp_detail['SALES_CREDIT_REP_EMAIL'] == eid]

        excel_file = r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\PyWorkbooks & Queries\COMP_STATEMENT.xlsx"

        # export the rep's tblPayout info and comp detail info to different tabs on COMP_STATEMENT.xlsx
        export_to_excel(excel_file, 'payout', payout_df)
        export_to_excel(excel_file, 'detail', comp_detail_df)

        # export the rep's name, email, rm email, and territory name to the 'info' tab

        wb = load_workbook(excel_file)
        sheet_name = 'info'
        sheet = wb[sheet_name]
        sheet['B1'].value = name
        sheet['B2'].value = eid
        sheet['B3'].value = rm
        sheet['B4'].value = terr
        if payees.am_info[am]['EMAIL'] in tms:
            sheet['B5'].value = 'Territory Manager'
        else:
            sheet['B5'].value = 'Account Manager'
        wb.save(excel_file)
        wb.close()
        # break

        # set kwarg export=True to generate a PDF statement from the Excel file
        export = kwargs.get('export', None)
        if export:
            export_to_pdf()
        if email is not None:
            break
