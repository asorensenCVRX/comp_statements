from export_pdf import export_to_pdf
import pandas as pd
from openpyxl import load_workbook
from VARIABLES import comp_month, tm_comp_file, cs_comp_file, asd_comp_file, atm_comp_file
from database import payees


def export_to_excel(excel_file: str, tab: str, dataframe: pd.DataFrame):
    """Writes a pandas dataframe pulled from database.py to the specified Excel file and tab."""
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        dataframe.to_excel(writer, sheet_name=tab, index=False)


def statement(role, email: list[str] | None = None, export: bool = False):
    if role == 'TM':
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
                recoverable_draw = payees.tm_info[tm]['RECOVERABLE_DRAW']
                payout_df = payees.tblpayout[(payees.tblpayout['EID'] == eid) & (payees.tblpayout['ROLE'] == 'TM')]
                comp_detail_df = payees.tm_comp_detail[payees.tm_comp_detail['SALES_CREDIT_REP_EMAIL'] == eid]

                # export the rep's tblPayout info and comp detail info to different tabs on COMP_STATEMENT.xlsx
                export_to_excel(tm_comp_file, 'payout', payout_df)
                export_to_excel(tm_comp_file, 'detail', comp_detail_df)

                # export the rep's name, email, asd email, and territory name to the 'info' tab

                wb = load_workbook(tm_comp_file)
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
                sheet['B9'].value = recoverable_draw
                wb.save(tm_comp_file)
                wb.close()
                # break

                # run the specified VBA script to export as a PDF
                if export:
                    export_to_pdf("AMExportPDF")
            else:
                continue

    elif role == 'CS':
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

                # export the rep's tblPayout info and comp detail info to different tabs on COMP_STATEMENT.xlsx
                export_to_excel(cs_comp_file, 'payout', payout_df)
                export_to_excel(cs_comp_file, 'detail', comp_detail_df)

                wb = load_workbook(cs_comp_file)
                sheet_name = 'info'
                sheet = wb[sheet_name]
                sheet['B1'].value = name
                sheet['B2'].value = eid
                sheet['B3'].value = rm
                sheet['B4'].value = terr
                sheet['B6'].value = base_bonus
                sheet['B7'].value = comp_month
                sheet['B8'].value = quota
                wb.save(cs_comp_file)
                wb.close()

                if export:
                    export_to_pdf("CSRExportPDF")
            else:
                continue

    elif role == 'ASD':
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

                # export the rep's tblPayout info and comp detail info to different tabs on COMP_STATEMENT.xlsx
                export_to_excel(asd_comp_file, 'payout', payout_df)
                export_to_excel(asd_comp_file, 'detail', comp_detail_df)

                wb = load_workbook(asd_comp_file)
                sheet_name = 'info'
                sheet = wb[sheet_name]
                sheet['B1'].value = name
                sheet['B2'].value = fname
                sheet['B3'].value = lname
                sheet['B4'].value = eid
                sheet['B5'].value = region
                sheet['B7'].value = comp_month
                wb.save(asd_comp_file)
                wb.close()

                if export:
                    export_to_pdf("RMExportPDF")
            else:
                continue

    elif role == 'ATM':
        print("Generating ATM Statements...")
        for atm in payees.atm_info:
            if email is None or payees.atm_info[atm]['EMAIL'] in email:
                # get the info for only the current loop rep
                name = atm
                eid = payees.atm_info[atm]['EMAIL']
                asd = payees.atm_info[atm]['RM_EMAIL']
                terr = payees.atm_info[atm]['TERR_NM']
                payout_df = payees.tblpayout[(payees.tblpayout['EID'] == eid) & (payees.tblpayout['ROLE'] == 'ATM')]
                comp_detail_df = payees.atm_comp_detail[payees.atm_comp_detail['ATM_EMAIL'] == eid]

                # export the rep's tblPayout info and comp detail info to different tabs on COMP_STATEMENT.xlsx
                export_to_excel(atm_comp_file, 'payout', payout_df)
                export_to_excel(atm_comp_file, 'detail', comp_detail_df)

                # export the rep's name, email, asd email, and territory name to the 'info' tab

                wb = load_workbook(atm_comp_file)
                sheet_name = 'info'
                sheet = wb[sheet_name]
                sheet['B1'].value = name
                sheet['B2'].value = eid
                sheet['B3'].value = asd
                sheet['B4'].value = terr
                sheet['B5'].value = 'Associate Territory Manager'
                sheet['B6'].value = comp_month
                wb.save(atm_comp_file)
                wb.close()
                # break

                # run the specified VBA script to export as a PDF
                if export:
                    export_to_pdf("ATMExportPDF")
            else:
                continue

    else:
        print("Please specify a role to generate statements for. (TM, ATM, CS, ASD)")
