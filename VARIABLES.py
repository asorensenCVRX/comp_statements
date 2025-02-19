# EVERY MONTH: change export path by editing the VBA in export_statements.xlsm

comp_mm = '01'
comp_month = 'January'

# email module
rep_prelim_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2025_COMP_OPS\COMP_STATEMENTS\TM Prelim msg.oft"
asd_prelim_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2025_COMP_OPS\COMP_STATEMENTS\ASD Prelim msg.oft"
rep_official_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2025_COMP_OPS\COMP_STATEMENTS\TM msg.oft"
asd_official_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2025_COMP_OPS\COMP_STATEMENTS\ASD msg.oft"

# exported pdf directories
tm_directory = fr"C:\Users\asorensen\OneDrive - CVRx Inc\2025_COMP_OPS\COMP_STATEMENTS\2025_{comp_mm}\TM"
tm_prelim_directory = tm_directory + r"\PRELIMINARIES"
asd_directory = fr"C:\Users\asorensen\OneDrive - CVRx Inc\2025_COMP_OPS\COMP_STATEMENTS\2025_{comp_mm}\ASD"
asd_prelim_directory = asd_directory + r"\PRELIMINARIES"
cs_directory = fr"C:\Users\asorensen\OneDrive - CVRx Inc\2025_COMP_OPS\COMP_STATEMENTS\2025_{comp_mm}\CS"
cs_prelim_directory = cs_directory + r"\PRELIMINARIES"

# excel files for generating statements
vba_excel_file = r"C:\Users\asorensen\PycharmProjects\comp_statements\Excel Files\export_statements.xlsm"
tm_comp_file = r"C:\Users\asorensen\PycharmProjects\comp_statements\Excel Files\COMP_STATEMENT.xlsx"
cs_comp_file = r"C:\Users\asorensen\PycharmProjects\comp_statements\Excel Files\COMP_STATEMENT_CS.xlsx"
asd_comp_file = r"C:\Users\asorensen\PycharmProjects\comp_statements\Excel Files\COMP_STATEMENT_ASD.xlsx"

# SQL files
payout_table = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\tblPayout.sql"
tm_info = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\TM_INFO.sql"
cs_info = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\CS_INFO.sql"
tm_comp = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\comp_TM.sql"
cs_comp = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\comp_CS.sql"
