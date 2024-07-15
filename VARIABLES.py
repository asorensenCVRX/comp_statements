# EVERY MONTH: change export path by editing the VBA in export_statements.xlsm

comp_mm = '06'
comp_month = 'June'

tms = ['ccraigo@cvrx.com', 'dduffy@cvrx.com', 'ecorson@cvrx.com', 'jbuxton@cvrx.com', 'jclemmons@cvrx.com',
       'jrussell@cvrx.com', 'jsantoli@cvrx.com', 'jwyatt@cvrx.com', 'tbarker@cvrx.com']

rms = ['jgarner@cvrx.com', 'kdenton@cvrx.com', 'jhorky@cvrx.com', 'ccastillo@cvrx.com']

# email module
am_prelim_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\AM Prelim msg.oft"
rm_prelim_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\RM Prelim msg.oft"
am_official_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\AM msg.oft"
rm_official_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\RM msg.oft"

# exported pdf directories
am_directory = fr"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\2024_{comp_mm}\AM"
am_prelim_directory = am_directory + r"\PRELIMINARIES"
rm_directory = fr"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\2024_{comp_mm}\RM"
rm_prelim_directory = rm_directory + r"\PRELIMINARIES"
csr_directory = fr"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\2024_{comp_mm}\CSR"
csr_prelim_directory = csr_directory + r"\PRELIMINARIES"

# excel files for generating statements
vba_excel_file = r"C:\Users\asorensen\PycharmProjects\comp_statements\Excel Files\export_statements.xlsm"
am_comp_file = r"C:\Users\asorensen\PycharmProjects\comp_statements\Excel Files\COMP_STATEMENT.xlsx"
csr_comp_file = r"C:\Users\asorensen\PycharmProjects\comp_statements\Excel Files\COMP_STATEMENT_CSR.xlsx"
rm_comp_file = r"C:\Users\asorensen\PycharmProjects\comp_statements\Excel Files\COMP_STATEMENT_RM.xlsx"

# SQL files
payout_table = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\tblPayout.sql"
am_info = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\AM_INFO.sql"
fce_info = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\FCE_INFO.sql"
am_comp = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\comp_AM.sql"
fce_comp = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\comp_FCE.sql"
tm_reports = r"C:\Users\asorensen\PycharmProjects\comp_statements\SQL Queries\TM_direct_reports.sql"
