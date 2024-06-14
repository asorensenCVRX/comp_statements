# EVERY MONTH: change export path by editing the VBA in export_statements.xlsm

comp_mm = '05'
comp_month = 'May'

tms = ['ccraigo@cvrx.com', 'dduffy@cvrx.com', 'ecorson@cvrx.com', 'jbuxton@cvrx.com', 'jclemmons@cvrx.com',
       'jrussell@cvrx.com', 'jsantoli@cvrx.com', 'jwyatt@cvrx.com', 'tbarker@cvrx.com']

rms = ['jgarner@cvrx.com', 'kdenton@cvrx.com', 'jhorky@cvrx.com', 'ccastillo@cvrx.com']

am_prelim_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\AM Prelim msg.oft"
rm_prelim_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\RM Prelim msg.oft"
am_official_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\AM msg.oft"
rm_official_email = r"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\RM msg.oft"

am_directory = fr"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\2024_{comp_mm}\AM"
am_prelim_directory = am_directory + r"\PRELIMINARIES"
rm_directory = fr"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\2024_{comp_mm}\RM"
rm_prelim_directory = rm_directory + r"\PRELIMINARIES"
csr_directory = fr"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS\2024_{comp_mm}\CSR"
csr_prelim_directory = csr_directory + r"\PRELIMINARIES"
