import os.path
from database import Payees
from pprint import pprint
from add_watermark import add_directory_watermark
from send_email import SendEmail, send_am_prelim_email, send_rm_prelim_email, send_csr_prelim_email
from statements import am_statement, rm_statement, csr_statement

# remember to change the templates to have the current comp month
payees = Payees()

# change export path by editing the VBA in export_statements.xlsm
# am_statement(payees, export=True, email='jclemmons@cvrx.com')
# rm_statement(payees, export=True)
csr_statement(payees, export=True)


# add watermark to prelim statements
add_directory_watermark(r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\Statements\CSR\2024_04",
                        r"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\Statements\CSR\2024_04\PRELIMINARIES")

# send preliminary emails
# send_am_prelim_email(payees, '04', 'April')
# send_rm_prelim_email(payees, '04', 'April')
send_csr_prelim_email(payees, '04', 'April')
