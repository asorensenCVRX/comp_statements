import os.path
from database import Payees
from pprint import pprint
from add_watermark import add_directory_watermark
from send_email import SendEmail, send_am_prelim_email
from statements import am_statement

# remember to change the templates to have the current comp month
payees = Payees()

# change export path by editing the VBA in export_statements.xlsm
am_statement(payees, export=True, email='jclemmons@cvrx.com')

# add watermark to prelim statements
# add_directory_watermark()

# send AM preliminary emails
# send_am_prelim_email(payees, '04', 'April')
