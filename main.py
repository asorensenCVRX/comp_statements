from database import Payees
from pprint import pprint
from add_watermark import add_directory_watermark
from send_email import (send_am_prelim_email, send_rm_prelim_email, send_csr_prelim_email,
                        send_am_official_email, send_rm_official_email, send_csr_official_email)
from statements import am_statement, rm_statement, csr_statement
from VARIABLES import (am_directory, am_prelim_directory, rm_directory, rm_prelim_directory, csr_directory,
                       csr_prelim_directory, comp_mm, comp_month)

#########################################################

# BEFORE RUNNING: update VARIABLES.py and check VARIABLES.py comp notes

payees = Payees()

# step 1: create comp statements
# am_statement(payees, export=True)
# rm_statement(payees, export=True)
# csr_statement(payees, export=True)

##################################################

# PRELIMS
# step 2: add watermark to statements to create prelim statements.
# add_directory_watermark(am_directory, am_prelim_directory)
# add_directory_watermark(rm_directory, rm_prelim_directory)
# add_directory_watermark(csr_directory, csr_prelim_directory)

# step 3: send preliminary statement emails
# send_am_prelim_email(payees, comp_mm, comp_month)
# send_rm_prelim_email(payees, comp_mm, comp_month)
# send_csr_prelim_email(payees, comp_mm, comp_month)

##################################################

# OFFICIAL
# step 2: send official comp statement emails
# send_am_official_email(payees, comp_mm, comp_month)
# send_rm_official_email(payees, comp_mm, comp_month)
# send_csr_official_email(payees, comp_mm, comp_month)
