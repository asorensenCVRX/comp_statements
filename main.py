from database import Payees, get_tblpayout
from pprint import pprint
import pandas as pd

payees = Payees()
print(payees.rm_comp_detail)

# for am in payees.am_info:
#     eid = payees.am_info[am]['EMAIL']
#     payees.tblpayout[payees.tblpayout['EID'] == eid].to_excel('./testing.xlsx', 'payout')
#     break


# TODO 3: Using tmp tables, create an Excel file pulling the necessary data for comp statements


# TODO 4: Generate a PDF comp statement based on the data from the Excel file
