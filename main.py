from database import Payees
from pprint import pprint
from add_watermark import add_directory_watermark
from send_email import send_am_email, send_rm_email, send_csr_email
from statements import am_statement, rm_statement, csr_statement
from VARIABLES import (am_directory, am_prelim_directory, rm_directory, rm_prelim_directory, csr_directory,
                       csr_prelim_directory, comp_mm, comp_month)
from tkinter import *

# payees = Payees()

window = Tk()
window.title("COMP")
window.config(pady=20, padx=20)


# To create only a single statement instead of looping through everyone (for debugging purposes):
# add email='employee_email' kwarg to the statement expression (after export=...)
def start():
    payees = Payees()
    if checked_state_am.get():
        am_statement(payees, export=get_radio(radio_state_export))
        if get_radio(radio_state_prelim):
            add_directory_watermark(am_directory, am_prelim_directory)
    if checked_state_csr.get():
        csr_statement(payees, export=get_radio(radio_state_export))
        if get_radio(radio_state_prelim):
            add_directory_watermark(csr_directory, csr_prelim_directory)
    if checked_state_rm.get():
        rm_statement(payees, export=get_radio(radio_state_export))
        if get_radio(radio_state_prelim):
            add_directory_watermark(rm_directory, rm_prelim_directory)
    if get_radio(radio_state_email):
        if get_radio(radio_state_prelim):
            send_am_email(payees, comp_mm, comp_month, is_prelim=True)
            send_rm_email(payees, comp_mm, comp_month, is_prelim=True)
            send_csr_email(payees, comp_mm, comp_month, is_prelim=True)
        else:
            send_am_email(payees, comp_mm, comp_month, is_prelim=False)
            send_rm_email(payees, comp_mm, comp_month, is_prelim=False)
            send_csr_email(payees, comp_mm, comp_month, is_prelim=False)
    window.destroy()


def get_radio(variable):
    radio_input = variable.get()
    return radio_input


# reminder header
reminder = Label(text="Remember to double check all variables on VARIABLES.py!!!")
reminder.pack()

# prelim statements?
prelim_label = Label(text="Are these prelim statements?", pady=10)
prelim_label.pack()

# radio buttons for prelim
radio_state_prelim = BooleanVar()
prelim_true = Radiobutton(text="TRUE", value=True, variable=radio_state_prelim)
prelim_false = Radiobutton(text="FALSE", value=False, variable=radio_state_prelim)
prelim_true.pack()
prelim_false.pack()

# header for checkbutton options
checkbutton_label = Label(text="Create statements for which roles?", pady=10)
checkbutton_label.pack()

# checkbutton options
checked_state_am = BooleanVar()
checked_state_csr = BooleanVar()
checked_state_rm = BooleanVar()
am_button = Checkbutton(text="AMs", variable=checked_state_am)
am_button.pack()
csr_button = Checkbutton(text="CSRs", variable=checked_state_csr)
csr_button.pack()
rm_button = Checkbutton(text="RMs", variable=checked_state_rm)
rm_button.pack()

# header for radio button options
export_label = Label(text="Export statements to PDF?", pady=10)
export_label.pack(side='top')

# radio buttons for export option
radio_state_export = BooleanVar()
radio_state_export.set(True)
export_true = Radiobutton(text="TRUE", value=True, variable=radio_state_export)
export_false = Radiobutton(text="FALSE", value=False, variable=radio_state_export)
export_true.pack()
export_false.pack()

# header for emails
email_label = Label(text="Send emails?", pady=10)
email_label.pack()

# radio buttons for email
radio_state_email = BooleanVar()
email_true = Radiobutton(text="TRUE", value=True, variable=radio_state_email)
email_false = Radiobutton(text="FALSE", value=False, variable=radio_state_email)
email_true.pack()
email_false.pack()

run = Button(text="GO", command=start, padx=10)
run.pack()

window.mainloop()

#########################################################
# pre-GUI version:

# BEFORE RUNNING: update VARIABLES.py and check VARIABLES.py comp notes

# payees = Payees()

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
