import win32com.client
import html
import os
from VARIABLES import (am_prelim_email, rm_prelim_email, am_official_email, rm_official_email, am_prelim_directory,
                       am_directory, csr_prelim_directory, csr_directory, rm_prelim_directory, rm_directory)


def send_am_email(payees, month_mm: str, month_name: str, is_prelim: bool):
    print("Sending AM prelim emails...") if is_prelim \
        else print("Sending AM official emails...")
    for key, value in payees.am_info.items():
        try:
            folder = am_prelim_directory if is_prelim \
                else am_directory
            file_name = f'PRELIMINARY_{key}_2024_{month_mm}.pdf' if is_prelim \
                else f'{key}_2024_{month_mm}.pdf'
            path = os.path.join(folder, file_name)
            subject = f"PRELIMINARY {month_name} Comp Statement: {value['TERR_NM']}" if is_prelim \
                else f"{month_name} Comp Statement: {value['TERR_NM']}"
            # Check if rep reports to a TM; if so, add the rep's TM and RM as managers so both get CC'd
            manager_email = None
            for tm, rep in payees.tm_reports.items():
                if value["EMAIL"] in rep:
                    manager_email = f"{value['RM_EMAIL']}; {tm}"
                    break
                else:
                    manager_email = value['RM_EMAIL']
            template = am_prelim_email if is_prelim \
                else am_official_email
            email = SendEmail(template=template, recipient_fullname=key, recipient_first_name=value['FNAME_REP'],
                              recipient_email=value['EMAIL'], manager_email=manager_email, subject=subject,
                              attachment_path=path)
            email.send_email()
        except Exception as e:
            print(f"There was an error {e}.\nUnable to send email to {key}: {value}")
            continue


def send_rm_email(payees, month_mm: str, month_name: str, is_prelim: bool):
    print("Sending RM prelim emails...") if is_prelim \
        else print("Sending RM official emails...")
    for key, value in payees.rm_info.items():
        try:
            folder = rm_prelim_directory if is_prelim \
                else rm_directory
            file_name = f'PRELIMINARY_{key}_2024_{month_mm}.pdf' if is_prelim \
                else f"{key}_2024_{month_mm}.pdf"
            path = os.path.join(folder, file_name)
            subject = f"PRELIMINARY {month_name} Comp Statement: {value['REGION']}" if is_prelim \
                else f"{month_name} Comp Statement: {value['REGION']}"
            template = rm_prelim_email if is_prelim \
                else rm_official_email
            email = SendEmail(template=template, recipient_fullname=key, recipient_first_name=value['FNAME'],
                              recipient_email=value['EMAIL'], manager_email='rjohn@cvrx.com', subject=subject,
                              attachment_path=path)
            email.send_email()
        except Exception as e:
            print(f"There was an error {e}.\nUnable to send email to {key}: {value}")
            continue


def send_csr_email(payees, month_mm: str, month_name: str, is_prelim: bool):
    print("Sending CSR prelim emails...") if is_prelim \
        else print("Sending CSR official emails...")
    for key, value in payees.csr_info.items():
        try:
            folder = csr_prelim_directory if is_prelim \
                else csr_directory
            file_name = f'PRELIMINARY_{key}_2024_{month_mm}.pdf' if is_prelim \
                else f'{key}_2024_{month_mm}.pdf'
            path = os.path.join(folder, file_name)
            subject = f"PRELIMINARY {month_name} Comp Statement: {value['TERR_NM']}" if is_prelim \
                else f"{month_name} Comp Statement: {value['TERR_NM']}"
            template = am_prelim_email if is_prelim \
                else am_official_email
            email = SendEmail(template=template, recipient_fullname=key, recipient_first_name=value['FNAME_REP'],
                              recipient_email=value['EMAIL'], manager_email=value['RM_EMAIL'], subject=subject,
                              attachment_path=path)
            email.send_email()
        except Exception as e:
            print(f"There was an error {e}.\nUnable to send email to {key}: {value}")
            continue


class SendEmail:
    def __init__(self, template, recipient_fullname, recipient_first_name, recipient_email, manager_email, subject,
                 attachment_path):
        self.email_template = template
        self.subject = subject
        self.email = recipient_email
        self.recipient_fullname = recipient_fullname
        self.fname = recipient_first_name
        self.attachment = fr'{attachment_path}'
        self.manager_email = manager_email

    def send_email(self):
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItemFromTemplate(self.email_template)
        replace_text_decoded = html.unescape(self.fname)
        # mail.BodyFormat = 2
        # for the .replace statement below, you must change "<Rep Name>" to HTML. Result is "&lt;Rep Name&gt;"
        mail.HTMLBody = mail.HTMLBody.replace('&lt;Rep Name&gt;', replace_text_decoded)
        mail.Subject = self.subject
        mail.To = self.email
        if self.manager_email is not None:
            mail.CC = f"{self.manager_email}; jmoore@cvrx.com"
        else:
            mail.CC = "jmoore@cvrx.com"
        mail.Attachments.Add(self.attachment)
        mail.Send()
