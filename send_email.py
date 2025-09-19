import win32com.client
import html
import os
from VARIABLES import (rep_prelim_email, asd_prelim_email, rep_official_email, asd_official_email, tm_prelim_directory,
                       tm_directory, cs_prelim_directory, cs_directory, asd_prelim_directory, asd_directory, atm_directory,
                       atm_official_email)


def send_tm_email(payees, month_mm: str, month_name: str, is_prelim: bool):
    print("Sending TM prelim emails...") if is_prelim \
        else print("Sending TM official emails...")
    for key, value in payees.tm_info.items():
        try:
            folder = tm_prelim_directory if is_prelim \
                else tm_directory
            file_name = f'PRELIMINARY_{key}_2025_{month_mm}.pdf' if is_prelim \
                else f'{key}_2025_{month_mm}.pdf'
            path = os.path.join(folder, file_name)
            subject = f"PRELIMINARY {month_name} Comp Statement: {value['TERR_NM']}" if is_prelim \
                else f"{month_name} Comp Statement: {value['TERR_NM']}"

            manager_email = value['RM_EMAIL']
            template = rep_prelim_email if is_prelim \
                else rep_official_email
            email = SendEmail(template=template, recipient_fullname=key, recipient_first_name=value['FNAME_REP'],
                              recipient_email=value['EMAIL'], manager_email=manager_email, subject=subject,
                              attachment_path=path)
            email.send_email()
        except Exception as e:
            print(f"There was an error {e}.\nUnable to send email to {key}: {value}")
            continue


def send_asd_email(payees, month_mm: str, month_name: str, is_prelim: bool):
    print("Sending ASD prelim emails...") if is_prelim \
        else print("Sending ASD official emails...")
    for key, value in payees.asd_info.items():
        try:
            folder = asd_prelim_directory if is_prelim \
                else asd_directory
            file_name = f'PRELIMINARY_{value["FNAME"]} {value["LNAME"]}_2025_{month_mm}.pdf' if is_prelim \
                else f"{value["FNAME"]} {value["LNAME"]}_2025_{month_mm}.pdf"
            path = os.path.join(folder, file_name)
            subject = f"PRELIMINARY {month_name} Comp Statement: {value['REGION']}" if is_prelim \
                else f"{month_name} Comp Statement: {value['REGION']}"
            template = asd_prelim_email if is_prelim \
                else asd_official_email
            email = SendEmail(template=template, recipient_fullname=key, recipient_first_name=value['FNAME'],
                              recipient_email=value['EMAIL'], manager_email='rjohn@cvrx.com', subject=subject,
                              attachment_path=path)
            email.send_email()
        except Exception as e:
            print(f"There was an error {e}.\nUnable to send email to {key}: {value}")
            continue


def send_cs_email(payees, month_mm: str, month_name: str, is_prelim: bool):
    print("Sending CS prelim emails...") if is_prelim \
        else print("Sending CS official emails...")
    for key, value in payees.cs_info.items():
        try:
            folder = cs_prelim_directory if is_prelim \
                else cs_directory
            file_name = f'PRELIMINARY_{key}_2025_{month_mm}.pdf' if is_prelim \
                else f'{key}_2025_{month_mm}.pdf'
            path = os.path.join(folder, file_name)
            subject = f"PRELIMINARY {month_name} Comp Statement: {value['TERR_NM']}" if is_prelim \
                else f"{month_name} Comp Statement: {value['TERR_NM']}"
            template = rep_prelim_email if is_prelim \
                else rep_official_email
            email = SendEmail(template=template, recipient_fullname=key, recipient_first_name=value['FNAME_REP'],
                              recipient_email=value['EMAIL'], manager_email=value['RM_EMAIL'], subject=subject,
                              attachment_path=path)
            email.send_email()
        except Exception as e:
            print(f"There was an error {e}.\nUnable to send email to {key}: {value}")
            continue


def send_atm_email(payees, month_mm: str, month_name: str, is_prelim: bool):
    if is_prelim:
        pass
    else:
        print("Sending ATM emails...")
        for key, value in payees.atm_info.items():
            try:
                file_name = f'{key}_2025_{month_mm}.pdf'
                path = os.path.join(atm_directory, file_name)
                subject = f"{month_name} Comp Statement: {value['TERR_NM']}"

                manager_email = value['RM_EMAIL']
                template = atm_official_email
                email = SendEmail(template=template, recipient_fullname=key, recipient_first_name=value['FNAME_REP'],
                                  recipient_email=value['EMAIL'], manager_email=manager_email, subject=subject,
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
            mail.CC = f"{self.manager_email}"
        mail.Attachments.Add(self.attachment)
        mail.Send()
