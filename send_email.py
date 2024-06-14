# must pip install pywin32
import win32com.client
import html
import os
from VARIABLES import am_prelim_email, rm_prelim_email, am_official_email, rm_official_email


def send_am_prelim_email(payees, month_mm, month_name):
    print("Sending AM Prelim Emails...")
    for key, value in payees.am_info.items():
        folder = fr"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\Statements\AM\2024_{month_mm}\PRELIMINARIES"
        file_name = f'PRELIMINARY_{key}_2024_{month_mm}.pdf'
        path = os.path.join(folder, file_name)
        subject = f"PRELIMINARY {month_name} Comp Statement: {value['TERR_NM']}"
        email = SendEmail(am_prelim_email, key, value['FNAME_REP'], value['EMAIL'], value['RM_EMAIL'], subject, path)
        email.send_email()


def send_rm_prelim_email(payees, month_mm, month_name):
    print("Sending RM Prelim Emails...")
    for key, value in payees.rm_info.items():
        folder = fr"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\Statements\RM\2024_{month_mm}\PRELIMINARIES"
        file_name = f'PRELIMINARY_{key}_2024_{month_mm}.pdf'
        path = os.path.join(folder, file_name)
        subject = f"PRELIMINARY {month_name} Comp Statement: {value['REGION']}"
        email = SendEmail(rm_prelim_email, key, value['FNAME'], value['EMAIL'], None, subject, path)
        email.send_email()


def send_csr_prelim_email(payees, month_mm, month_name):
    print("Sending CSR Prelim Emails...")
    for key, value in payees.csr_info.items():
        folder = fr"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\Statements\CSR\2024_{month_mm}\PRELIMINARIES"
        file_name = f'PRELIMINARY_{key}_2024_{month_mm}.pdf'
        path = os.path.join(folder, file_name)
        subject = f"PRELIMINARY {month_name} Comp Statement: {value['TERR_NM']}"
        email = SendEmail(am_prelim_email, key, value['FNAME_REP'], value['EMAIL'], value['RM_EMAIL'], subject, path)
        email.send_email()


def send_am_official_email(payees, month_mm, month_name, **kwargs):
    """Use the kwarg skip_reps='email' in a list form to skip sending comp statements to specific reps."""
    for key, value in payees.am_info.items():
        folder = fr"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\Statements\AM\2024_{month_mm}"
        file_name = f'{key}_2024_{month_mm}.pdf'
        path = os.path.join(folder, file_name)
        subject = f"{month_name} Comp Statement: {value['TERR_NM']}"
        email = SendEmail(am_official_email, key, value['FNAME_REP'], value['EMAIL'], value['RM_EMAIL'], subject,
                          path)
        email.send_email()
        print(f"Email sent to {value['EMAIL']}")


def send_rm_official_email(payees, month_mm, month_name):
    for key, value in payees.rm_info.items():
        folder = fr"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\Statements\RM\2024_{month_mm}"
        file_name = f'{key}_2024_{month_mm}.pdf'
        path = os.path.join(folder, file_name)
        subject = f"{month_name} Comp Statement: {value['REGION']}"
        email = SendEmail(rm_official_email, key, value['FNAME'], value['EMAIL'], None, subject, path)
        email.send_email()
        print(f"Email sent to {value['EMAIL']}")


def send_csr_official_email(payees, month_mm, month_name):
    for key, value in payees.csr_info.items():
        folder = fr"C:\Users\asorensen\OneDrive - CVRx Inc\Calculate Comp\Statements\CSR\2024_{month_mm}"
        file_name = f'{key}_2024_{month_mm}.pdf'
        path = os.path.join(folder, file_name)
        subject = f"{month_name} Comp Statement: {value['TERR_NM']}"
        email = SendEmail(am_official_email, key, value['FNAME_REP'], value['EMAIL'], value['RM_EMAIL'], subject, path)
        email.send_email()
        print(f"Email sent to {value['EMAIL']}")


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
        # self.send_email()

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
