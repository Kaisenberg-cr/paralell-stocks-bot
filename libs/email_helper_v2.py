"""Python script to configure email communication"""

import inspect
import os
import smtplib

# import imaplib
import sys
import time
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

import psutil
import win32com.client


class EmailHelper:
    """Class to handle all the communication with emails"""

    @staticmethod
    def send_email(
        email_from,
        email_to,
        subject,
        body,
        cc=False,
        bcc=False,
        html_body=False,
        html_images=[],
        attachments=None,
    ):
        """Method to send emails using the SMTP Protocol"""
        try:
            print("Creating the email message...")
            message = MIMEMultipart()
            message["From"] = email_from
            rcpt = []
            if type(email_to) == str:
                message["To"] = email_to
                rcpt.append(email_to)
            else:
                email_to_string = "; ".join(email_to)
                message["To"] = email_to_string
                rcpt.extend(email_to)
            message["Date"] = formatdate(localtime=True)
            message["Subject"] = subject
            if cc:
                if type(cc) == str:
                    message["Cc"] = cc
                    rcpt.append(cc)
                else:
                    cc_to_string = "; ".join(cc)
                    message["Cc"] = cc_to_string
                    rcpt.extend(cc)
            if bcc:
                if type(bcc) == str:
                    message["Bcc"] = bcc
                    rcpt.append(bcc)
                else:
                    bcc_to_string = "; ".join(bcc)
                    message["Bcc"] = bcc_to_string
                    rcpt.extend(bcc)
            if html_body:
                if html_images:
                    index = 1
                    for item in html_images:
                        image1 = MIMEImage(open(item, "rb").read())
                        image1.add_header("Content-ID", f"MyId{index}")
                        message.attach(image1)
                        index += 1
                mt_html = MIMEText(body, "html")
                message.attach(mt_html)
                # This part below is unique to the bot, Use of needed to add images to HTML
                # image1 = MIMEImage(open("C:\\RPA\\BIS\\BIS127\\BIS127 - Target IQ Welcome E-mail\\Files\\HTML Images\\EmailImage1.jpg", 'rb').read())
                # Define the image's ID as referenced in the HTML body above
                # image1.add_header('Content-ID', 'MyId1')
                # message.attach(image1)
                # This part above is unique to the bot, Use of needed to add images to HTML
            else:
                message.attach(MIMEText(body))
            if attachments is not None:
                for file in attachments:
                    file_basename = os.path.basename(file)
                    print("Attaching ", file_basename, " file")
                    with open(file, "rb") as fil:
                        part = MIMEApplication(fil.read(), Name=file_basename)
                    part["Content-Disposition"] = (
                        'attachment; filename="%s"' % file_basename
                    )
                    message.attach(part)
                    print("Attached ", file_basename, " file")
            print("Email Message created")
            print("Start sending email message...")
            try:
                print("Start connecting to the SMTP service...")
                with smtplib.SMTP("relay.experian.com", 25, timeout=5.00) as smtp_obj:
                    print("Conected to the SMTP service...")
                    smtp_obj.set_debuglevel(2)
                    smtp_obj.sendmail(email_from, rcpt, message.as_string())
                    smtp_obj.quit()
                    print("Email sent successfully")
                    return True
            except Exception as ex:
                print("Email message not sent")
                raise Exception("Not able to conect to SMTP service")
        except Exception as ex:
            error_line = sys.exc_info()[-1].tb_lineno
            print(
                "An error has occurred while sending the email, error details:"
                + str(ex)
                + ". Error line :"
                + str(error_line)
            )
            return False

    def send_error_email(
        email_from, email_to, bot_ID, error_note, cc=False, attachments=None
    ):
        """Method to send emails using the SMTP Protocol"""
        try:
            subject = "An error occurred on " + bot_ID
            print("Creating the email message...")
            message = MIMEMultipart()
            message["From"] = email_from
            if type(email_to) == str:
                message["To"] = email_to
            else:
                email_to = "; ".join(email_to)
                message["To"] = email_to
            message["Date"] = formatdate(localtime=True)
            message["Subject"] = subject
            rcpt = [email_to]
            if cc:
                message["Cc"] = cc
                rcpt.append(cc)
            body = (
                "Hello "
                "\n"
                "\n"
                "There was an error on line: "
                + str(error_note.__traceback__.tb_lineno)
                + " Please review the below error message: "
                "\n"
                "\n" + str(error_note) + "\n"
                "\n"
                "Thank you!"
            )
            message.attach(MIMEText(body))
            if attachments is not None:
                for file in attachments:
                    file_basename = os.path.basename(file)
                    print("Start attaching", file_basename, "file")
                    with open(file, "rb") as fil:
                        part = MIMEApplication(fil.read(), Name=file_basename)
                    part["Content-Disposition"] = (
                        'attachment; filename="%s"' % file_basename
                    )
                    message.attach(part)
                    print("Attached ", file_basename, " file")
            print("Email Message created")
            print("Start sending email message...")
            try:
                print("Start connecting to the SMTP service...")
                with smtplib.SMTP("relay.experian.com", 25, timeout=5.00) as smtp_obj:
                    print("Conected to the SMTP service...")
                    smtp_obj.set_debuglevel(2)
                    smtp_obj.sendmail(email_from, rcpt, message.as_string())
                    smtp_obj.quit()
                    print("Email sent successfully")
                    # Call your logfile method here
                    return True
            except Exception as ex:
                print("Email message not sent")
                raise Exception("Not able to conect to SMTP service")
        except Exception as ex:
            error_line = sys.exc_info()[-1].tb_lineno
            print(
                "An error has occurred while sending the email, error details:"
                + str(ex)
                + ". Error line :"
                + str(error_line)
            )
            return False

    @staticmethod
    def send_outlook_email(
        self,
        outlook_account_email_from,
        subject,
        email_body,
        email_to,
        send_on_behalf_email=None,
        attachment=None,
        html=False,
        image=False,
        email_cc=None,
    ):
        """This method is used to send e-mails from an Outlook account"""
        try:
            sent_flag = False
            print(f"Start {inspect.stack()[0][3]}")
            outlook = win32com.client.Dispatch("Outlook.Application")
            account_to_use = None
            for account in outlook.Session.Accounts:
                if account.DisplayName == outlook_account_email_from:
                    account_to_use = account
                    break

            if account_to_use is None:
                raise Exception(
                    f"Could not find account with display name {outlook_account_email_from}"
                )

            mail = outlook.CreateItem(0)
            mail.SendUsingAccount = account_to_use
            if send_on_behalf_email:
                mail.SentOnBehalfOfName = send_on_behalf_email
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account_to_use))
            # Set Email Recipients
            mail.To = email_to
            if email_cc is not None:
                mail.CC = email_cc  # Set CC address
            if subject is not None:
                mail.Subject = subject  # Set Subject
            if not html:
                mail.Body = email_body  # Set body of email
            else:
                mail.HTMLBody = email_body
            if image:
                if hasattr(self, "img_header"):
                    embed1 = mail.Attachments.Add(self.img_header)
                    embed1.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                        "Header",
                    )
                if hasattr(self, "img_footer"):
                    embed2 = mail.Attachments.Add(self.img_footer)
                    embed2.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                        "Footer",
                    )
                if hasattr(self, "img_banner"):
                    embed3 = mail.Attachments.Add(self.img_banner)
                    embed3.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                        "Banner",
                    )
            if attachment:
                for attach in attachment:
                    self.log_file(f"Adding attachment {attach}")
                    mail.Attachments.Add(
                        attach
                    )  # If attachment is going to be sent out
            mail.Send()
            sent_flag = True
            logged_message = f"Email sent to {email_to} with subject {subject}."
            print(logged_message)
            # Call your Logfile method here
            self.utils.log_to_file(
                f"Email has been sent to: {email_to}\n\
                                    CC: {email_cc}\n\
                                    Subject: {subject}"
            )
            return sent_flag
        except Exception as err:
            err = str(err)
            method = inspect.stack()[0][3]
            error_line = str(sys.exc_info()[-1].tb_lineno)
            message = f"Error {self.bot_id}. Error on {method} Method. Error Line: {error_line}. Error {err}."
            print(message)
            sent_flag = False
            self.utils.log_to_file(message)
            self.utils.log_to_file(
                f"Error sending Email to: {email_to}\n\
                                            CC: {email_cc}\n\
                                            Subject: {subject}"
            )
            return sent_flag

    @staticmethod
    def send_encrypted_email(
        email_to,
        subject,
        body,
        html_body=False,
        html_images=[],
        attachments=[],
        email_cc="",
    ):
        """This method is used when needing to send an encrypted email through the Outlook Service.
        Args:
            email_to (str): Email of who will receive the email
            subject (str): Subject of the email
            body (str): Body of the email. It can be in plain text or the HTML Code. If HTML Code, you must set html_body to True
            html_body (bool, optional): Must set to True if body being passed is HTML Code. Defaults to False.
            html_images (list, optional): If html has images, you would need to pass the paths as a list. NOTE: Inside the HTML you would need to place 'src="cid:MyId1"', 'src="cid:MyId2"' and so on, depending on the images added. Defaults to [].
            attachments (list, optional): A list of the paths to the attachments added to the email. Defaults to [].
            email_cc (str, optional): If a different CC is used. Defaults to "BISClientCareEmail@experian.com".
        Return:
            Sent (bool) : Confirmation if email was sent or not.
        """

        def open_outlook():
            """This method is used to open outlook.
            Returns:
                opened (bool): Returns a bool if outlook was opened
            """
            opened = False
            try:
                os.startfile("outlook")
                time.sleep(3)
                opened = True
            except:
                print("Outlook didn't open successfully")
            return opened

        def is_outlook_opened():
            """This method is used to check if outlook is opened.
            Returns:
                opened (bool): Returns a bool if outlook is opened
            """
            opened = False
            for app in psutil.pids():
                process = psutil.Process(app)
                if process.name() == "OUTLOOK.EXE":
                    print("Outlook is opened")
                    opened = True
                    break
                else:
                    opened = False
            return opened

        sent = False
        try:
            # Check if outlook is open
            outlook_opened = is_outlook_opened()
            if not outlook_opened:
                counter = 1
                while counter < 6:
                    open_outlook()
                    outlook_opened = is_outlook_opened()
                    if outlook_opened:
                        counter = 9
                        break
                    else:
                        counter += 1
            if outlook_opened:
                outlook = win32com.client.Dispatch("outlook.application")
                mail = outlook.CreateItem(0)
                mail.To = email_to
                mail.CC = email_cc
                mail.Subject = "[Encrypt] " + str(subject)
                # if html body
                if html_body:
                    print(html_images)
                    if html_images:
                        index = 1
                        for item in html_images:
                            html_image = mail.Attachments.Add(item)
                            print(item)
                            image_id = f"MyId{index}"
                            print(image_id)
                            html_image.PropertyAccessor.SetProperty(
                                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                                image_id,
                            )
                            index += 1
                    mail.HTMLBody = body
                else:
                    mail.Body = body
                # To attach a file to the email (optional)
                if attachments:
                    for item in attachments:
                        mail.Attachments.Add(item)
                mail.Send()
                sent = True
                # Call your log file method here
            else:
                print("Unable to send email, unable to open outlook.")
                raise Exception("Unable to open Outlook to send email")
        except Exception as ex:
            error_line = sys.exc_info()[-1].tb_lineno
            print(
                "An error has occurred while sending the email, error details:"
                + str(ex)
                + ". Error line :"
                + str(error_line)
            )
            sent = False
        return sent

    @staticmethod
    def send_rpa_confirmation_email(
        self,
        message,
        error_line=None,
        method=None,
        attachment=None,
        subject=None,
        html=False,
        email_to=None,
        image=False,
        email_cc=None,
    ):
        """This method is used to send e-mails through Outlook"""
        try:
            print(f"Start {inspect.stack()[0][3]}")
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            if email_to is None:
                mail.To = self.failed_email_to  # Set TO address
                email_to = self.failed_email_to
            else:
                mail.To = email_to
            if email_cc is not None:
                mail.CC = email_cc  # Set CC address
            if subject is None:
                mail.Subject = (
                    f"Error {self.bot_id}. Method: {method}. Error Line: {error_line}"
                )
                subject = (
                    f"Error {self.bot_id}. Method: {method}. Error Line: {error_line}"
                )
            else:
                mail.Subject = subject  # Set Subject
            if not html:
                mail.Body = message  # Set body of email
            else:
                mail.HTMLBody = message
            if image:
                embed1 = mail.Attachments.Add(self.img_header_rpa)
                embed1.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "Header_RPA"
                )
                embed2 = mail.Attachments.Add(self.img_footer_rpa)
                embed2.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "Footer_RPA"
                )
            if attachment:
                for attach in attachment:
                    mail.Attachments.Add(
                        attach
                    )  # If attachment is going to be sent out
            mail.Send()
            print(f"End {inspect.stack()[0][3]}")
            # Call your log file method here
            # self.utils.log_to_file(
            #     "RPA confirmation Email successfully sent to SME."
            # )
        except Exception as err:
            err = str(err)
            method = inspect.stack()[0][3]
            error_line = str(sys.exc_info()[-1].tb_lineno)
            message = (
                f"Error on {method} Method. Error Line: {error_line}. Error {err}."
            )
            print(message)
