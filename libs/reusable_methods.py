"""The script is to store reusable methods"""

import glob
import math
import os
import sys
import time
from datetime import datetime, timedelta

import pandas as pd
from mss import mss
from styleframe import StyleFrame

from .email_helper_v2 import EmailHelper


class ReusableMethods:
    """Class to store reusable method"""

    @staticmethod
    def wait_for_file_to_be_downloaded(
        downloads_folder_path,
        log_file_file_path=None,
        download_file_extension="",
        wait_in_seconds=60,
    ):
        """Method to wait for a file to be downloaded"""
        number_of_iterations = math.ceil(wait_in_seconds / 2)
        last_downloaded_file_name_recent = ""
        try:
            ReusableMethods.fprint(
                "Start downloading file", log_file_file_path=log_file_file_path
            )
            list_of_files = []
            if download_file_extension == "":
                list_of_files = glob.glob(os.path.join(downloads_folder_path, "*.*"))
            else:
                list_of_files = glob.glob(
                    os.path.join(downloads_folder_path, download_file_extension)
                )
            last_downloaded_file_name = (
                "" if not list_of_files else max(list_of_files, key=os.path.getctime)
            )
            last_downloaded_file_name_recent = (
                "" if not list_of_files else max(list_of_files, key=os.path.getctime)
            )
            ReusableMethods.fprint(
                last_downloaded_file_name, log_file_file_path=log_file_file_path
            )
            ReusableMethods.fprint(
                last_downloaded_file_name_recent, log_file_file_path=log_file_file_path
            )
            count = 0
            while (
                last_downloaded_file_name == last_downloaded_file_name_recent
                and count < number_of_iterations
            ):
                time.sleep(2)
                if download_file_extension == "":
                    list_of_files = glob.glob(
                        os.path.join(downloads_folder_path, "*.*")
                    )
                else:
                    list_of_files = glob.glob(
                        os.path.join(downloads_folder_path, download_file_extension)
                    )
                print(list_of_files, download_file_extension, downloads_folder_path)
                last_downloaded_file_name_recent = (
                    ""
                    if not list_of_files
                    else max(list_of_files, key=os.path.getctime)
                )
                ReusableMethods.fprint(
                    last_downloaded_file_name_recent,
                    log_file_file_path=log_file_file_path,
                )
                count += 1
            if last_downloaded_file_name == last_downloaded_file_name_recent:
                last_downloaded_file_name_recent = None
                ReusableMethods.fprint(
                    "Time out exceeded while waiting for a file to be dowloaded",
                    log_file_file_path=log_file_file_path,
                )
                return last_downloaded_file_name_recent
                # raise ValueError('Time out exceeded while waiting for a file to be dowloaded')
            time.sleep(2)
            ReusableMethods.fprint(
                "End downloading file", log_file_file_path=log_file_file_path
            )
        except Exception as ex:
            ReusableMethods.fprint(
                "Exception in wait_for_file_to_be_downloaded",
                log_file_file_path=log_file_file_path,
            )
            ReusableMethods.fprint(ex, log_file_file_path=log_file_file_path)
        return last_downloaded_file_name_recent

    @staticmethod
    def wait_for_file_to_exist(
        file_path, full_file_name, log_file_file_path=None, wait_in_seconds=60
    ):
        """Method to wait for a file to exist, the file name and extension is required"""
        number_of_iterations = math.ceil(wait_in_seconds / 2)
        count = 0
        try:
            while count < number_of_iterations:
                if os.path.isfile(os.path.join(file_path, full_file_name)):
                    # FILE FOUND, EXIT LOOP
                    return True
                else:
                    count += 1
                    time.sleep(2)
        except Exception as ex:
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            ReusableMethods.fprint(
                "wait_for_file_to_exist",
                "line: " + str(line),
                str(ex),
                log_file_file_path=log_file_file_path,
            )
        return False

    @staticmethod
    def create_excel_report(
        report_file_path,
        report_sheet_name,
        headers_names,
        data,
        columns_width_dict=None,
    ):
        """Method to create an excel report with desire format"""
        report_successfully_created = True
        # ----------------------- columns_width_dict example -----------------------
        # This object must be constructed where you are calling this method
        # and must be passed as parameter to this method, please take in
        # count that the column width dict must have the columns names
        # that you want to format
        #   columns_width_dict = {
        #        ('Project Account', 'Project Revenue', 'Project Type',\
        #        'Project Start Date', 'All BAs', 'Assigned To'): 20,
        #        ('First Filter - Max Project Count', 'Second Filter - Max Conflicting Projecs',\
        #        'Third Filter - Min project Count', 'Fourth Filter - Last Assigment Date'): 75
        #    }
        # ----------------------- columns_width_dict example -----------------------
        try:
            report_data_frame = pd.DataFrame(data, columns=headers_names)
            excel_writer = StyleFrame.ExcelWriter(report_file_path)
            styled_frame = StyleFrame(report_data_frame)
            if columns_width_dict is not None:
                styled_frame.set_column_width_dict(col_width_dict=columns_width_dict)
            styled_frame.to_excel(
                excel_writer=excel_writer, sheet_name=report_sheet_name, index=False
            )
            excel_writer.save()
        except Exception as ex:
            print(ex)
            report_successfully_created = False
        return report_successfully_created

    @staticmethod
    def fprint(*args, log_file_file_path=None, session_id=""):
        """Method to print a desire message in the console
        and save it in the bot log file
        """
        message = ""
        print("before log")
        for arg in args:
            message += str(arg) + " "
        message.strip()
        if log_file_file_path is not None:
            with open(log_file_file_path, "a", 1) as file:
                file.write(
                    "------- "
                    + session_id
                    + " ------- "
                    + message
                    + " ------- "
                    + session_id
                    + " -------\n"
                )
                file.close()
                print("file updated")
        print("-------", message, "-------")

    @staticmethod
    def handle_exeption(
        exception,
        project_folder_path="",
        email_from="",
        email_to=[],
        subject="",
        body="",
        log_file_file_path=None,
        attachment=None,
    ):
        """Reusable method to handle exception"""
        # Add better exception handling
        # Direct selenium error should be logged directly,
        # \ but the method on this class should also trigger an exception
        # so the main class can detect an error.
        screenshot_file_path = None
        if project_folder_path != "":
            screenshot_handler = mss()
            screenshot_handler.shot()
            now = datetime.now()
            screenshot_file_path = os.path.join(
                project_folder_path,
                "Error "
                + str(now.month)
                + "."
                + str(now.day)
                + "."
                + str(now.year)
                + "."
                + str(now.hour)
                + "."
                + str(now.minute)
                + "."
                + str(now.second)
                + "."
                + str(now.microsecond)
                + ".png",
            )
            screenshot_files = screenshot_handler.save(output=screenshot_file_path)
            screenshot_handler.close()
            # Must be in the code if you remove it the library will not save the file
            lst_attachments = []
            for screenshot_file_path in screenshot_files:
                if screenshot_file_path not in lst_attachments:
                    lst_attachments.append(screenshot_file_path)
            body = body.replace("[exception_message]", str(exception))
            if attachment is not None:
                lst_attachments.append(attachment)
            EmailHelper.send_email(
                email_from, email_to, subject, body, lst_attachments, True
            )
        ReusableMethods.fprint(exception, log_file_file_path=log_file_file_path)

    @staticmethod
    def pause_execution():
        """Method to pause python execution"""
        try:
            pause = input("Press any key to continue the execution")
            print(pause)
        except Exception as ex:
            print("Error while pausing execution", ex)

    @staticmethod
    def add_business_days(date, days_to_add):
        """Method to add business days to a date"""
        try:
            weekend_days = 0
            for i in range(days_to_add):
                if days_to_add < 0:
                    date = date - timedelta(days=1)
                else:
                    date = date + timedelta(days=1)
                if date.weekday() == 5 or date.weekday() == 6:
                    weekend_days += 1
            if weekend_days > 0:
                if days_to_add < 0:
                    date = date - timedelta(days=weekend_days)
                else:
                    date = date + timedelta(days=weekend_days)
            if days_to_add < 0:
                if date.weekday() == 5:
                    date = date - timedelta(days=1)
                if date.weekday() == 6:
                    date = date - timedelta(days=2)
            else:
                if date.weekday() == 5:
                    date = date + timedelta(days=2)
                if date.weekday() == 6:
                    date = date + timedelta(days=1)
        except Exception as ex:
            print("Error while adding business days", ex)
        return date

    @staticmethod
    def log_to_file(log_file_file_path, text, trucate_before_write=False):
        """Method to write into a file with out any formatting"""
        with open(log_file_file_path, "a", 1) as file:
            if trucate_before_write:
                file.truncate()
            now = str(datetime.today().strftime("(%m/%d/%Y %H:%M:%S) "))
            file.write(now + text + "\n")
            file.close()
