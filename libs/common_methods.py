"""The script is to store reusable methods"""

import inspect
import os
import sys
import time
from datetime import datetime, timedelta
from string import Template

import openpyxl
import psutil
import pyautogui
import win32com.client
from mss import mss
from smartAutomationToolBox.email_lib.email_helper import (send_email,
                                                           send_outlook_email)

"""Class to store reusable method"""
def log_file(message, log_file_path, bot_id):
    """record log files"""
    try:
        now = str(datetime.today().strftime("(%m/%d/%Y %H:%M:%S) "))
        with open(log_file_path, "a", encoding="utf8") as foxtrot:
            foxtrot.write(now + message + "\n")
    except Exception as error:
        print("Unable to log to file")
        subject = "Error on log to file process"
        trace_back = sys.exc_info()[-1]
        line = trace_back.tb_lineno
        body = f"Error on log_file method. Bot ID: {bot_id}. Line {line}. Error: {error}."
        print(body)

def error(message, error_line, method, log_file_path, error_file_path, bot_id, splunk_calls, \
    email_to, subject):
    """This method is used for when a method errors out"""
    errorflag = True
    try:
        log_file(f"Start {inspect.stack()[0][3]}", log_file_path, bot_id)
        print(f"Start {inspect.stack()[0][3]}")
        log_file(message, log_file_path, bot_id)
        now = str(datetime.today().strftime("%m-%d-%Y - %H-%M-%S"))
        error_file_path = Template(error_file_path).substitute(date=now)
        pyautogui.screenshot(error_file_path)
        attach = [error_file_path, log_file_path]
        send_outlook_email(message=message, log_file_path=log_file_path, bot_id=bot_id, \
            splunk_calls=splunk_calls, email_to=email_to, subject=subject, attachment=attach)
        time.sleep(0.5)
        if os.path.isfile(error_file_path):
            os.remove(error_file_path)
        time.sleep(3)
    except Exception as err:
        err = str(err)
        method = inspect.stack()[0][3]
        error_line = str(sys.exc_info()[-1].tb_lineno)
        message = f"Error on {method} Method. Error Line: {error_line}. Error {err}."
        send_outlook_email(message, log_file_path=log_file_path, bot_id=bot_id, \
            splunk_calls=splunk_calls, email_to=email_to, subject=subject, attachment=attach)
    return errorflag

def fprint(*args, log_file_file_path=None, session_id=''):
    """Method to print a desire message in the console
        and save it in the bot log file"""
    message = ''
    print("before log")
    for arg in args:
        message += str(arg) + " "
    message.strip()
    if log_file_file_path is not None:
        now = str(datetime.today().strftime("(%m/%d/%Y %H:%M:%S) "))
        with open(log_file_file_path, 'a', 1, encoding="utf8") as file:
            file.write(now + "------- " + session_id +
                " ------- " + message + " ------- " +\
                session_id + " -------\n")
            file.close()
            print("file updated")
    print("-------", message, "-------")

def handle_exeption(exception, project_folder_path='',\
    email_from='', email_to=[], subject='', body='',\
    log_file_file_path=None, attachment=None):
    """Reusable method to handle exception"""
    # Add better exception handling
    # Direct selenium error should be logged directly,
    # \ but the method on this class should also trigger an exception
    # so the main class can detect an error.
    screenshot_file_path = None
    if project_folder_path != '':
        screenshot_handler = mss()
        screenshot_handler.shot()
        now = datetime.now()
        screenshot_file_path = os.path.join(project_folder_path, 'Error '\
            + str(now.month)\
                + "." + str(now.day)\
                    + "." + str(now.year) + "."\
                        + str(now.hour)\
                            + "." + str(now.minute) + "."\
                                + str(now.second)\
                                    + "." + str(now.microsecond)\
                                        + '.png')
        screenshot_files = screenshot_handler.save(output=screenshot_file_path)
        screenshot_handler.close()
        #Must be in the code if you remove it the library will not save the file
        lst_attachments = []
        for screenshot_file_path in screenshot_files:
            if screenshot_file_path not in lst_attachments:
                lst_attachments.append(screenshot_file_path)
        body = body.replace('[exception_message]', str(exception))
        if attachment is not None:
            lst_attachments.append(attachment)
        email_to = ", ".join(email_to)
        send_email(email_from, email_to, subject, body, lst_attachments, True)
    fprint(exception, log_file_file_path=log_file_file_path)

def pause_execution():
    """Method to pause python execution"""
    try:
        pause = input('Press any key to continue the execution')
        print(pause)
    except Exception as ex:
        print('Error while pausing execution', ex)

def calculate_fiscal_year():
    """Method to calculate current FY"""
    fiscal_year = int(datetime.strftime(datetime.now(), "%y"))
    if int(datetime.strftime(datetime.now(), "%m")) >= 4:
        fiscal_year += 1
    return fiscal_year

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
        print('Error while adding business days', ex)
        raise("Exception adding business days: " + str(Exception))

    return date

def run_macro(macro_workbook, macro_name, macro_error_file, log_file_path, bot_id, \
    project_folder_path, email_from="", email_to="", macro_variables=[]):
    """run macro, take up to 256 variables"""
    try:
        if os.path.isfile(macro_error_file):
            os.remove(macro_error_file)
        log_file(f"Start Run Macro Method {macro_name}", log_file_path, bot_id)
        print(f"Start Run Macro Method {macro_name}")############################
        excl = win32com.client.Dispatch("Excel.Application")
        excl.Workbooks.Open(Filename=macro_workbook, ReadOnly=1)
        excl.Application.Run(macro_name, macro_variables) #########################
        print("Finish Run Macro")
        excl.Workbooks(1).Close(SaveChanges=0)
        excl.Application.Quit()
        excl = 0
        del excl
        log_file("Run Macro Method complete", log_file_path, bot_id)
        if os.path.isfile(macro_error_file):
            raise Exception(f"Error running Macro Method {macro_name},\
                please refer to log file or {macro_error_file} for more information.")
    except Exception as err:
        err = str(err)
        method = "Run Macro"
        error_line = str(sys.exc_info()[-1].tb_lineno)
        message = f"Error on {method} Method. Macro: {macro_name}.\
             Error Line: {error_line}. Error {err}."
        handle_exeption(err, project_folder_path, email_from, email_to, 'Error running macro',\
             message, log_file_path)

def kill_process(process_name, log_file_path, bot_id, email_to, splunk_calls):
    """Method to  find and kill running applications"""
    
    try:
        for proc in psutil.process_iter():
            try:
                # Check if process name contains the given name string.
                if process_name.lower() in proc.name().lower():
                    print(f'Killing {proc.name()}')
                    proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        return True
    except Exception as error:
        subject = "Error on kill_process"
        trace_back = sys.exc_info()[-1]
        line = trace_back.tb_lineno
        body = f"Error on Chrome Driver kill_process. Bot ID: {bot_id}. Line {line}. Error: {error}."
        log_file(body, log_file_path, bot_id)
        send_outlook_email(body, log_file_path=log_file_path, bot_id=bot_id, \
            splunk_calls=splunk_calls, email_to=email_to, subject=subject)
        return False

def get_excel_sheet_names(file_path):
    """Use this method to get the names of the sheets on an Excel file"""
    wb = openpyxl.load_workbook(file_path)
    return wb.sheetnames


