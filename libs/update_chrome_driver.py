"""Method to update Chrome Driver"""

import io
import re
import sys
import time
import zipfile

from requests import Session
from win32com.client import Dispatch

from .common_methods import kill_process, log_file
from .email_helper_v2 import send_email


class UpdateChromeDriver:
    """Class to write all Update Chrome Driver reusable methods"""

    def __init__(
        self,
        destination,
        log_file_path,
        bot_id,
        loc_chrome_exe,
        email_from="",
        email_to="",
    ):
        self.url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
        self.destination = destination
        self.log_file_path = log_file_path
        self.bot_id = bot_id
        self.email_from = email_from
        self.email_to = email_to
        self.loc_chrome_exe = loc_chrome_exe
        print(self.loc_chrome_exe)
        self.session = Session()

    def download(self):
        """function to download chrome driver"""
        download_success = False
        try:
            success = kill_process("chromedriver.exe")
            if success:
                # get the current installed chrome version number
                current_version = self.current_chrome_version()
                # get the latest chrome driver version number
                wshell = Dispatch("wscript.shell")
                wshell.run(self.url)
                time.sleep(5)
                response = self.session.get(self.url)
                time.sleep(5)
                chrome_driver_version = response.text
                response.close()
                # wshell.SendKeys("^w")
                chrome_driver_version_compare = chrome_driver_version.split(".")[0]
                print(chrome_driver_version_compare)
                # chrome_driver_version_compare = 115
                pattern = re.compile(r"^[0-9\.]+$")
                # if chrome version less than chrome driver, we lower chrome driver version to chrome version
                if int(current_version) < int(chrome_driver_version_compare):
                    wshell = Dispatch("wscript.shell")
                    wshell.run(f"{self.url}_{current_version}")
                    response = self.session.get(f"{self.url}_{current_version}")
                    if bool(pattern.match(response.text)):
                        download_url = f"http://chromedriver.storage.googleapis.com/{response.text}/chromedriver_win32.zip"
                    else:
                        raise Exception(
                            f"Version Number format incorrect {response.text}"
                        )
                    response.close()
                    wshell.SendKeys("^w")
                # if chrome version higher or equal than chrome driver, we download latest chrome driver version to chrome version
                else:
                    if bool(pattern.match(chrome_driver_version)):
                        download_url = f"http://chromedriver.storage.googleapis.com/{chrome_driver_version}/chromedriver_win32.zip"
                    else:
                        raise Exception(
                            f"Version Number format incorrect {chrome_driver_version}"
                        )
                # download the zip file using the url built above
                latest_driver_zip = self.session.get(download_url)
                pattern = re.compile(b"chromedriver\\.exe")
                latest_driver_zip.close()
                # extract the zip file
                if latest_driver_zip.status_code == 200:
                    if bool(pattern.search(latest_driver_zip.content)):
                        # if bool(re.search(regex, chromefile)):
                        with zipfile.ZipFile(
                            io.BytesIO(latest_driver_zip.content), mode="r"
                        ) as zip_ref:
                            zip_ref.extractall(self.destination)
                        print("Chrome Driver Update Successful")
                        log_file(
                            "Chrome Driver Update Successful",
                            self.log_file_path,
                            self.bot_id,
                            self.email_from,
                            self.email_to,
                        )
                        download_success = True
                    else:
                        raise Exception(
                            "Downloaded Chromedriver file didn't pass regex test"
                        )
                else:
                    raise Exception(
                        f"Incorrect API Reponse code: {latest_driver_zip.status_code}"
                    )
            else:
                raise Exception(
                    "Unable to kill running chrome or chromedriver application."
                )
        except Exception as error:
            subject = "Error on ChromeDriver Download"
            trace_back = sys.exc_info()[-1]
            line = trace_back.tb_lineno
            body = f"Error on Chrome Driver Download. Bot ID: {self.bot_id}. Line {line}. Error: {error}."
            log_file(
                "Chrome Driver Update Successful",
                self.log_file_path,
                self.bot_id,
                self.email_from,
                self.email_to,
            )
            send_email(self.email_from, self.email_to, subject, body)
        return download_success

    def current_chrome_version(self):
        """Method to find the chrome.exe version"""
        try:
            parser = Dispatch("Scripting.FileSystemObject")
            version = parser.GetFileVersion(self.loc_chrome_exe)
            version = version.split(".")[0]
            print(version)
        except Exception as error:
            subject = "Error on kill_process"
            trace_back = sys.exc_info()[-1]
            line = trace_back.tb_lineno
            body = f"Error on Chrome Driver current_chrome_version. Bot ID: {self.bot_id}. Line {line}. Error: {error}."
            log_file(
                "Chrome Driver Update Successful",
                self.log_file_path,
                self.bot_id,
                self.email_from,
                self.email_to,
            )
            version = None
            send_email(self.email_from, self.email_to, subject, body)
        return version


# update_chrome_driver implements Class UpdateChromeDriver
def update_chrome_driver(
    destination, log_file_path, bot_id, loc_chrome_exe, email_from="", email_to=""
):
    """Update chrome driver to latest version in page"""
    update_chrome_driver = UpdateChromeDriver(
        destination, log_file_path, bot_id, loc_chrome_exe, email_from, email_to
    )
    if not update_chrome_driver.download():
        raise Exception(
            "Unable to download the latest chrome driver. For more information "
            + "please see the system logs"
        )
