"""Script to write all selenium interaction with web pages"""

import inspect
import sys
import time

import autoit
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from .common_methods import fprint, handle_exeption


class SeleniumExtensions:
    """Class to write all selenium reusable methods"""

    def __init__(self, driver, log_file_path, debugging_mode=False, time_out=30):
        self.browser = driver
        self.log_file_path = log_file_path
        self.retry = 0
        self.main_window_handle = None
        self.debugging_mode = debugging_mode
        self.browser.maximize_window()
        self.continue_flag = False
        self.time_out = time_out
        self.wait = WebDriverWait(self.browser, self.time_out)

    def open_web_page(self, web_page_url, wait_element_xpath=""):
        """Method to open a web page and wait for an element to exist"""
        self.reset_variables()
        while self.retry < 3 and not self.continue_flag:
            try:
                self.browser.get(web_page_url)
                self.browser.maximize_window()
                if self.debugging_mode:
                    fprint("Web page open successfully, waiting for it to loads")
                if wait_element_xpath != "":
                    self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, wait_element_xpath)),
                        "Timeout Element not found by XPATH " + wait_element_xpath,
                    )
                if self.debugging_mode:
                    fprint("Web page loads successfully")
                self.continue_flag = True
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "open_web_page",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + wait_element_xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "open_web_page",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the "
                        + "HTML element due a Timeout error "
                        + wait_element_xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "open_web_page",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + wait_element_xpath
                    )

    def refresh_web_page(self, wait_element_xpath=""):
        """This method is used to refresh a site and wait
        until an element is visible"""
        self.reset_variables()
        while self.retry < 3 and not self.continue_flag:
            try:
                self.browser.refresh()
                if self.debugging_mode:
                    fprint("Web page refreshed succesfuly, waiting for it to load")
                if wait_element_xpath != "":
                    self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, wait_element_xpath)),
                        "Timeout Element not found by XPATH " + wait_element_xpath,
                    )
                if self.debugging_mode:
                    fprint("Web page loaded successfully")
                self.continue_flag = True
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "refresh_web_page",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + wait_element_xpath
                    ) from ex
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "refresh_web_page",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the "
                        + "HTML element due a Timeout error "
                        + wait_element_xpath
                    ) from ex
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "refresh_web_page",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + wait_element_xpath
                    ) from ex

    def clear_text_box(self, xpath):
        """Method to clear the text from a textbox"""
        self.reset_variables()
        if self.debugging_mode:
            fprint("clear_text_box", xpath)
        while self.retry < 3 and not self.continue_flag:
            try:
                self.browser.find_element("xpath", xpath).clear()
                text = self.get_text(xpath)
                if self.debugging_mode:
                    fprint("clear_text_box", text)
                if text != "":
                    if self.debugging_mode:
                        fprint("clear_text_box", text)
                    self.browser.find_element("xpath", xpath).send_keys(
                        Keys.CONTROL, "a"
                    )
                    self.browser.find_element("xpath", xpath).send_keys(Keys.BACKSPACE)
                text = self.get_text(xpath)
                if self.debugging_mode:
                    fprint("clear_text_box", text)
                if text != "":
                    text_length = len(text)
                    if self.debugging_mode:
                        fprint("clear_text_box_length", text, text_length)
                    for i in range(text_length + 1):
                        self.browser.find_element("xpath", xpath).send_keys(
                            Keys.BACKSPACE
                        )
                else:
                    self.continue_flag = True
                if self.debugging_mode and self.get_text(xpath) == "":
                    fprint("element", xpath, "was cleared successfully")
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "clear_text_box",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "clear_text_box",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML "
                        + "element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "clear_text_box",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def wait_for_element(self, xpath):
        """Method to wait for an element to exist"""
        self.reset_variables()
        if self.debugging_mode:
            fprint("wait for element", xpath)
        while self.retry < 3 and not self.continue_flag:
            try:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, xpath)),
                    "Timeout Element not found by XPATH " + xpath,
                )
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("element", xpath, "is ready")
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "wait_for_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "wait_for_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "wait_for_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def wait_for_text_in_element(self, xpath, text):
        """Method to wait for an element to exist"""
        self.reset_variables()
        if self.debugging_mode:
            fprint("wait for element", xpath)
        while self.retry < 3 and not self.continue_flag:
            try:
                self.wait.until(
                    EC.text_to_be_present_in_element((By.XPATH, xpath), text),
                    "Timeout Element not found by XPATH " + xpath,
                )
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("text", text, "in element", xpath, "is ready")
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "wait_for_text_in_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "wait_for_text_in_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "wait_for_text_in_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def wait_for_invisibility_of_element(self, xpath):
        """Method to wait for the invisibility of an element"""
        self.reset_variables()
        if self.debugging_mode:
            fprint("wait for invisibility of element", xpath)
        while self.retry < 3 and not self.continue_flag:
            try:
                self.wait.until(
                    EC.invisibility_of_element((By.XPATH, xpath)),
                    "Timeout Element not found by XPATH " + xpath,
                )
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("element", xpath, "is ready")
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "wait_for_invisibility_of_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "wait_for_invisibility_of_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "wait_for_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def wait_for_iframe_and_switch_to_it(self, xpath):
        """Method to waiit for an iframe and switch to it"""
        self.reset_variables()
        if self.debugging_mode:
            fprint("wait for iframe", xpath)
        while self.retry < 3 and not self.continue_flag:
            try:
                self.main_window_handle = self.browser.window_handles[0]
                self.wait.until(
                    EC.frame_to_be_available_and_switch_to_it((By.XPATH, xpath)),
                    "Timeout Element not found by XPATH " + xpath,
                )
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("iframe", xpath, "is ready")
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "wait_for_iframe_and_switch_to_it",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "wait_for_iframe_and_switch_to_it",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "wait_for_iframe_and_switch_to_it",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def switch_to_main_window(self):
        """Method to wait for an iframe and switch to it"""
        self.reset_variables()
        while self.retry < 3 and not self.continue_flag:
            try:
                self.browser.switch_to.window(self.main_window_handle)
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("switch to main window ready")
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "switch_to_main_window",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension was not able to switch to main window"
                    )

    def element_exist(self, xpath):
        """Method to check if an element is present on the current screen"""
        element_found = False
        self.reset_variables()
        while self.retry < 3 and not self.continue_flag:
            try:
                if len(self.browser.find_elements("xpath", xpath)) > 0:
                    element_found = True
                self.continue_flag = True
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "element_exist",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension was not able to find the element " + xpath
                    )
        return element_found

    def extract_table_column_attribute(
        self, xpath, attribute="text", remove_headers=False
    ):
        """Method to extract all the information of a HTML table"""
        column_information = []
        self.reset_variables()
        while self.retry < 3 and not self.continue_flag:
            try:
                column_elements = self.browser.find_elements("xpath", xpath)
                if self.debugging_mode:
                    print("Found " + str(len(column_elements)) + " to extract")
                for element in column_elements:
                    attribute_value = "[empty]"
                    try:
                        if attribute == "text":
                            attribute_value = element.text
                        else:
                            attribute_value = element.get_attribute(attribute)
                        print(attribute_value)
                    except Exception:
                        if self.debugging_mode:
                            print(
                                "Failed to extract text from element["
                                + str(column_elements.index(element))
                                + "], setting it to empty"
                            )
                    column_information.append(attribute_value)
                if remove_headers:
                    column_information = column_information[1:]
                self.continue_flag = True
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "extract_table_column_attribute",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "extract_table_column_attribute",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "extract_table_column_attribute",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )
        return column_information

    def click(self, xpath="", max_retry=3, web_element=None):
        """Method to click on an element"""
        exception_max_retry = max_retry - 1
        self.reset_variables()
        while self.retry < max_retry and not self.continue_flag:
            try:
                if self.debugging_mode:
                    fprint("Start clicking", xpath)
                error_while_clicking = False
                # Change it because when there in an element above the element that
                # you want to click then it will fail
                try:
                    if web_element is not None:
                        if self.debugging_mode:
                            print("Clicking web element")
                        web_element.click()
                    else:
                        print("Clicking web element using the xpath")
                        self.wait.until(
                            EC.presence_of_element_located((By.XPATH, xpath)),
                            "Timeout Element not found by XPATH " + xpath,
                        )
                        self.browser.find_element("xpath", xpath).click()
                except Exception as ex:
                    error_while_clicking = True
                    if self.debugging_mode:
                        print("error_while_clicking", error_while_clicking)
                if error_while_clicking:
                    element_to_be_clicked = None
                    if web_element is not None:
                        # action = ActionChains(self.browser)
                        # action.move_to_element(web_element).click().perform()
                        # print('Need to be fixed')# ActionChain it's no longer working after update chromedriver 90.xx
                        element_to_be_clicked = web_element
                    else:
                        element_to_be_clicked = self.browser.find_element(
                            "xpath", xpath
                        )
                        # action = ActionChains(self.browser)
                        # action.move_to_element(
                        #    self.browser.find_element_by_xpath(xpath)).click().perform()
                        # print('Need to be fixed')# ActionChain it's no longer working after update chromedriver 90.xx
                    if element_to_be_clicked is not None:
                        self.browser.execute_script(
                            "arguments[0].click();", element_to_be_clicked
                        )
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("End clicking", xpath)
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > exception_max_retry:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "click",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > exception_max_retry:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "click",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > exception_max_retry:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "click",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def double_click(self, xpath="", web_element=None):
        """Method to click on an element"""
        self.reset_variables()
        while self.retry < 3 and not self.continue_flag:
            try:
                if self.debugging_mode:
                    fprint("Start double clicking", xpath)
                error_while_clicking = False
                # Change it because when there in an element above the element that
                # you want to click then it will fail
                try:
                    if web_element is not None:
                        if self.debugging_mode:
                            print("Clicking web element")
                        web_element.double_click()
                    else:
                        print("Clicking web element using the xpath")
                        self.wait.until(
                            EC.presence_of_element_located((By.XPATH, xpath)),
                            "Timeout Element not found by XPATH " + xpath,
                        )
                        self.browser.find_element("xpath", xpath).double_click()
                except Exception as ex:
                    error_while_clicking = True
                    if self.debugging_mode:
                        print("error_while_clicking", error_while_clicking)
                if error_while_clicking:
                    if web_element is not None:
                        action = ActionChains(self.browser)
                        action.move_to_element(web_element).double_click().perform()
                    else:
                        action = ActionChains(self.browser)
                        action.move_to_element(
                            self.browser.find_element("xpath", xpath)
                        ).double_click().perform()
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("End double clicking", xpath)
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "doubleclick",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "doubleclick",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "doubleclick",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def set_text(self, xpath, input_text):
        """Method to set the text on a html textbox"""
        self.reset_variables()
        if self.debugging_mode:
            fprint("set text")
        while self.retry < 3 and not self.continue_flag:
            try:
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, xpath)),
                    "Timeout Element not found by XPATH " + xpath,
                )
                self.move_to_element(xpath)
                self.browser.find_element("xpath", xpath).send_keys(input_text)
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("text ready")
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "set_text",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "set_text",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "set_text",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def get_text_from_element(self, element):
        """Method to the the text of a HTML element"""
        self.reset_variables()
        if self.debugging_mode:
            fprint("get_text_from_element")
        while self.retry < 3 and not self.continue_flag:
            try:
                element_text = str(element.text).strip()
                if self.debugging_mode:
                    fprint("get_text_from_element", element_text)
                if element_text == "":
                    if self.debugging_mode:
                        fprint("get_text_from_element getting the value")
                    web_element_attribute = element.get_attribute("value")
                    if web_element_attribute is not None:
                        element_text = str(element.get_attribute("value")).strip()
                if self.debugging_mode:
                    fprint("get_text_from_element", element_text)
                self.continue_flag = True
                return element_text
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "get_text_from_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                    )

    def get_text(self, xpath):
        """Method to the the text of a HTML element"""
        self.reset_variables()
        if self.debugging_mode:
            fprint("get text")
        while self.retry < 3 and not self.continue_flag:
            try:
                self.move_to_element(xpath)
                element_text = str(
                    self.browser.find_element("xpath", xpath).text
                ).strip()
                if self.debugging_mode:
                    fprint("get text", element_text)
                if element_text == "":
                    if self.debugging_mode:
                        fprint("get text getting the value")
                    web_element_attribute = self.browser.find_element(
                        "xpath", xpath
                    ).get_attribute("value")
                    if web_element_attribute is not None:
                        element_text = str(
                            self.browser.find_element("xpath", xpath).get_attribute(
                                "value"
                            )
                        ).strip()
                if self.debugging_mode:
                    fprint("get text", element_text)
                self.continue_flag = True
                return element_text
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "get_text",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def get_elements_from_path(self, xpath):
        """Method to extract all the elements with a path"""
        self.reset_variables()
        lst_elements = []
        if self.debugging_mode:
            fprint("get elements from path")
        while self.retry < 3 and not self.continue_flag:
            try:
                lst_elements = self.browser.find_elements("xpath", xpath)
                if self.debugging_mode:
                    fprint("Number of elements extracted:", str(len(lst_elements)))
                self.continue_flag = True
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "get_text",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )
        return lst_elements

    def get_child_elements(self, web_element, child_xpath=""):
        """Method to get child element of a web element"""
        child_elements = []
        self.reset_variables()
        while self.retry < 3 and not self.continue_flag:
            try:
                if self.debugging_mode:
                    fprint(
                        "Start getting child elements with" + " xpath: " + child_xpath
                    )
                if web_element is not None and child_xpath != "":
                    child_elements = web_element.find_elements("xpath", child_xpath)
                else:
                    raise ValueError(
                        "Invalid parameters for method get_child_element()"
                    )
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("Getting child element " + " ready")
                if self.debugging_mode:
                    fprint("End getting child elements with" + " xpath: " + child_xpath)
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "get_child_elements",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + child_xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "get_child_elements",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + child_xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "get_child_elements",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + child_xpath
                    )
        return child_elements

    def get_attribute_from_element(self, attribute, web_element=None, xpath=""):
        """Method to get an specific attribute value
        of a web element"""
        attribute_value = ""
        self.reset_variables()
        while self.retry < 3 and not self.continue_flag:
            try:
                if web_element is not None:
                    if self.debugging_mode:
                        fprint(
                            "Start getting attribute '",
                            attribute,
                            "' from ",
                            str(web_element),
                            log_file_file_path=self.log_file_path,
                        )
                    attribute_value = str(web_element.get_attribute(attribute))
                elif xpath != "":
                    if self.debugging_mode:
                        fprint(
                            "Start getting attribute '",
                            attribute,
                            "' from xpath ",
                            xpath,
                            log_file_file_path=self.log_file_path,
                        )
                    attribute_value = str(
                        self.browser.find_element("xpath", xpath).get_attribute(
                            attribute
                        )
                    )
                else:
                    raise ValueError(
                        "Invalid parameters for method get_attribute_from_element()"
                    )
                self.continue_flag = True
                if self.debugging_mode:
                    fprint("Getting attribute " + attribute + " ready")
            except NoSuchElementException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "get_attribute_from_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize the HTML element "
                        + xpath
                    )
            except TimeoutException as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    fprint(
                        "get_attribute_from_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise Exception(
                        "Selenium Extension was not able to recognize "
                        + "the HTML element due a Timeout error "
                        + xpath
                    )
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "get_attribute_from_element",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )
        return attribute_value

    def scroll_into_panel(
        self,
        xpath,
        element_xpath,
        number_of_scrolls=None,
        scroll_until_end=False,
        scroll_up=False,
    ):
        """Method to scroll up or donw in a container"""
        self.reset_variables()
        try:
            self.wait_for_element(xpath)
            container = self.browser.find_element("xpath", xpath)
            if scroll_until_end:
                number_of_elements_previous = 0
                number_of_elements_current = len(
                    self.browser.find_elements_by_xpath(element_xpath)
                )
                print(number_of_elements_previous, number_of_elements_current)
                count = 0
                while (
                    number_of_elements_current != number_of_elements_previous
                    and count < 1000
                ):
                    number_of_elements_previous = number_of_elements_current
                    for i in range(count + 1 if count == 0 else count * 5):
                        if scroll_up:
                            container.send_keys(Keys.PAGE_UP)
                        else:
                            container.send_keys(Keys.PAGE_DOWN)
                    time.sleep(2)
                    number_of_elements_current = len(
                        self.browser.find_elements("xpath", element_xpath)
                    )
                    print(number_of_elements_previous, number_of_elements_current)
                    count += 1
            elif number_of_scrolls is not None:
                for i in range(number_of_scrolls):
                    if scroll_up:
                        container.send_keys(Keys.PAGE_UP)
                    else:
                        container.send_keys(Keys.PAGE_DOWN)
            else:
                raise ValueError("Invalid parameters for the method " + xpath)
        except NoSuchElementException as ex:
            self.retry += 1
            handle_exeption(ex)
            if self.retry > 2:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                # If retry is bigger than 3 it means the script WAS NOT
                # able to perform the action, raise exception to caller
                fprint(
                    "scroll_into_panel",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Selenium Extension was not able to recognize the HTML element "
                    + xpath
                )
        except TimeoutException as ex:
            self.retry += 1
            handle_exeption(ex)
            if self.retry > 2:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                # If retry is bigger than 3 it means the script WAS NOT
                # able to perform the action, raise exception to caller
                fprint(
                    "scroll_into_panel",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Selenium Extension was not able to recognize "
                    + "the HTML element due a Timeout error "
                    + xpath
                )
        except Exception as ex:
            self.retry += 1
            handle_exeption(ex)
            if self.retry > 2:
                # If retry is bigger than 3 it means the script WAS NOT
                # able to perform the action, raise exception to caller
                # class
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                fprint(
                    "scroll_into_panel",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise ValueError(
                    "Selenium Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + xpath
                )

    # INTERNAL METHODS, NOT NEEDED TO BE USED OUTSIDE THIS CLASS
    def move_to_element(self, xpath):
        """Method to scroll to the web element"""
        action = ActionChains(self.browser)
        action.move_to_element(self.browser.find_element("xpath", xpath))

    def reset_variables(self):
        """TODO @Richie we need to see if this is needed"""
        self.continue_flag = False
        self.retry = 0

    def get_selected_option(self, xpath):
        """Returns the text of a selected item in a
        Select html element"""
        self.reset_variables()
        element_text = ""
        if self.debugging_mode:
            fprint("get text")
        while self.retry < 3 and not self.continue_flag:
            try:
                select = Select(self.browser.find_element("xpath", xpath))
                element_text = select.first_selected_option.text
                self.continue_flag = True
                return element_text
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "get_selected_option",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def submit_alert_window(self):
        """Dimisses a browser alert window, this method.
        This method only clicks on Accept, it does not
        validate the message of the alert."""
        try:
            WebDriverWait(self.browser, 10).until(
                EC.alert_is_present(), "Time out waiting for alert window to appear"
            )
            alert = self.browser.switch_to.alert
            alert.accept()
        except Exception as ex:
            self.retry += 1
            handle_exeption(ex)
            if self.retry > 2:
                # If retry is bigger than 3 it means the script WAS NOT
                # able to perform the action, raise exception to caller
                # class
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                fprint(
                    "submit_alert_window",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise ValueError(
                    "Selenium Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details"
                )

    def print_page_to_pdf(self):
        """Prints a page as a PDF file. The image covers all the
        the page, therefore depending on the size of the same
        the pdf can contain multiple pages.
        Please use HC50_logic.py as a reference how to configure
        your browser to print a page.
        * Have the following configuration:
        pdf_settings = {
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
            "isCssBackgroundEnabled": True
        }
        prefs = {
            'download.default_directory' : [variable],
            'savefile.default_directory': [variable],
            'printing.print_preview_sticky_settings.appState': json.dumps(pdf_settings)
        }
        """
        try:
            self.browser.execute_script("window.print();")
        except Exception as ex:
            self.retry += 1
            handle_exeption(ex)
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            fprint(
                "Error on method print_page_to_pdf",
                "line: " + str(line),
                str(ex),
                log_file_file_path=self.log_file_path,
            )
            raise ValueError(
                "Selenium Extension an unexpected error happened, "
                + "please refer to the internal LOG file for more details"
            )

    def wait_for_new_window(self, expected_windows_number=2):
        """Method to wait for a new tab to open.
        Varaibles:
        * expected_windows_number: contains how many
        tabs a browser should have, for example if
        you only have one tab and click on an element
        that opens a new one, the variable should contain 2"""
        try:
            self.wait.until(EC.number_of_windows_to_be(expected_windows_number))
        except Exception as ex:
            self.retry += 1
            handle_exeption(ex)
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            fprint(
                "Error switching to a different window",
                "line: " + str(line),
                str(ex),
                log_file_file_path=self.log_file_path,
            )
            raise ValueError(
                "Selenium Extension an unexpected error happened on method, "
                + "switch_to_other_window"
                + "please refer to the internal LOG file for more details"
            )

    def switch_to_other_window(self, index_number):
        """Small method to switch the self.browser variable
        to a different window (tab). Ths index of pages (tabs)
        starts with 0, in other words the first original tab of
        Chrome has the index_number 0"""
        try:
            self.browser.switch_to.window(self.browser.window_handles[index_number])
        except Exception as ex:
            self.retry += 1
            handle_exeption(ex)
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            fprint(
                "Error switching to a different window",
                "line: " + str(line),
                str(ex),
                log_file_file_path=self.log_file_path,
            )
            raise ValueError(
                "Selenium Extension an unexpected error happened on method, "
                + "switch_to_other_window"
                + "please refer to the internal LOG file for more details"
            )

    def close_window(self):
        """Small method to close a
        Chrome's tab. Ths index of pages (tabs)
        starts with 0, in other words the first original tab of
        Chrome has the index_number 0"""
        try:
            self.browser.close()
        except Exception as ex:
            self.retry += 1
            handle_exeption(ex)
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            fprint(
                "Error closing window",
                "line: " + str(line),
                str(ex),
                log_file_file_path=self.log_file_path,
            )
            raise ValueError(
                "Selenium Extension an unexpected error happened on method, "
                + "close_window"
                + "please refer to the internal LOG file for more details"
            )

    def is_checkbox_selected(self, xpath):
        """Returns a flag indicating if a Checkbox
        element is checked.
        * True: checked
        * False: unchecked"""
        self.reset_variables()
        if self.debugging_mode:
            fprint("get checkbox flag")
        while self.retry < 3 and not self.continue_flag:
            try:
                self.move_to_element(xpath)
                check_flag = self.browser.find_element("xpath", xpath).is_selected()
                if self.debugging_mode:
                    fprint("get select status", check_flag)
                return check_flag
            except Exception as ex:
                self.retry += 1
                handle_exeption(ex)
                if self.retry > 2:
                    # If retry is bigger than 3 it means the script WAS NOT
                    # able to perform the action, raise exception to caller
                    # class
                    trace_back = sys.exc_info()[2]
                    line = trace_back.tb_lineno
                    fprint(
                        "is_checkbox_selected",
                        "line: " + str(line),
                        str(ex),
                        log_file_file_path=self.log_file_path,
                    )
                    raise ValueError(
                        "Selenium Extension an unexpected error happened, "
                        + "please refer to the internal LOG file for more details "
                        + xpath
                    )

    def remove_attribute_from_element(self, id, attribute):
        """Method to run a JavaScript which removes
        the disabled a single attribute in an element.
        Important: it does not support more than one attribute
        at the same time"""
        try:
            self.browser.execute_script(
                "document.getElementById('"
                + id
                + "').removeAttribute('"
                + attribute
                + "')"
            )
        except Exception as ex:
            handle_exeption(ex)
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            fprint(
                "is_checkbox_selected",
                "line: " + str(line),
                str(ex),
                log_file_file_path=self.log_file_path,
            )
            raise ValueError(
                "Selenium Extension an unexpected error happened, "
                + "please refer to the internal LOG file for more details "
                + "ID: "
                + str(id)
                + " - Attribute: "
                + str(attribute)
            )

    def popup_login(self, user, password, success_xpath):
        """Method to log into pop up"""
        # Authentication is VM email and password
        success = True
        try:
            fprint(f"Start {inspect.stack()[0][3]}", log_file_path=self.log_file_path)
            # Autoit Commands to fill in URL authentication
            # Make sure you give blank since the cursor is at userid
            autoit.win_wait_active("", 30)
            autoit.send(user)
            autoit.send("{TAB}")
            time.sleep(0.5)
            autoit.send(password)
            autoit.send("{ENTER}")
            time.sleep(0.5)

            # Verify xpath exists due to duccessful log in
            self.wait_for_element(success_xpath)

        except Exception as err:
            err = str(err)
            method = inspect.stack()[0][3]
            error_line = str(sys.exc_info()[-1].tb_lineno)
            message = (
                f"Error on {method} Method. Error Line: {error_line}. Error {err}."
            )
            fprint(message, log_file_file_path=self.log_file_path)
            success = False
        return success
