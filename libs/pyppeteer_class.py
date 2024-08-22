import asyncio
import sys
import warnings

import pandas as pd
import pyppeteer

from .reusable_methods import ReusableMethods


class pyppeteer_extension:
    """Class to write all pyppeteer reusable methods
    Parameters:
    driver = Executable path for Chrome.
    data = Loaded json data
    """

    def __init__(self, browser: str, data: str, time_out=120000):

        warnings.filterwarnings("ignore")
        self.log_file_path = data
        self.retry = 0
        self.continue_flag = False
        self.browser = browser
        self.time_out = time_out
        self.original_main_page = None
        self.current_page = None
        self.set_current_page_pp()

    def set_current_page_pp(self):
        """Method to set the current page"""
        response = asyncio.get_event_loop().run_until_complete(
            self.set_current_page_ppp()
        )
        return response

    async def set_current_page_ppp(self):
        """
        Sets the current page to be used for web scraping.

        This method retrieves the pages from the browser and sets the current page to be used for web scraping.
        If the original main window is not set, it sets the original main window to be the first page.
        If the current page is not set, it sets the current page to be the original main window.
        Finally, it sets the continue flag to True and returns it.

        Returns:
            bool: The value of the continue flag indicating if the operation was successful.
        """
        self.continue_flag = False
        page = await self.browser.pages()
        if self.original_main_page is None:
            self.original_main_page = page[0]
        if self.current_page is None:
            self.current_page = self.original_main_page
        self.continue_flag = True
        return self.continue_flag

    def set_downloads_folder(self, folder_path):
        """Use to setup the downloads folder

        Args:
            folder_path (str): The path of the downloads folder

        Returns:
             response(bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.set_downloads_folder_pp(folder_path)
        )
        return response

    async def set_downloads_folder_pp(self, folder_path):
        """*** For Class use only - Refer to set_downloads_folder method instead ***"""
        attempt = self.retry
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                page = self.current_page
                await page._client.send(
                    "Page.setDownloadBehavior",
                    {"behavior": "allow", "downloadPath": folder_path},
                )
                ReusableMethods.fprint("Web page opened successfully")
                self.continue_flag = True
                return self.continue_flag
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "open_web_page",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + "Setting downloads path"
                )
        return self.continue_flag

    def open_web_page(self, web_page_url, wait_element_xpath=""):
        """Use to open a website in chrome.

        Parameters:
            a (str) web_page_url: Provide the website URL you want to open.
            b (str) wait_element_xpath(str): *Optional - Provide an xpath for the function to fisnish until that element is loaded.

        Returns:
            a (bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.open_web_page_pp(web_page_url, wait_element_xpath)
        )
        return response

    async def open_web_page_pp(self, web_page_url, wait_element_xpath=""):
        """*** For Class use only - Refer to open_web_page method instead ***"""
        attempt = self.retry
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                page = self.current_page

                await page.goto(web_page_url, timeout=self.time_out)
                await page.setViewport({"width": 1920, "height": 1040})
                if wait_element_xpath != "":
                    attempt = 0
                    while attempt < 3:
                        try:
                            await page.waitForXPath(
                                wait_element_xpath, timeout=self.time_out
                            )
                            ReusableMethods.fprint(
                                "Web page opened successfully, xpath load completed"
                            )
                            self.continue_flag = True
                            return self.continue_flag
                        except pyppeteer.errors.NetworkError:
                            attempt = attempt + 1
                            ReusableMethods.fprint(
                                f"Attempt {attempt}: Timed out while waiting for element."
                            )
                else:
                    ReusableMethods.fprint("Web page opened successfully")
                    self.continue_flag = True
                    return self.continue_flag
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "open_web_page",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + wait_element_xpath
                )
        return self.continue_flag

    def refresh_web_page(self, wait_element_xpath=""):
        """Use to reload a website in chrome.

        Parameters:
            a (str) wait_element_xpath(str): *Optional - Provide an xpath for the function to fisnish until that element is loaded.

        Returns:
            a (bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.refresh_web_page_pp(wait_element_xpath)
        )
        return response

    async def refresh_web_page_pp(self, wait_element_xpath=""):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        self.continue_flag = False
        attempt = self.retry
        while attempt < 3 and not self.continue_flag:
            try:
                page = self.current_page

                await page.reload()
                if wait_element_xpath != "":
                    attempt = 0
                    while attempt < 3:
                        try:
                            await page.waitForXPath(
                                wait_element_xpath, timeout=self.time_out
                            )
                            ReusableMethods.fprint(
                                "Web page opened successfully, xpath load completed"
                            )
                            self.continue_flag = True
                            return self.continue_flag
                        except pyppeteer.errors.NetworkError:
                            attempt = attempt + 1
                            ReusableMethods.fprint(
                                f"Attempt {attempt}: Timed out while waiting for element."
                            )
                else:
                    ReusableMethods.fprint("Web page reloaded successfully")
                    self.continue_flag = True
                    return self.continue_flag
            except Exception as ex:
                attempt += 1
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "refresh_web_page",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + wait_element_xpath
                )
        return self.continue_flag

    def clear_text_box(self, xpath, iframe=False):
        """Use to reload a website in chrome.

        Parameters:
            a (str) xpath(str): Provide an xpath for the field to clear.

        Returns:
            a (bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.clear_text_box_pp(xpath, iframe)
        )
        return response

    async def clear_text_box_pp(self, xpath, iframe):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        attempt = self.retry
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                if not iframe:
                    page = self.current_page

                    textbox = await page.waitForXPath(xpath)
                    await textbox.click()
                    await page.keyboard.down("Control")
                    await page.keyboard.press("KeyA")
                    await page.keyboard.up("Control")
                    await page.keyboard.press("Backspace")
                    ReusableMethods.fprint(f"Text box cleared atXPath '{xpath}'.")
                    self.continue_flag = True
                    return self.continue_flag
                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    textbox = await content.waitForXPath(xpath)
                    await textbox.click()
                    await content.keyboard.down("Control")
                    await content.keyboard.press("KeyA")
                    await content.keyboard.up("Control")
                    await content.keyboard.press("Backspace")
                    ReusableMethods.fprint(f"Text box cleared atXPath '{xpath}'.")
                    self.continue_flag = True
                    return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "refresh_web_page",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + xpath
                )
        ReusableMethods.fprint(f"Unable to clear text box at XPath '{xpath}'.")
        return self.continue_flag

    def wait_for_element(self, xpath, time=60, iframe=False):
        """Use to reload a website in chrome.

        Parameters:
            a (str) wait_element_xpath(str): *Optional - Provide an xpath for the function to fisnish until that element is loaded.

        Returns:
            a (bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.wait_for_element_pp(xpath, time, iframe)
        )
        print("Response: " + str(response))
        return response

    async def wait_for_element_pp(self, xpath, time, iframe):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        attempt = self.retry
        time = time * 1000
        element = ""
        self.continue_flag = False
        while attempt <= 3:
            try:
                if not iframe:
                    page = self.current_page

                    element = await page.waitForXPath(xpath, timeout=time)
                    ReusableMethods.fprint(f"Element exists on page")
                    self.continue_flag = True
                    return self.continue_flag
                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=time)
                    content = await frames.contentFrame()
                    click = await content.waitForXPath(xpath, timeout=time)
                    ReusableMethods.fprint(f"Element exists on page")
                    self.continue_flag = True
                    return self.continue_flag
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                continue
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                continue
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "refresh_web_page",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + xpath
                )
        return self.continue_flag

    def set_text(self, xpath, input_text, iframe=False):
        """Use to reload a website in chrome.

        Parameters:
            a (str) wait_element_xpath(str): *Optional - Provide an xpath for the function to fisnish until that element is loaded.

        Returns:
            a (bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.set_text_pp(xpath, input_text, iframe)
        )
        return response

    async def set_text_pp(self, xpath, input_text, iframe):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        attempt = self.retry
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                if not iframe:
                    page = self.current_page

                    element = await page.xpath(xpath)
                    if element:
                        await element[0].type(input_text)
                        ReusableMethods.fprint(
                            f"Text set successfully at XPath '{xpath}'."
                        )
                        self.continue_flag = True
                        return self.continue_flag
                    else:
                        ReusableMethods.fprint(f"No element found at XPath '{xpath}'.")
                        self.continue_flag = True
                        return self.continue_flag
                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    text_set = await content.waitForXPath(xpath, timeout=10000)
                    element = await text_set.xpath(xpath)
                    if element:
                        await element[0].type(input_text)
                        ReusableMethods.fprint(
                            f"Text set successfully at XPath '{xpath}'."
                        )
                        self.continue_flag = True
                        return self.continue_flag
                    else:
                        ReusableMethods.fprint(f"No element found at XPath '{xpath}'.")
                        return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "set_text",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + xpath
                )
        return self.continue_flag

    def click_drop_down(self, select_selector, option_value, max_retry=3, iframe=False):
        """This method is to select an option from a drop down
        Args:
            select_selector (str): The xpath of the selector
            option_value (str): The attribute value for the option tag
            max_retry (int, optional): The amount of times it will attempt to select

        Returns:
             response(bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.click_drop_down_pp(select_selector, option_value, max_retry, iframe)
        )
        return response

    async def click_drop_down_pp(
        self, select_selector, option_value, max_retry, iframe
    ):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        attempt = 0
        self.continue_flag = False
        while attempt < max_retry and not self.continue_flag:
            option_xpath = f'{select_selector}/option[@value="{option_value}"]'
            try:
                if not iframe:
                    page = self.current_page

                    # Evaluate XPath to get the <select> element and <option> element
                    select_element = await page.xpath(select_selector)
                    option_element = await page.xpath(option_xpath)

                    # Select the option using the <select> element's value attribute
                    option_value = await (
                        await option_element[0].getProperty("value")
                    ).jsonValue()
                    await page.evaluate(
                        f"""
                        (selectElement, optionValue) => {{
                            selectElement.value = optionValue;
                            selectElement.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        }}
                    """,
                        select_element[0],
                        option_value,
                    )

                    await asyncio.sleep(2)  # Just to see the result
                    ReusableMethods.fprint("Clicked successfully!")
                    return True  # Break the loop if clicked successfully
                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    # Evaluate XPath to get the <select> element and <option> element
                    select_element = await content.xpath(select_selector)
                    option_element = await content.xpath(option_value)

                    # Select the option using the <select> element's value attribute
                    option_value = await (
                        await option_element[0].getProperty("value")
                    ).jsonValue()
                    await content.evaluate(
                        f"""
                        (selectElement, optionValue) => {{
                            selectElement.value = optionValue;
                            selectElement.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        }}
                    """,
                        select_element[0],
                        option_value,
                    )

                    await asyncio.sleep(2)  # Just to see the result
                    ReusableMethods.fprint("Clicked successfully!")
                    self.continue_flag = True
                    return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "PYP Extension an unexpected error happened, on line"
                    + str(line)
                    + ".\nPlease refer to the internal LOG file for more details "
                    + option_xpath
                )
        ReusableMethods.fprint(f"Unable to complete click on element")
        return self.continue_flag

    def send_keys(self, keys: list):
        """This method is to send keys from a keyboard
        Args:
            keys (list): The list of the keys you want to have pressed in the corresponding order. Example:["UP","ENTER"]
        Returns:
             response(bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(self.send_keys_pp(keys))
        return response

    async def send_keys_pp(self, keys):
        """*** For Class use only - Refer to send_keys method instead ***"""
        attempt = 0
        self.continue_flag = False
        while attempt <= 3 and not self.continue_flag:
            try:
                page = self.current_page

                for key in keys:
                    await page.keyboard.press(key)
                    await asyncio.sleep(2)
                ReusableMethods.fprint("Clicked successfully!")
                self.continue_flag = True
                return self.continue_flag  # Break the loop if clicked successfully
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + "unable to hit keys"
                )
        ReusableMethods.fprint(f"Unable to complete click on element")
        return self.continue_flag

    def switch_to_other_window(self, index_number):
        """
        Switches to the other window/tab based on the given index number.
        Note: Allow enough time in your code prior calling this method
              to let the tab fully load in case it is recently opened
        Args:
            index_number (int): The index number of the window/tab to switch to.
            Example: 0 for the first tab, 1 for the second, etc.

        Returns:
            str: The response after switching to the other window/tab.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.switch_to_other_window_pp(index_number)
        )
        return response

    async def switch_to_other_window_pp(self, index_number):
        """Private method to switch the self.browser variable
        to a different window (tab). The index of pages (tabs)
        starts with 0, in other words the first original tab of
        Chrome has the index_number 0"""
        attempt = self.retry
        self.continue_flag = False
        try:
            while attempt < 20 and not self.continue_flag:
                # Wait for a short delay
                await asyncio.sleep(3)
                pages = await self.browser.pages()
                print(pages)
                if index_number < len(pages):
                    self.current_page = pages[index_number]
                    await self.current_page.setViewport({"width": 1920, "height": 1040})
                    self.continue_flag = True
                else:
                    attempt += 1
            if attempt == 20:
                raise Exception(
                    "No window found with index number " + str(index_number)
                )
        except Exception as ex:
            attempt += 1
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            ReusableMethods.fprint(
                "Error switching to a different window",
                "line: " + str(line),
                str(ex),
                log_file_file_path=self.log_file_path,
            )
            raise ValueError(
                "Pyppeteer extension an unexpected error happened on method, "
                + "switch_to_other_window"
                + "please refer to the internal LOG file for more details"
            )
        return self.continue_flag

    def switch_to_main_window(self):
        """
        Switches the focus to the main tab in browser.

        Returns:
            str: The response from switching to the main window.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.switch_to_main_window_pp()
        )
        return response

    async def switch_to_main_window_pp(self):
        """Method to switch the self.browser variable
        back to the original main tab in browser."""
        attempt = self.retry
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                self.current_page = self.original_main_page
                self.continue_flag = True
            except Exception as ex:
                attempt += 1
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "Error switching to the main window",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise ValueError(
                    "Pyppeteer extension an unexpected error happened on method, "
                    + "switch_to_main_window"
                    + "please refer to the internal LOG file for more details"
                )
        return self.continue_flag

    def click(self, xpath, max_retry=3, iframe=False):
        """Use to reload a website in chrome.

        Parameters:
            a (str) wait_element_xpath(str): *Optional - Provide an xpath for the function to fisnish until that element is loaded.

        Returns:
            a (bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.click_pp(xpath, max_retry, iframe)
        )
        return response

    async def click_pp(self, xpath, max_retry, iframe):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        attempt = 0
        self.continue_flag = False
        while attempt < max_retry and not self.continue_flag:
            try:
                if not iframe:
                    page = self.current_page

                    element = await page.waitForXPath(xpath, timeout=10000)
                    await element.click()
                    ReusableMethods.fprint("Clicked successfully!")
                    self.continue_flag = True
                    return self.continue_flag
                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    click = await content.waitForXPath(xpath, timeout=10000)
                    await click.click()
                    ReusableMethods.fprint(f"Element exists on page")
                    self.continue_flag = True
                    return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + xpath
                )
        ReusableMethods.fprint(f"Unable to complete click on element")
        return self.continue_flag

    def element_exist(self, xpath, wait_time=1, iframe=False):
        """Use to reload a website in chrome.

        Parameters:
            a (str) wait_element_xpath(str): *Optional - Provide an xpath for the function to fisnish until that element is loaded.

        Returns:
            a (bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.element_exist_pp(xpath, wait_time, iframe)
        )
        return response

    async def element_exist_pp(self, xpath, wait_time, iframe):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        attempt = 0
        wait = wait_time * 1000
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                if not iframe:
                    ReusableMethods.fprint(xpath)
                    page = self.current_page

                    await page.waitForXPath(xpath, timeout=wait)
                    ReusableMethods.fprint(f"Element exists on page")
                    self.continue_flag = True
                    return self.continue_flag
                else:
                    ReusableMethods.fprint(xpath)
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    await content.waitForXPath(xpath, timeout=wait)
                    ReusableMethods.fprint(f"Element exists on page")
                    self.continue_flag = True
                    return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(5)
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
        ReusableMethods.fprint(f"Element not on page")
        return self.continue_flag

    def get_text(self, xpath="", iframe=False):
        """Use to reload a website in chrome.

        Parameters:
            a (str) wait_element_xpath(str): *Optional - Provide an xpath for the function to fisnish until that element is loaded.

        Returns:
            a (bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.get_text_pp(xpath, iframe)
        )
        return response

    async def get_text_pp(self, xpath, iframe):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        attempt = 0
        self.continue_flag = False
        text = ""
        while attempt < 3 and not self.continue_flag:
            try:
                if not iframe:
                    page = self.current_page

                    if xpath:
                        elements = await page.xpath(xpath)
                        for element in elements:
                            # Check if the element is an input field
                            tag_name = await page.evaluate(
                                "(element) => element.tagName", element
                            )
                            if tag_name.lower() == "input":
                                # If it's an input field, get the value attribute
                                text += await page.evaluate(
                                    "(element) => element.value", element
                                )
                            else:
                                text += await page.evaluate(
                                    "(element) => element.textContent", element
                                )
                    else:
                        text = await page.evaluate("() => document.body.textContent")
                    self.continue_flag = True
                    return text
                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    if xpath:
                        elements = await content.xpath(xpath)
                        for element in elements:
                            text += await content.evaluate(
                                "(element) => element.textContent", element
                            )
                    else:
                        text = await content.evaluate("() => document.body.textContent")
                    return text
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(5)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + xpath
                )
        ReusableMethods.fprint(f"No xpath found to get text")
        return text

    def take_screenshot(self, path, iframe=False):
        """This method takes screenshot of the screen
        Args:
            path (str): The path where you want to save.

        Returns:
             response(bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.take_screenshot_pp(path, iframe)
        )
        return response

    async def take_screenshot_pp(self, path, iframe):
        """*** For Class use only - Refer to take_screenshot method instead ***"""
        attempt = 0
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                if not iframe:
                    page = self.current_page

                    await page.screenshot({"path": path})
                    self.continue_flag = True
                    return self.continue_flag
                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    await content.screenshot({"path": path})
                    ReusableMethods.fprint(f"Element exists on page")
                    self.continue_flag = True
                    return self.continue_flag
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(5)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                )
        ReusableMethods.fprint(f"No xpath found to get text")
        return self.continue_flag

    def accept_alert(self, xpath, timeout=3):
        """This method is to accept an alert popup
        Args:
            timeout (int, optional): The amount of seconds it will wait for the alert. Defaulted to "3"

        Returns:
             response(bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.accept_alert_pp(xpath, timeout)
        )
        return response

    async def accept_alert_pp(self, xpath, timeout=3):
        attempt = 0
        wait_for = timeout * 1000
        self.continue_flag = False
        while attempt <= 3 and not self.continue_flag:
            try:
                page = self.current_page

                # Intercept dialog events (alerts, confirms, prompts, etc.)
                await page.evaluateOnNewDocument(
                    "() => {"
                    "   window.alert = () => {};"
                    "   window.confirm = () => true;"
                    "   window.prompt = () => null;"
                    "}"
                )
                # Handle any alerts that may appear (optional)
                page.on("dialog", lambda dialog: asyncio.ensure_future(dialog.accept()))
                element = await page.waitForXPath(
                    xpath, visibility=True, timeout=wait_for
                )
                # Use page.evaluate to click on the element
                await element.click()
                await asyncio.sleep(1)  # Adjust the sleep time if necessary
                ReusableMethods.fprint("Clicked successfully!")
                self.continue_flag = True
                return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + xpath
                )
        ReusableMethods.fprint(f"Unable to complete click on element")
        return self.continue_flag

    def close_all_browsers(self):
        """Use to reload a website in chrome.

        Parameters:
            a (str) wait_element_xpath(str): *Optional - Provide an xpath for the function to finish until that element is loaded.

        Returns:
            a (bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.close_all_browsers_pp()
        )
        return response

    async def close_all_browsers_pp(self):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        attempt = 0
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                pages = await self.browser.pages()
                for page in pages:
                    await page.close()
                await self.browser.close()
                self.continue_flag = True
                return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(5)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
        ReusableMethods.fprint(f"Can't close the browser")
        return self.continue_flag

    def authenticate_in_alert_window(self, username, password, page_url):
        """
        Authenticates the user in the alert window.

        Args:
            username (str): The username to be entered in the alert window.
            password (str): The password to be entered in the alert window.
            page_url (str): The URL of the page where the alert window is present.

        Returns:
            str: The response from the authentication process.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.authenticate_in_alert_window_pp(username, password, page_url)
        )
        return response

    async def authenticate_in_alert_window_pp(self, username, password, page_url):
        attempt = 0
        self.continue_flag = False
        while attempt <= 3 and not self.continue_flag:
            try:
                page = self.current_page
                #  HTTP Basic Authentication using username and password
                await page.authenticate({"username": username, "password": password})
                await asyncio.sleep(2)
                # Navigate to authenticated webpage
                await page.goto(page_url)
                await page.setViewport({"width": 1920, "height": 1040})
                self.continue_flag = True
                return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + "Unable to perform HTTP Basic Authentication in page"
                )
        ReusableMethods.fprint(f"Unable to complete click on element")
        return self.continue_flag

    def close_current_tab(self):
        """
        Closes the current tab in the browser.

        Returns:
            str: The response from closing the tab.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.close_current_tab_pp()
        )
        return response

    async def close_current_tab_pp(self):
        """
        Closes the current tab in the browser.

        Returns:
            bool: True if the tab was successfully closed, False otherwise.
        """
        attempt = 0
        self.continue_flag = False
        while attempt <= 3 and not self.continue_flag:
            try:
                # Get all pages (tabs)
                pages = await self.browser.pages()

                # Close the current tab
                await self.current_page.close()
                # Update the current_page to the previous tab
                if len(pages) > 1:
                    self.current_page = pages[-2]
                else:
                    self.current_page = None
                self.continue_flag = True
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details: "
                    + "Unable to close current page in browser."
                )
            ReusableMethods.fprint(f"Unable to complete click on element")
        return self.continue_flag

    def extract_table_column_attribute(
        self, xpath, attribute="text", remove_headers=False
    ):
        """
        Extracts the specified attribute from a column in a table using the given XPath.

        Args:
            xpath (str): The XPath expression to locate the column elements (Usually for multiple elements).
            attribute (str, optional): The attribute to extract from each cell in the column. Defaults to "text".
            remove_headers (bool, optional): Whether to remove the header row from the extracted data. Defaults to False.

        Returns:
            list: The extracted attribute values from the column.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.extract_table_column_attribute_pp(xpath, attribute, remove_headers)
        )
        return response

    async def extract_table_column_attribute_pp(self, xpath, attribute, remove_headers):
        """
        Extracts the specified attribute from a column of elements identified by the given XPath.

        Args:
            xpath (str): The XPath expression to locate the column elements (Usually for multiple elements).
            attribute (str): The attribute to extract from each element in the column (text is default value).
            remove_headers (bool): Flag indicating whether to remove the header element from the extracted column.

        Returns:
            list: A list of attribute values extracted from the column elements.
        """
        column_information = []
        attempt = self.retry
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                page = self.current_page
                await page.waitForXPath(xpath)
                column_elements = await page.xpath(xpath)

                for element in column_elements:
                    attribute_value = "[empty]"
                    if attribute == "text":
                        attribute_value = await (
                            await element.getProperty("textContent")
                        ).jsonValue()
                    else:
                        attribute_value = await (
                            await element.getProperty(attribute)
                        ).jsonValue()
                    print(attribute_value)
                    column_information.append(attribute_value)
                if remove_headers:
                    column_information = column_information[1:]
                self.continue_flag = True
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details: "
                    + "Unable to extract table column information"
                )
        return column_information

    def page_html(self):
        """
        Capture page html

        Returns:
            str: The response from closing the tab and html.
        """
        response = False
        html = None
        response = self.set_current_page_pp()
        if response:
            response = False
            response, html = asyncio.get_event_loop().run_until_complete(
                self.page_html_pp()
            )
        return response, html

    async def page_html_pp(self):
        """
        Capture page html.

        Returns:
            bool: True if the tab was successfully closed, False otherwise.
            html: HTML Code
        """
        attempt = 0
        self.continue_flag = False
        html = None
        while attempt <= 3 and not self.continue_flag:
            try:
                # Get all html
                html = await self.current_page.content()
                self.continue_flag = True
                return self.continue_flag, html
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details: "
                    + "Unable to close current page in browser."
                )
        ReusableMethods.fprint(f"Unable to complete html extraction")
        return self.continue_flag, html

    def bring_page_to_front(self):
        """
        Bring page to front

        Returns:
            bool: True if the tab was successfully moved, False otherwise.
        """
        response = False
        response = self.set_current_page_pp()
        if response:
            response = False
            response = asyncio.get_event_loop().run_until_complete(
                self.bring_page_to_front_pp()
            )
        return response

    async def bring_page_to_front_pp(self):
        """
        Bring page to front

        Returns:
            bool: True if the tab was successfully moved, False otherwise.
        """
        attempt = 0
        self.continue_flag = False
        while attempt <= 3 and not self.continue_flag:
            try:
                # Get all html
                await self.current_page.bringToFront()
                self.continue_flag = True
                return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details: "
                    + "Unable to close current page in browser."
                )
        ReusableMethods.fprint(f"Unable to complete html extraction")
        return self.continue_flag

    def extract_table_data(self, xpath):
        """
        Extracts the data from a table using the given XPath.

        Args:
            xpath (str): The XPath expression to locate the table elements.

        Returns:
            list: The extracted data from the table.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.extract_table_data_pp(xpath)
        )
        return response

    async def extract_table_data_pp(self, xpath):
        """
        Extracts the data from a table using the given XPath and saves it in a DataFrame.

        Args:
            xpath (str): The XPath expression to locate the table elements.

        Returns:
            DataFrame: The extracted data from the table.
        """
        table_data = []
        headers = []
        attempt = self.retry
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                page = self.current_page
                await page.waitForXPath(xpath)
                table_rows = await page.xpath(f"{xpath}//tr")

                for i, row in enumerate(table_rows):
                    cols = await row.xpath(".//td | .//th")
                    row_data = []
                    a_tags = []
                    for col in cols:
                        a_tag = await col.xpath(".//a")
                        if a_tag:
                            for a in a_tag:
                                a_tags.append(
                                    await page.evaluate("(element) => element.href", a)
                                )
                            # a_tags.append(await page.evaluate("(element) => element.href", a_tag[0]))
                        row_data.append(
                            await page.evaluate("(element) => element.textContent", col)
                        )
                    if i == 0:
                        headers = row_data + ["aTags"]
                    else:
                        table_data.append(row_data + [a_tags])
                    self.continue_flag = True
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details: "
                    + "Unable to extract table data"
                )
        df = pd.DataFrame(table_data, columns=headers)
        return df

    def check_checkbox(self, xpath):
        """
        Check the checkbox element using the given XPath.

        Args:
            xpath (str): The XPath expression to locate the checkbox element.

        Returns:
            bool: True if the checkbox was successfully checked, False otherwise.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.check_checkbox_pp(xpath)
        )
        return response

    async def check_checkbox_pp(self, xpath):
        """
        Checks if a checkbox is selected using the given XPath.

        Args:
            xpath (str): The XPath expression to locate the checkbox.

        Returns:
            bool: True if the checkbox is selected, False otherwise.
        """
        attempt = self.retry
        self.continue_flag = False

        while attempt < 3 and not self.continue_flag:
            try:
                page = self.current_page
                await page.waitForXPath(xpath)
                checkbox_element = await page.xpath(xpath)

                # Get the checked property of the checkbox
                is_checked = await page.evaluate(
                    "(element) => element.checked", checkbox_element[0]
                )

                self.continue_flag = True
                return is_checked

            except (pyppeteer.errors.NetworkError, pyppeteer.errors.TimeoutError) as ex:
                attempt += 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                if isinstance(ex, pyppeteer.errors.TimeoutError):
                    await asyncio.sleep(3)

            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details: "
                    + "Unable to check checkbox"
                )

        return False

    def get_header_index(self, xpath, header_text):
        """
        Get the index of the header in the table using the given XPath.

        Args:
            xpath (str): The XPath expression to locate the table elements.
            header_text (str): The text of the header to find the index of.

        Returns:
            int: The index of the header in the table.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.get_header_index_pp(xpath, header_text)
        )
        return response

    async def get_header_index_pp(self, xpath, header_text):
        """
        Extracts the index of a td element in a table based on a specific text.

        Args:
            xpath (str): The XPath of the table.
            header_text (str): The text of the header.

        Returns:
            int: The index of the header, or -1 if the header is not found.
        """
        attempt = self.retry
        self.continue_flag = False

        while attempt < 3 and not self.continue_flag:
            try:
                page = self.current_page
                await page.waitForXPath(xpath)
                headers = await page.xpath(f"{xpath}/thead/tr/td")

                for i, header in enumerate(headers):
                    text = await page.evaluate(
                        "(element) => element.textContent", header
                    )
                    if text.strip() == header_text:
                        self.continue_flag = True
                        return i + 1  # Add 1 to the index to get the column number

                attempt += 1

            except (pyppeteer.errors.NetworkError, pyppeteer.errors.TimeoutError):
                attempt += 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                if isinstance(ex, pyppeteer.errors.TimeoutError):
                    await asyncio.sleep(3)

            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "get_header_index",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details: "
                    + "Unable to extract header index"
                )

        return -1

    def select_dropdown_option(
        self, select_xpath, option_text, max_retry=3, iframe=False
    ):
        """
         Clicks a dropdown and selects an option based on the text.

        Args:
            dropdown_xpath (str): The XPath of the dropdown.
            option_text (str): The text of the option.

        Returns:
            None
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.select_dropdown_option_pp(select_xpath, option_text, max_retry, iframe)
        )
        return response

    async def select_dropdown_option_pp(
        self, select_xpath, option_text, max_retry=3, iframe=False
    ):
        """
        Clicks a dropdown and selects an option based on the text.

        Args:
            dropdown_xpath (str): The XPath of the dropdown.
            option_text (str): The text of the option.

        Returns:
            None
        """
        attempt = 0
        max_retry = 3
        self.continue_flag = False
        while attempt < max_retry and not self.continue_flag:
            option_xpath = f'{select_xpath}/option[normalize-space(.)="{option_text}"]'
            try:
                if not iframe:
                    page = self.current_page

                    # Evaluate XPath to get the <select> element and <option> element
                    select_element = await page.xpath(select_xpath)
                    option_element = await page.xpath(option_xpath)

                    # Select the option using the <select> element's value attribute
                    option_value = await (
                        await option_element[0].getProperty("value")
                    ).jsonValue()
                    await page.evaluate(
                        f"""
                        (selectElement, optionValue) => {{
                            selectElement.value = optionValue;
                            selectElement.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        }}
                    """,
                        select_element[0],
                        option_value,
                    )

                    await asyncio.sleep(2)  # Just to see the result
                    ReusableMethods.fprint("Clicked successfully!")
                    return True  # Break the loop if clicked successfully
                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    # Evaluate XPath to get the <select> element and <option> element
                    select_element = await content.xpath(select_xpath)
                    option_element = await content.xpath(option_xpath)

                    # Select the option using the <select> element's value attribute
                    option_value = await (
                        await option_element[0].getProperty("value")
                    ).jsonValue()
                    await content.evaluate(
                        f"""
                        (selectElement, optionValue) => {{
                            selectElement.value = optionValue;
                            selectElement.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        }}
                    """,
                        select_element[0],
                        option_value,
                    )

                    await asyncio.sleep(2)  # Just to see the result
                    ReusableMethods.fprint("Clicked successfully!")
                    self.continue_flag = True
                    return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "PYP Extension an unexpected error happened, on line"
                    + str(line)
                    + ".\nPlease refer to the internal LOG file for more details "
                    + option_xpath
                )
        ReusableMethods.fprint(f"Unable to complete click on element")
        return self.continue_flag

    def select_items_from_list(
        self, select_selector, list_options, max_retry=3, iframe=False
    ):
        """Selects an option from a list
        Args:
            select_selector (str): The xpath of the selector
            list_options (list): The list of options to select
            max_retry (int, optional): The amount of times it will attempt to select

        Returns:
             response(bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.select_items_from_list_pp(
                select_selector, list_options, max_retry, iframe
            )
        )
        return response

    async def select_items_from_list_pp(
        self, select_selector, list_options, max_retry, iframe
    ):
        attempt = 0
        self.continue_flag = False
        while attempt < max_retry and not self.continue_flag:
            try:
                if not iframe:
                    page = self.current_page

                    # Press and hold the control key
                    await page.keyboard.down("Control")

                    # Iterate over the list of options and select each one
                    for option_index in list_options:
                        # Select each option
                        await page.click(f"{select_selector}/option[{option_index}]")

                    # Release the control key
                    await page.keyboard.up("Control")

                    await asyncio.sleep(2)  # Just to see the result
                    ReusableMethods.fprint("Clicked successfully!")

                    return True  # Break the loop if clicked successfully

                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    # Evaluate XPath to get the <select> element
                    select_element = await content.xpath(select_selector)

                    # Select the options using the <select> element's value attribute
                    for option_value in list_options:
                        await content.evaluate(
                            f"""
                            (selectElement, optionValue) => {{
                                selectElement.value = optionValue;
                                selectElement.dispatchEvent(new Event('change', {{ bubbles: true }}));
                            }}
                        """,
                            select_element[0],
                            option_value,
                        )

                    await asyncio.sleep(2)  # Just to see the result
                    ReusableMethods.fprint("Clicked successfully!")
                    self.continue_flag = True
                    return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "PYP Extension an unexpected error happened, on line"
                    + str(line)
                    + ".\nPlease refer to the internal LOG file for more details "
                    + list_options
                )
        ReusableMethods.fprint(f"Unable to complete click on element")
        return self.continue_flag

    def get_checkbox_status(self, xpath, iframe=False):
        """Returns the status of a checkbox
        Args:
            xpath (str): The xpath of the checkbox

        Returns:
             response(bool) This method returns a True/False. Depending on the succeess of the operation.
        """
        response = asyncio.get_event_loop().run_until_complete(
            self.get_checkbox_status_pp(xpath, iframe)
        )
        return response

    async def get_checkbox_status_pp(self, xpath, iframe):
        """*** For Class use only - Refer to refresh_web_page method instead ***"""
        attempt = 0
        self.continue_flag = False
        while attempt < 3 and not self.continue_flag:
            try:
                if not iframe:
                    page = self.current_page

                    element = await page.waitForXPath(xpath, timeout=10000)
                    status = await page.evaluate(
                        "(element) => element.checked", element
                    )
                    self.continue_flag = status
                    return self.continue_flag
                else:
                    page = self.current_page

                    frames = await page.waitForXPath(iframe, timeout=10000)
                    content = await frames.contentFrame()
                    element = await content.waitForXPath(xpath, timeout=10000)
                    status = await content.evaluate(
                        "(element) => element.checked", element
                    )
                    self.continue_flag = status
                    return self.continue_flag
            except pyppeteer.errors.NetworkError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
            except pyppeteer.errors.TimeoutError:
                attempt = attempt + 1
                ReusableMethods.fprint(
                    f"Attempt {attempt}: Timed out while waiting for element."
                )
                await asyncio.sleep(3)
            except Exception as ex:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                ReusableMethods.fprint(
                    "click",
                    "line: " + str(line),
                    str(ex),
                    log_file_file_path=self.log_file_path,
                )
                raise Exception(
                    "Pyppeteer Extension an unexpected error happened, "
                    + "please refer to the internal LOG file for more details "
                    + xpath
                )
        ReusableMethods.fprint(f"Unable to complete click on element")
        return self.continue_flag

    # For testing only

    # driver = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    # web_page_url = "https://login.salesforce.com/ "

    # #need to start the browser outside the class
    # browser = asyncio.get_event_loop().run_until_complete(pyppeteer.launch(headless = False, executablePath = driver))

    # #Instanciate the class
    # pyppeteer_extensions = pyppeteer_extension(browser)

    # pyppeteer_extensions.open_web_page(web_page_url)
    # #asyncio.get_event_loop().run_until_complete(pyppeteer_extensions.open_web_page(web_page_url))

    # input("wait")
