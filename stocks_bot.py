import asyncio
import json
import os
import time

import pyppeteer
from dotenv import load_dotenv
from openpyxl import load_workbook

from libs.files_and_folders_utils import files_and_folders
from libs.pyppeteer_class import pyppeteer_extension


class StocksBot:
    def __init__(self, json_file_path):
        """
        Initialize the StocksBot with configuration from a JSON file.

        Args:
            json_file_path (str): Path to the JSON configuration file.
        """
        try:
            with open(json_file_path) as read_file:
                data = json.load(read_file)
        except FileNotFoundError:
            print(f"Configuration file not found: {json_file_path}")
            return
        except json.JSONDecodeError:
            print(f"Error decoding JSON from the file: {json_file_path}")
            return

        # Read variables from .env file
        self.login_email = os.getenv("login_email")
        self.login_password = os.getenv("login_password")
        # Read variables from config.json file
        self.file_name = data.get("file_name")
        chrome_executable_path = data.get("chrome_executable_path")

        # set the custom exception handler for the event loop
        loop = asyncio.get_event_loop()
        loop.set_exception_handler(self.custom_exception_handler)

        # set up log file
        self.utils = files_and_folders(data, "StocksBot")

        # Set pyppeteer to use the custom event loop
        try:
            self.browser = loop.run_until_complete(
                pyppeteer.launch(
                    headless=False,
                    executablePath=chrome_executable_path,
                    args=["--start-maximized"],
                )
            )
        except pyppeteer.errors.BrowserError as e:
            print(f"Error launching browser: {e}")
            return

        self.pyppeteer_extensions = pyppeteer_extension(
            self.browser, self.utils.log_file_path
        )
        self.start_automation()

    def custom_exception_handler(self, loop, context):
        """
        Custom exception handler for the event loop.

        Args:
            loop (asyncio.AbstractEventLoop): The event loop.
            context (dict): The context of the exception.
        """
        print(f"Exception in event loop: {context['message']}")

    def login(self):
        """
        Log in to the website using the provided credentials.

        Returns:
            bool: True if login is successful, False otherwise.
        """
        try:
            self.pyppeteer_extensions.open_web_page(
                "https://app.koyfin.com/login", '//input[@type="email"]'
            )
            self.pyppeteer_extensions.set_text(
                '//input[@type="email"]', self.login_email
            )
            time.sleep(0.5)
            self.pyppeteer_extensions.set_text(
                '//input[@name="password"]', self.login_password
            )
            self.pyppeteer_extensions.click('//label[text()="Sign in"]')
            self.pyppeteer_extensions.wait_for_element(
                '//div[text()="Search for a name, ticker, or function "]'
            )
            if self.pyppeteer_extensions.element_exist('//button[text()="Accept All"]'):
                self.pyppeteer_extensions.click('//button[text()="Accept All"]')

            return True
        except Exception as e:
            print(f"Login failed: {e}")
            return False

    def start_automation(self):
        """
        Start the automation process to gather stock data and write it to an Excel file.
        """
        char = 67
        tickers_list = self.get_ticker_list()
        if not tickers_list:
            print("No tickers found.")
            return

        login_success = self.login()
        if not login_success:
            return

        for ticker in tickers_list:
            try:
                self.collect_and_write_data(ticker, char)
                char += 1
            except Exception as e:
                print(f"Error processing ticker {ticker}: {e}")

    def collect_and_write_data(self, ticker, char):
        """
        Collect stock data for a given ticker and write it to the Excel file.

        Args:
            ticker (str): The stock ticker symbol.
            char (int): The column character code in the Excel sheet.
        """
        pe_ratio = ""
        pb_ratio = ""
        stock_price = ""
        current_ratio = ""
        return_on_equity = ""
        asset_turnover = ""
        eps_current_fiscal_year = ""
        eps_last_fiscal_year = ""
        total_liabilities_total_assets = ""
        self.pyppeteer_extensions.click(
            '//div[text()="Search for a name, ticker, or function "]'
        )
        self.pyppeteer_extensions.wait_for_element(
            '//input[@placeholder="Type by name or ticker"]'
        )
        self.pyppeteer_extensions.set_text(
            '//input[@placeholder="Type by name or ticker"]', ticker
        )
        self.pyppeteer_extensions.wait_for_element(
            f'(//div[@class="rc-dialog-content"]/descendant::div[text()="{ticker}"])[1]'
        )
        self.pyppeteer_extensions.click(
            f'(//div[@class="rc-dialog-content"]/descendant::div[text()="{ticker}"])[1]'
        )
        self.pyppeteer_extensions.wait_for_element(
            '//div[@class="console-popup__functionTitle___wZJdU"][text()="Overview"]'
        )
        self.pyppeteer_extensions.click(
            '//div[@class="console-popup__functionTitle___wZJdU"][text()="Overview"]'
        )
        self.pyppeteer_extensions.wait_for_element(
            '//div[text()="P/E"]/following-sibling::div/div/div'
        )
        while not pe_ratio:
            pe_ratio = self.pyppeteer_extensions.get_text(
                '//div[text()="P/E"]/following-sibling::div/div/div'
            )
        self.pyppeteer_extensions.wait_for_element(
            '//div[text()="Price/Book"]/following-sibling::div/div/div'
        )
        while not pb_ratio:
            pb_ratio = self.pyppeteer_extensions.get_text(
                '//div[text()="Price/Book"]/following-sibling::div/div/div'
            )
        self.pyppeteer_extensions.wait_for_element(
            f'//span[text()="{ticker}"]/parent::div/div/label'
        )
        while not stock_price:
            stock_price = self.pyppeteer_extensions.get_text(
                f'//span[text()="{ticker}"]/parent::div/div/label'
            )
        time.sleep(2)
        self.pyppeteer_extensions.wait_for_element(
            '//div[text()="Financial Analysis"]/parent::div/parent::div'
        )
        self.pyppeteer_extensions.click(
            '//div[text()="Financial Analysis"]/parent::div/parent::div'
        )
        time.sleep(1)
        self.pyppeteer_extensions.click('//div[text()="Profitability"]')
        self.pyppeteer_extensions.wait_for_element(
            '(//div[text()="Current Ratio"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
        )
        while not current_ratio:
            current_ratio = self.pyppeteer_extensions.get_text(
                '(//div[text()="Current Ratio"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
            )
        self.pyppeteer_extensions.wait_for_element(
            '(//div[text()="Return On Equity"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
        )
        while not return_on_equity:
            return_on_equity = self.pyppeteer_extensions.get_text(
                '(//div[text()="Return On Equity"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
            )
        self.pyppeteer_extensions.wait_for_element(
            '(//div[text()="Asset Turnover"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
        )
        while not asset_turnover:
            asset_turnover = self.pyppeteer_extensions.get_text(
                '(//div[text()="Asset Turnover"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
            )
        time.sleep(1)
        self.pyppeteer_extensions.click('//span[text()="Highlights"]')
        self.pyppeteer_extensions.wait_for_element(
            '(//div[text()="Diluted EPS"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
        )
        while not eps_current_fiscal_year:
            eps_current_fiscal_year = self.pyppeteer_extensions.get_text(
                '(//div[text()="Diluted EPS"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
            )
        self.pyppeteer_extensions.wait_for_element(
            '(//div[text()="Diluted EPS"]/parent::div/parent::div/parent::div/parent::div/div)[position()=last()-3]'
        )
        while not eps_last_fiscal_year:
            eps_last_fiscal_year = self.pyppeteer_extensions.get_text(
                '(//div[text()="Diluted EPS"]/parent::div/parent::div/parent::div/parent::div/div)[position()=last()-3]'
            )
        time.sleep(2)
        self.pyppeteer_extensions.click('//span[text()="Solvency"]')
        self.pyppeteer_extensions.wait_for_element(
            '(//div[text()="Total Liabilities / Total Assets"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
        )
        while not total_liabilities_total_assets:
            total_liabilities_total_assets = self.pyppeteer_extensions.get_text(
                '(//div[text()="Total Liabilities / Total Assets"]/parent::div/parent::div/parent::div/parent::div/div)[last()]'
            )
        time.sleep(3)
        # Clean up the data
        pe_ratio = pe_ratio.replace(",", "")
        pb_ratio = pb_ratio.replace(",", "")
        stock_price = stock_price.replace(",", "")
        current_ratio = current_ratio.replace(",", "")
        return_on_equity = return_on_equity.replace(",", "")
        asset_turnover = asset_turnover.replace(",", "")
        eps_current_fiscal_year = eps_current_fiscal_year.replace(",", "")
        eps_last_fiscal_year = eps_last_fiscal_year.replace(",", "")
        total_liabilities_total_assets = total_liabilities_total_assets.replace(",", "")
        asset_turnover = asset_turnover.replace("x", "")
        current_ratio = current_ratio.replace("x", "")
        return_on_equity = return_on_equity.replace("%", "")
        total_liabilities_total_assets = total_liabilities_total_assets.replace("%", "")
        pe_ratio = float(pe_ratio) if pe_ratio != "Upgrade" else pe_ratio
        pb_ratio = float(pb_ratio) if pb_ratio != "Upgrade" else pb_ratio
        stock_price = float(stock_price) if stock_price != "Upgrade" else stock_price
        current_ratio = (
            float(current_ratio) if current_ratio != "Upgrade" else current_ratio
        )
        return_on_equity = (
            float(return_on_equity)
            if return_on_equity != "Upgrade"
            else return_on_equity
        )
        asset_turnover = (
            float(asset_turnover) if asset_turnover != "Upgrade" else asset_turnover
        )
        eps_current_fiscal_year = (
            float(eps_current_fiscal_year)
            if eps_current_fiscal_year != "Upgrade"
            else eps_current_fiscal_year
        )
        eps_last_fiscal_year = (
            float(eps_last_fiscal_year)
            if eps_last_fiscal_year != "Upgrade"
            else eps_last_fiscal_year
        )
        total_liabilities_total_assets = (
            float(total_liabilities_total_assets)
            if total_liabilities_total_assets != "Upgrade"
            else total_liabilities_total_assets
        )
        # Write the data to the Excel file

        self.write_to_excel(
            char,
            pe_ratio,
            pb_ratio,
            stock_price,
            current_ratio,
            return_on_equity,
            asset_turnover,
            eps_current_fiscal_year,
            eps_last_fiscal_year,
            total_liabilities_total_assets,
        )

    def write_to_excel(
        self,
        char,
        pe_ratio,
        pb_ratio,
        stock_price,
        current_ratio,
        return_on_equity,
        asset_turnover,
        eps_current_fiscal_year,
        eps_last_fiscal_year,
        total_liabilities_total_assets,
    ):
        """
        Write the collected stock data to the Excel file.

        Args:
            char (int): The column character code in the Excel sheet.
            pe_ratio (str): The P/E ratio.
            pb_ratio (str): The Price/Book ratio.
            stock_price (str): The stock price.
            current_ratio (str): The current ratio.
            return_on_equity (str): The return on equity.
            asset_turnover (str): The asset turnover.
            eps_current_fiscal_year (str): The EPS for the current fiscal year.
            eps_last_fiscal_year (str): The EPS for the last fiscal year.
            total_liabilities_total_assets (str): The total liabilities to total assets ratio.
        """
        try:
            workbook = load_workbook(filename=self.file_name)
            sheet = workbook.active
            column = chr(char)
            sheet[column + "3"] = pe_ratio
            sheet[column + "4"] = eps_last_fiscal_year
            sheet[column + "5"] = eps_current_fiscal_year
            sheet[column + "8"] = pb_ratio
            sheet[column + "9"] = return_on_equity
            sheet[column + "10"] = current_ratio
            sheet[column + "11"] = total_liabilities_total_assets
            sheet[column + "12"] = asset_turnover
            sheet[column + "15"] = stock_price
            workbook.save(self.file_name)
        except Exception as e:
            print(f"Error writing to Excel file: {e}")

    def get_ticker_list(self):
        """
        Get the list of stock tickers from the Excel file.

        Returns:
            list: A list of stock ticker symbols.
        """
        try:
            workbook = load_workbook(filename=self.file_name)
            sheet = workbook.active
            tickers_list = [
                str(sheet["C1"].value),
                str(sheet["D1"].value),
                str(sheet["E1"].value),
                str(sheet["F1"].value),
            ]
            return tickers_list
        except FileNotFoundError:
            print(f"Excel file not found: {self.file_name}")
            return []
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return []


if __name__ == "__main__":
    # Load environment variables from .env file
    load_dotenv()
    json_file_path = "C:\\stocks_bot\\config.json"
    start_time = time.perf_counter()
    bot = StocksBot(json_file_path)
    end_time = time.perf_counter()
    total_time = end_time - start_time
    print(f"El bot ha tardado {total_time:.2f} segundos en ejecutarse.")
