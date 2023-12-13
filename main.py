import os
import time
import warnings

import pandas as pd
from decouple import config
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# Suppress openpyxl UserWarning
warnings.simplefilter("ignore", category=UserWarning)

BASE_URL = "https://auth.synergysportstech.com/Account/Login"
leaderbord_page = "https://apps.synergysports.com/basketball/leaderboards?leagueId=54457dce300969b132fcfb37&seasonId=64da359a0d288f7495c0bdc9&competitionIds=non-exhibition-54457dce300969b132fcfb37&comparisonGroupId=648ac7b0a79824aa31db2b35"


class SynergyBot:
    def __init__(self, teardown=True):
        # Specify the path to the locally installed ChromeDriver binary
        s = Service(ChromeDriverManager().install())

        self.options = webdriver.ChromeOptions()
        self.options.add_argument('headless')
        self.options.add_experimental_option("detach", True)
        self.options.add_experimental_option("excludeSwitches", ['enable-logging'])

        # Use the specified service and options to create the Chrome WebDriver
        self.login_url = "https://auth.synergysportstech.com/Account/Login"
        self.leaderboard_page = leaderbord_page

        # Set the download path
        self.download_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'downloads')

        print(self.download_path)
        self.options.add_experimental_option("prefs", {
            "download.default_directory": self.download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        self.driver = webdriver.Chrome(options=self.options, service=s)

        self.driver.implicitly_wait(50)
        super(SynergyBot, self).__init__()

    def __enter__(self):
        self.driver.get(BASE_URL)

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.teardown:
            self.driver.quit()

    def land_first_page(self):
        self.driver.get(BASE_URL)

    def fill_login_form(self):
        self.driver.get(self.login_url)
        username_field = self.driver.find_element("id", "Username")
        password_field = self.driver.find_element("id", "Password")
        remember_login_checkbox = self.driver.find_element("id", "RememberLogin")

        # Clear existing input (optional)
        username_field.clear()
        password_field.clear()

        # Fill in the form fields
        username = config("EMAIL", default="")
        password = config("PASSWORD", default="")

        username_field.send_keys(username)
        password_field.send_keys(password)

        remember_login_checkbox.click()

    def submit_login_form(self):
        login_button = self.driver.find_element("name", "button")
        login_button.click()

    def select_team_tag(self, text):
        self.driver.get(self.leaderboard_page)
        # Find the ng-select dropdown element
        ng_select = self.driver.find_element(
            by=By.CSS_SELECTOR,
            value='ng-select[class="ng-select-searchable ng-select ng-select-single ng-untouched ng-pristine ng-valid"]')

        # Click on the ng-select to open the dropdown
        ng_select.click()

        # Wait for the dropdown options to be visible
        dropdown_options = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_all_elements_located((By.CLASS_NAME, "ng-option"))
        )

        # Find and click on the "Team Offensive" option
        for option in dropdown_options:
            if option.text == text:
                option.click()
                break

        time.sleep(3)
        self.click_on_all_play_types(text)

    def click_on_all_play_types(self, text):

        for stats_counter in range(0, 11):  # Get the stats
            element = self.driver.find_elements(By.CSS_SELECTOR, 'div[class="p-2 xl:p-4"]')[2]
            overall_selector = element.find_element(by=By.CSS_SELECTOR, value='div[class="ng-select"]')
            overall_selector.click()

            player_type_unordered_list = \
                self.driver.find_elements(by=By.CSS_SELECTOR, value='div[class="mt-6 px-3 ng-star-inserted"]')[1]
            player_stat = player_type_unordered_list.find_elements(by=By.TAG_NAME, value="li")[stats_counter]

            print(player_stat.get_attribute("innerHTML"))  # click the stat
            player_stat.click()
            time.sleep(2)

            self.download_play_type()
            if stats_counter == 10:
                print("Done With ", text)
                time.sleep(5)

    def download_play_type(self):
        element = self.driver.find_elements(by=By.CSS_SELECTOR, value='div[class="p-2 xl:p-4"]')[4]
        download_button = element.find_element(by=By.CSS_SELECTOR, value='div[class="ng-select"]')
        download_button.click()
        time.sleep(2)

    def get_files(self):
        files = []
        for f in os.listdir(self.download_path):
            if f.endswith('.xlsx'):
                files.append(f)
        return files

    def get_sheet_name(self, file):
        try:
            base_name = os.path.splitext(file)[0]
            if len(base_name) > 31:
                base_name = base_name[64:94]
        except:
            base_name = base_name[:31]
        return base_name

    def merge_data(self):
        # Loop through each file, read it into a DataFrame, and write it to a new Excel file with
        # a sheet name corresponding to the file name
        with pd.ExcelWriter('merged_file.xlsx') as writer:
            files = self.get_files()
            for file in files:
                try:
                    df = pd.read_excel(os.path.join(self.download_path, file), engine='openpyxl')
                    sheet_name = self.get_sheet_name(file)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                except Exception as e:
                    print("Error reading file :", e)


try:

    bot = SynergyBot(teardown=True)
    bot.land_first_page()
    bot.fill_login_form()
    bot.submit_login_form()
    bot.select_team_tag(text="Team Offensive")
    bot.select_team_tag(text="Team Defensive")
    time.sleep(10)
    print("MERGING FILES")
    bot.merge_data()
except Exception as a:
    print(a)
