import os
import time
import re
import shutil
import csv
# pip install pywin32
import win32clipboard

# pip install Pillow
from PIL import Image
from io import BytesIO

import json
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains



class ByteHiCrawler:
    def __init__(self):
    
        self.by_method = By.ID
        self.driver = None
        self.wait = None
        self.options = webdriver.ChromeOptions()
        self.dir_path = os.getcwd()

        # Set default path and url
        self.CHROME_DRIVER_PATH = self.dir_path + "\\Web Driver\\chromedriver.exe"
        self.default_folder_path = self.dir_path + "\\temp download files"
        self.USER_DATA_DIR = self.dir_path + "\\User Data"
        self.URL = 'Your Web URL'


    def initialize_driver(self):
        service = Service(executable_path=self.CHROME_DRIVER_PATH)
        # Configure Chrome options
        self.options.add_argument(f'--user-data-dir={self.USER_DATA_DIR}')
        self.options.add_experimental_option("prefs", {
            "download.default_directory": self.default_folder_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        self.driver = webdriver.Chrome(service=service, options=self.options)
        self.wait = WebDriverWait(self.driver, 20)
        self.driver.get(self.URL)
        
    
    def click_object(self, by_selector, string):
        self.wait.until(EC.presence_of_element_located((by_selector, string)))
        obj = self.driver.find_element(by=by_selector, value=string)
        obj.click()

    def delete_all_files_in_folder(self):
        files = os.listdir(self.default_folder_path)
        for file in files:
            file_path = os.path.join(self.default_folder_path, file)
            if os.path.isfile(file_path):
                os.remove(file_path)

    def update_filename_after_download(self, new_file_name):
        files_with_same_name = [f for f in os.listdir(self.default_folder_path) if f == new_file_name]
        for file_with_same_name in files_with_same_name:
            file_path = os.path.join(self.default_folder_path, file_with_same_name)
            os.remove(file_path)
        filename = max([os.path.join(self.default_folder_path, f) for f in os.listdir(self.default_folder_path)], key=os.path.getctime)
        shutil.move(filename, os.path.join(self.default_folder_path, new_file_name))

    def click_by_class_text(self, class_name, text_input):
        elements = self.driver.find_elements(By.CLASS_NAME, class_name)
        for element in elements:
            if text_input in element.text:
                element.click()
                break
        time.sleep(5)

    def download_table_field_management(self, table_xpath, filename):
        self.click_object(By.XPATH, table_xpath)
        time.sleep(10)
        self.click_by_class_text('united_helpdesk_field_management_i18n-button-content', 'Export')
        time.sleep(5)
        self.update_filename_after_download(filename)

    def select_today(self, left_picker_xpath):
        left_picker = self.driver.find_element(By.CLASS_NAME, left_picker_xpath)
        days_left_picker = left_picker.find_elements(By.XPATH, './/div')
        current_date = datetime.now()
        formatted_date = current_date.strftime('%Y-%m-%d')
        for day in days_left_picker:
            title = day.get_attribute('title')
            if title == formatted_date:
                day.click()
                time.sleep(2)
                day.click()
                time.sleep(2)

    def get_ticket_number(self):
        time.sleep(3)
        row_number = self.driver.find_element(By.XPATH, '//*[@id="root"]/section/main/div/div/div/div/div[2]/div[3]/div/div/div/div/div[2]/span[1]').text
        if row_number != '':
            ticket_number = int(re.findall(r'\d+', row_number)[-1])
            return ticket_number
        else:
            return 0

    def choose_role(self, role):

        # Expand avatar
        self.click_object(By.XPATH, '//*[@id="united-helpdesk-main"]/div[1]/div[2]/div/div[4]')
        time.sleep(2)

        # click to expand business role
        self.click_by_class_text('united-dropdown-item','Business Line')

        # click to role
        self.click_by_class_text('_24XffIbzNnI6gUxkHzUjGg', role)

    def open_message_page(self):
        # click on bell Message
        self.click_object(By.CLASS_NAME, 'notify-sdk-bell-wrapper')

        # click on view all
        self.click_by_class_text('united-button-content', 'View All')
    def extract_last_number(self, str):
        pattern = r'\b\d+\b'
        try:
            last_number = int(re.findall(pattern, str)[-1])
        except:
            last_number = 0
        return last_number


    def open_download_data_lark_links(self, link, name):
        
        self.driver.get(link)

        # Click more
        self.click_object(By.CLASS_NAME, 'suite-more-menu')
        time.sleep(3)
        # Click download
        self.click_by_class_text('navigation-bar__moreMenu_v3-item__text', 'Download As')

        # Click download
        self.click_by_class_text('ud__menu-normal-item-group-list', 'Excel (.xlsx)')

        time.sleep(10)

        self.update_filename_after_download(name)

    def click_inprogress_ticket(self):

        # click ticket
        self.click_object(By.XPATH, '//*[@id="united-helpdesk-main"]/div[1]/div[1]/div[2]/div/div[6]/a/div')
        time.sleep(5)

        # click inprogress tickets in my group
        self.click_by_class_text('line--DWSbm','In-progress' )
        time.sleep(5)        

    def send_to_clipboard(self, clip_type, data):

         # clip_type for picture is win32clipboard.CF_DIB
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()
    

    def convert_image_to_bit(self,screenshot_filename):
        image = Image.open(screenshot_filename)

        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:] 
        output.close()
        return data
    
    def open_new_message_tab(self):

        # open new tab
        self.driver.execute_script("window.open('about:blank','secondtab')")
        time.sleep(2)

        # change driver to new tab
        self.driver.switch_to.window(self.driver.window_handles[1])

        # open message page
        self.driver.get('Your Lark Chat Message Page URL')
        time.sleep(10)

        # Click on bot group
        self.click_by_class_text('feed-shortcut-item', 'Group Chat Name')


        # Change driver to bytehi tab
        self.driver.switch_to.window(self.driver.window_handles[0])
        time.sleep(2)

    def switch_tab_and_paste_clipboard(self, alias = 'j' , task ='' ,status = 'Break', name = '', duration = '' ):
        
        # swtich to chat tab
        self.driver.switch_to.window(self.driver.window_handles[1])
        time.sleep(3)

        # click on expand message
        self.click_object(By.CLASS_NAME, 'open-inner-editor')
        time.sleep(2)

        # select typing
        self.click_object(By.CLASS_NAME, 'post-edit-zone')
        time.sleep(4)

        # paste clipboard
        action = ActionChains(self.driver)
        action.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        time.sleep(2)

        # Tag teamlead
            # Tag 
        if task in ['Tier2', 'Seller Call Back']:
            for alias in ['3200527', '3230228']:
                action.send_keys('@').perform()
                time.sleep(3)
                action.send_keys(alias).perform()
                time.sleep(5)
                action.key_down(Keys.ENTER).perform()
                time.sleep(2)
        else:
            action.send_keys('@').perform()
            time.sleep(3)

            action.send_keys(alias).perform()
            time.sleep(5)
            action.key_down(Keys.ENTER).perform()
            time.sleep(1)

        # Attach message to team lead
        action.send_keys('\nPlease assist to check Agent status as ' + status + "\nAgent name: " + name + "\nDuration: " + duration).perform()
        time.sleep(3)

        # click send
        self.click_by_class_text('uni-btn', 'Send')

        self.driver.switch_to.window(self.driver.window_handles[0])


    # Function read csv alias
    def read_alias(self, name):
        with open('alias_leader.csv','r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            headers = next(reader)

            data = []
            for row in reader:
                record = {}
                for i, value in enumerate(row):
                    record[headers[i]] = value
                data.append(record)

        search_name = name

        email = None

        for entry in data:
            if entry['Customer service name'] == search_name:
                alias = entry['Line Manager Alias']
                task = entry['Task']
        return [alias, task]
    

