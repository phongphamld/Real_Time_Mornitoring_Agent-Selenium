import time
import re
import os
from selenium.webdriver.common.by import By
from main import ByteHiCrawler

# pip install pywin32
import win32clipboard


class realtime_mornitoring(ByteHiCrawler):
    def hide_browser(self):
        #self.options.add_argument("--headless")
        self.options.add_argument("--window-size=1920x1080")

    def go_to_montoring_page(self):

        # Run browser in background
        self.hide_browser()


        self.initialize_driver()
        time.sleep(10)

        self.driver.set_window_size(1920,1080)

        # go to field management
        self.driver.get('Mornitoring_web page')
        time.sleep(4)

        # go to field
        self.click_by_class_text('eU0sx', 'All Mornitor Status')

        # Expand 100 rows
        self.click_by_class_text('united_helpdesk_field_management_i18n-select-selection','items / page')
        self.click_by_class_text('united_helpdesk_field_management_i18n-select-option', '100 items')


    
    def arlet_condition(self, tier, status, duration):
        if tier == 'TL':
            return False
        elif (status in ['Busy', 'Break']) and (int(duration[3:5]) > 14):
            return True
        elif status == 'Abnormal':
            return True
        elif (status == 'Lunch') and (int(duration[0:2]) > 0):
            return True
        else:
            return False
    
    def screenshot_agent(self, object):

        try:
            # Record time stamp and set name for capture
            timestamp = int(time.time())
            screenshot_filename = f'Screen shot\\element_screenshot_{timestamp}.png'

            counter = 1
            while os.path.exists(screenshot_filename):
                screenshot_filename = f'Screen shot\\element_screenshot_{timestamp}_{counter}.png'
                counter += 1

            element_screenshot = object.screenshot_as_png
            with open(screenshot_filename, 'wb') as f:
                f.write(element_screenshot)
        except:
            print('error when export')
# try
            # convert image to bit
        bit_image = self.convert_image_to_bit(screenshot_filename)   

            # copy image to clipboard
        self.send_to_clipboard(win32clipboard.CF_DIB, bit_image)
  

    def rta(self):
        # go to page
        self.go_to_montoring_page()

        # open new message tab
        self.open_new_message_tab()

        # get the table and row
        session_num = 1
        agents_status = {}

        while True:
                print("RTA Session : " + str(session_num) )
                print(agents_status)
                rows = self.driver.find_elements(By.CLASS_NAME, 'united_helpdesk_field_management_i18n-table-row')
                for row in rows:

                    try: 
                        groups = re.split(r'\n', row.text)
                        name = groups[0]
                        tier = groups[1]
                        status = groups[3]
                        duration = groups[-1]

                        check = self.arlet_condition(tier,status,duration)
                        #print(name +' | '+  status +' | '+  duration[3:5] +'|' + str(check))

                        if check and (name not in agents_status or session_num - agents_status[name] >= 60):
                            #self.screenshot_agent(name, tier, status, duration, row)
                            agents_status[name] = session_num
                            
                            #Screeen shot Agent and Copy to clipboard
                            self.screenshot_agent(row)

                            # Read alias of egent
                            alias = self.read_alias(name)[0]
                            task = self.read_alias(name)[1]

                            # Send message
                            self.switch_tab_and_paste_clipboard(alias,task ,status, name, duration)

                    except:
                        self.driver.switch_to.window(self.driver.window_handles[0])
                        pass

                time.sleep(5)
                session_num+=1
if __name__ == '__main__':
    realtime_mornitoring().rta()


    

