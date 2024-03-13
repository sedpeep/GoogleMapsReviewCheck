import datetime
import sys

current_time = datetime.datetime.now()
expiration_date = current_time + datetime.timedelta(days=2)
if(current_time >expiration_date):
    print("The script has expired")
    sys.exit()


import importlib
import subprocess
try:
    importlib.import_module('selenium')
except ImportError:
    print("Selenium is not installed. Installing ...")
    subprocess.check_call(['pip','install','selenium'])

from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time


options = Options()
options.add_experimental_option("detach", True)
options.add_argument("--disable-popup-blocking")
driver = webdriver.Edge(options=options)

# Function to find the corresponding link in 'links.xlsx'
def find_link(link_number, file_path='links.xlsx'):
    df = pd.read_excel(file_path)
    link = df.loc[df['Number'] == int(link_number), 'Link'].values
    if len(link) > 0:
        return link[0]
    else:
        return None

# Function to check if the username exists in the reviews
def search_for_username(link, username):
    try:

        reviews_panel = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[3]/div[8]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[3]/div/div/button[2]/div[2]/div[2]')))
        reviews_panel.click()


        search_button_css = '#QA0Szd > div > div > div.w6VYqd > div.bJzME.tTVLSc > div > div.e07Vkf.kA9KIf > div > div > div.m6QErb.DxyBCb.kA9KIf.dS8AEf > div.m6QErb.Pf6ghf.KoSBEe.ecceSd.tLjsW > div.i7mKJb.fontBodyMedium > div.m3rned > div.pV4rW.q8YqMd > div > button > span > img'
        search_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, search_button_css)))
        search_button.click()


        reviews_container_css = '#QA0Szd > div > div > div.w6VYqd > div.bJzME.tTVLSc > div > div.e07Vkf.kA9KIf > div > div > div.m6QErb.DxyBCb.kA9KIf.dS8AEf'
        reviews_container = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, reviews_container_css)))

        user_found = False
        elements_texts_seen = set()
        scroll_pause_time = 1.5  # value can be changed

        # Scroll and search
        while True:
            # Scroll down the page incrementally
            driver.execute_script("arguments[0].scrollTop += 800", reviews_container)
            time.sleep(scroll_pause_time)

            elements = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class,'d4r55')]")))
            new_elements = [element for element in elements if element.text.lower() not in elements_texts_seen]
            # checking username in reviews
            for element in new_elements:
                element_text = element.text.lower()
                if element_text not in elements_texts_seen:
                    elements_texts_seen.add(element_text)
                    #print(f"Element text:{element_text}")
                    if element_text in username.lower():
                        print(f"User Found: {username}")
                        user_found = True # Return immediately after finding the user
                        break

            if user_found:
                return True

            # checking if end of page reached
            current_scroll_position = driver.execute_script("return arguments[0].scrollTop", reviews_container)
            max_scroll_position = driver.execute_script("return arguments[0].scrollHeight", reviews_container) - \
                                  reviews_container.size['height']

            #print(f"Current Scroll: {current_scroll_position}, Max Scroll: {max_scroll_position}")

            if current_scroll_position+10 >= max_scroll_position:
                print("End of scrollable area reached")
                break

        print("User not found after reaching the bottom")
        return False

    except TimeoutException as e:
        print(f"Timeout occurred while waiting for elements on the page: {str(e)}")
        return False
    except Exception as e:
        import traceback
        print(f"An unexpected error occurred: {traceback.format_exc()}")
        return False


# Read 'list.xlsx'
df_list = pd.read_excel('list.xlsx')

# Loop through each row in 'list.xlsx'
for index, row in df_list.iterrows():
    if pd.isna(row['Status']):
        print(f"Skipping row {index + 1} due to NaN values in the Status column.")
        continue

    # Extract the link number and username from the Status column
    status_parts = row['Status'].split(maxsplit=1)
    if len(status_parts) == 2 and status_parts[0].startswith('#'):
        link_number = status_parts[0].lstrip('#')
        username_to_search = status_parts[1]
        print(link_number)
        print(username_to_search)

        # Find the corresponding link from 'links.xlsx'
        link = find_link(link_number)

        if link:
            driver.get(link)
            time.sleep(3)
            user_found = search_for_username(link, username_to_search)

            # Update the status in 'list.xlsx' DataFrame
            if user_found:
                status_update = f"✔️ (checked - {datetime.now().strftime('%Y-%m-%d')})"
            else:
                status_update = f"❌ (checked - {datetime.now().strftime('%Y-%m-%d')})"

            df_list.at[index, 'Status'] = status_parts[0] + " " + username_to_search + " " + status_update
        else:
            print(f"Link not found in 'links.xlsx' for link number: {link_number}")
    else:
        print(f"Invalid or missing link number and user information in 'list.xlsx' at row {index + 1}")

# Save the updated 'list.xlsx' DataFrame
df_list.to_excel('list.xlsx', index=False)
driver.quit()
