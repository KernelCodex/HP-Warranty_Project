from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException




def interact_with_page(serial_number):

    # Remove popups
    remove_clutter()

    if wait_loading_screen():
        print("loading screen stop")
        return None

    try:
        # Enter current Serial
        input_field = WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.ID, "inputtextpfinder"))
        )
        input_field.clear()
        input_field.send_keys(serial_number)

        # Submit current device
        button_submit = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((By.ID, "FindMyProduct"))
        )
        button_submit.click()

        time.sleep(1)

        if wait_loading_screen():
            print("loading screen stop")
            return None
        if check_serial_exist():
            print("serial check stop")
            return None
        if check_requires_prod_num():
            return None

        # Load required Information
        WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.CLASS_NAME, "info-section"))
        )

        # Target html content to extract info from
        html_content = driver.page_source

        # Go back to home page
        link_back = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((By.ID, "back"))
        )
        link_back.click()

        return html_content

    except:
        print("interact page block exception stop")
        return None

def extract_info(html_content):

    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        info_sections = soup.find_all('div', class_='info-section')

        serial_info = {}
        coverage_types = set()

        

        # Loop through all sections
        for section in info_sections:
            items = section.find_all('div', class_='info-item')
            section_info = {}

            # Extract Info from current section
            for item in items:
                label = item.find('div', class_='label').text.strip()
                text = item.find('div', class_='text').text.strip()
                key = label.replace(' ', '_').replace('-', '_')
                section_info[key] = text
                coverage_types.add(section_info.get('Coverage_type', 'Unknown'))

            # Formatting Keys in dictionary
            for key, value in section_info.items():
                if key != 'Coverage_type':
                    new_key = f"{key}_{section_info.get('Coverage_type', 'Unknown')}"
                    serial_info[new_key] = value

        coverage_types_formatted = ', '.join(coverage_types)
        serial_info['CoverageTypes'] = coverage_types_formatted

        formatted_serial_info = {}
        for old_key, value in serial_info.items():
            new_key = old_key.replace(' ', '_').replace('-', '_')
            formatted_serial_info[new_key] = value
        print(formatted_serial_info)
        return formatted_serial_info
    except:
        print("extract info block exception stop")
        return None



def remove_clutter():
    try:
        driver.execute_script("document.getElementById('onetrust-consent-sdk').style.display = 'none';")
    except:
        pass
    try:
        driver.execute_script("document.getElementById('MDigitalLightboxWrapper').style.display = 'none';")
    except:
        pass
    try:
        driver.execute_script("document.body.style.overflow = 'auto';")
    except:
        pass

def wait_loading_screen():
    time.sleep(1.5)
    try:
        # Wait until the loading screen element is no longer displayed
        WebDriverWait(driver, 120).until_not(
            EC.presence_of_element_located((By.CLASS_NAME, "loading_screen_wrapper"))
        )
        return False
    except TimeoutException:
        return True

def check_serial_exist():
    try:
        serial_not_found_message = WebDriverWait(driver, 1.5).until(
            EC.visibility_of_element_located((By.XPATH, "//p[contains(@class, 'errorTitle') and contains(text(), 'We were unable to match your product based on the information provided')]"))
        )
        return True
    except:
        return False

def check_requires_prod_num():
    try:
        error_message = WebDriverWait(driver, 1.5).until(
            EC.visibility_of_element_located((By.XPATH, "//p[contains(@class, 'field-info') and contains(@class, 'errorTxt') and contains(@class, 'is-invalid-field') and contains(text(), 'This product cannot be identified using the serial number alone. Please add a product number in the field below:')]"))
        )
        return True
    except:
        return False

