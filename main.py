import sys
from PyQt5.QtWidgets import QApplication, QWidget, QHBoxLayout, QVBoxLayout, QLabel, QLineEdit, QPushButton
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup

#"C:\Users\ruanr\OneDrive\Documents\HP Warranty\Serials.xlsx"
#self.general_label.setText('Button 1 was clicked!')

class MyWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Set Window Title and icon
        self.setWindowTitle(' ')
        icon_path = './assets/icon.png'
        icon_pixmap = QPixmap(icon_path)
        self.setWindowIcon(QIcon(icon_pixmap))

        # Create the main vertical layout
        main_layout = QVBoxLayout()

        # Create labels
        self.general_label = QLabel('HP Warranty Check')

        # Create text field
        self.path_text_field = QLineEdit()
        self.path_text_field.setPlaceholderText('Enter Excel File Path')
        self.path_text_field.setText('"C:\\Users\\ruanr\\OneDrive\\Documents\\HP Warranty\\Serials.xlsx"')

        # Create horizontal layout for buttons
        button_layout = QHBoxLayout()

        # Create buttons
        upload_btn = QPushButton('Upload')
        self.download_btn = QPushButton('Download')

        # Add buttons to the horizontal layout
        button_layout.addWidget(upload_btn)
        button_layout.addWidget(self.download_btn)

        # Add items to the main layout
        main_layout.addWidget(self.general_label)
        main_layout.addWidget(self.path_text_field)
        main_layout.addLayout(button_layout)

        # Set the main layout for the window
        self.setLayout(main_layout)

        # Automatically adjust window size to fit contents
        self.adjustSize()

        # Lock window sizing
        self.setFixedSize(self.size())

        # Set window to always stay on top
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)

        upload_btn.clicked.connect(self.upload_clicked)
        self.download_btn.clicked.connect(self.download_clicked)

    def upload_clicked(self):

        file_path = self.path_text_field.text().strip('"\'')

        try:
            # Initialize Browser
            self.initialize_selenium()
            composite_dict = {}
            machine_dict = {}
            failed_serial = {}
            serial_count = 0
            serial_total = 0

            # Read Excel file into DataFrame
            df = pd.read_excel(file_path, sheet_name='Serial')

            # Extract data from 'Serial' column and store in a list
            serial_data = df['Serial'].tolist()

            # Total len of serials
            serial_total = len(serial_data)

            for serial_number in serial_data:
                info_dict = {}
                content = self.interact_with_page(serial_number)
                if content == "Timed out":
                    info_dict = {serial_number: "Timed out"}
                    failed_serial.update(info_dict)
                    driver.quit()
                    self.initialize_selenium()
                    continue
                elif content == "Serial unable to match":
                    info_dict = {serial_number: "Serial unable to match"}
                    failed_serial.update(info_dict)
                    continue
                elif content == "Serial requires prod num":
                    info_dict = {serial_number: "Serial requires prod num"}
                    failed_serial.update(info_dict)
                    continue
                elif content == "Warranty sections appear blank":
                    info_dict = {serial_number: "Warranty sections appear blank"}
                    failed_serial.update(info_dict)
                    continue
                else:
                    soup = BeautifulSoup(content, 'html.parser')

                    machine_info_dict = {}
                    product_number = ""
                    page_link = ""
                    product_name = ""

                    # Get product info div
                    product_info_div = soup.find(class_='product-info-text')
                    if product_info_div:
                        # Get product name from h2 tag
                        product_name_elem = product_info_div.find('h2', class_='ng-tns-c75-0')
                        if product_name_elem:
                            product_name = product_name_elem.string.strip()

                        # Get product number
                        product_number_elem = product_info_div.find('p', string='Product: ')
                        if product_number_elem:
                            product_number_span = product_number_elem.find_next('span', class_='ng-tns-c75-0')
                            if product_number_span:
                                product_number = product_number_span.string.strip()

                    # Get page link
                    page_link_elem = soup.find(id="Support_visitMyProductPage").find('a')
                    if page_link_elem:
                        page_link = "https://support.hp.com" + page_link_elem['href']

                    machine_info_dict = {
                        "serial": serial_number,
                        "product_number": product_number,
                        "product_name": product_name,
                        "page_link": page_link,
                    }
                    machine_dict.update(machine_info_dict)

                    info_sections = soup.find_all(class_='info-section')
                    for section in info_sections:
                        info_dict = {}
                        coverage_type = ""
                        service_type = ""
                        start_date = ""
                        end_date = ""
                        service_level = ""
                        deliverables = ""

                        # Extract relevant information from each info section
                        coverage_type_elem = section.find(class_='label', string='Coverage type')
                        coverage_type = coverage_type_elem.find_next(class_='text').string.strip() if coverage_type_elem else ""

                        service_type_elem = section.find(class_='label', string='Service type')
                        service_type = service_type_elem.find_next(class_='text').string.strip() if service_type_elem else ""

                        start_date_elem = section.find(class_='label', string='Start date')
                        start_date = start_date_elem.find_next(class_='text').string.strip() if start_date_elem else ""

                        end_date_elem = section.find(class_='label', string='End date')
                        end_date = end_date_elem.find_next(class_='text').string.strip() if end_date_elem else ""

                        service_level_elem = section.find(class_='label', string='Service level')
                        service_level = ', '.join(p.string.strip() for p in service_level_elem.find_next(class_='text').find_all('p')) if service_level_elem else ""

                        deliverables_elem = section.find(class_='label', string='Deliverables')
                        deliverables = ', '.join(p.string.strip() for p in deliverables_elem.find_next(class_='text').find_all('p')) if deliverables_elem else ""

                        # Create info dict for the current info section
                        info_dict = {
                            'serial': serial_number,
                            'product_number': product_number,
                            "product_name": product_name,
                            'page_link': page_link,
                            'coverage_type': coverage_type,
                            'service_type': service_type,
                            'start_date': start_date,
                            'end_date': end_date,
                            'service_level': service_level,
                            'deliverables': deliverables
                        }

                        # Update composite_dict with info_dict
                        composite_dict.update(info_dict)

            print(machine_dict)
            print(composite_dict)
            print(failed_serial)

        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"Unhandled Exception: {e}")
            self.general_label.setText(f"An Unhandled exception has occurred: {e}")


    def download_clicked(self):
        pass

    def interact_with_page(self, serial_number):

        self.remove_clutter()

        # Check Loading screen
        if self.wait_loading_screen():
            print(f"Timed out for serial: {serial_number}")
            return "Timed out"

        # Enter current Serial
        input_field = WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.ID, "inputtextpfinder"))
        )
        self.remove_clutter()
        input_field.clear()
        self.remove_clutter()
        input_field.send_keys(serial_number)

        # Submit current device
        button_submit = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((By.ID, "FindMyProduct"))
        )
        self.remove_clutter()
        button_submit.click()
        time.sleep(2)

        # Check Loading screen
        if self.wait_loading_screen():
            print(f"Timed out for serial: {serial_number}")
            return "Timed out"

        if self.check_serial_exist():
            print(f"Serial can't be matched: {serial_number}")
            return "Serial unable to match"

        if self.check_requires_prod_num():
            print(f"Serial requires prod num: {serial_number}")
            return "Serial requires prod num"

        # Load required Information
        try:
            WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((By.CLASS_NAME, "info-section"))
            )
        except:
            print(f"Warranty sections appear blank: {serial_number}")
            link_back = WebDriverWait(driver, 120).until(
                EC.element_to_be_clickable((By.ID, "back"))
            )
            self.remove_clutter()
            link_back.click()
            return "Warranty sections appear blank"

        # Target html content to extract info from
        html_content = driver.page_source

        # Go back to home page
        link_back = WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((By.ID, "back"))
        )
        self.remove_clutter()
        link_back.click()

        return html_content

    def initialize_selenium(self):
        global driver
        url = "https://support.hp.com/za-en/check-warranty"
        options = Options()
        options.add_experimental_option("detach", True)
        driver = webdriver.Chrome(options=options) 
        driver.get(url)
        driver.maximize_window()

    def remove_clutter(self):
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

    def wait_loading_screen(self):
        time.sleep(1.5)
        try:
            # Wait until the loading screen element is no longer displayed
            WebDriverWait(driver, 120).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "loading_screen_wrapper"))
            )
            return False #Loading screen disappeared
        except TimeoutException:
            return True #Loading screen failed to complete

    def check_serial_exist(self):
        try:
            element_locator = (By.CSS_SELECTOR, "p.errorTitle")
            error_message = "We were unable to match your product based on the information provided"
            serial_not_found_message = WebDriverWait(driver, 1.5).until(
                EC.visibility_of_element_located(element_locator)
            )
            return True
        except:
            return False #If block not found

    def check_requires_prod_num(self):
        try:
            element_locator = (By.CSS_SELECTOR, "p.field-info.errorTxt.is-invalid-field")
            error_message = "This product cannot be identified using the serial number alone. Please add a product number in the field below:"
            error_message = WebDriverWait(driver, 1.5).until(
                EC.visibility_of_element_located(element_locator)
            )
            return True
        except:
            return False #If block not found

# Create the application instance
app = QApplication(sys.argv)

# Create the main window instance
window = MyWindow()

# Show the main window
window.show()

# Execute the application
sys.exit(app.exec_())
