import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Load Excel data
df = pd.read_excel('data.xlsx')

form_url = 'https://forms.gle/DjRAeagg8ZWFMhPGA'  # **REPLACE WITH YOUR FORM URL**

# Open the Google Form
driver.get(form_url)

# Wait until the form is loaded and the first field container is present
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//form[@id='mG61Hd']"))) # Wait for the form element

for index, row in df.iterrows():
    try:  # Added try-except block for error handling

        # **Finding Field Containers - More Robust Approach**
        form_element = driver.find_element(By.XPATH, "//form[@id='mG61Hd']") # Find the form once and use it as context

        # Name Field - Find by label text (if label is stable) and then relative input
        name_label_xpath = '/html/body/div/div[2]/form/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input' # Adjust label class if needed
        name_container = form_element.find_element(By.XPATH, name_label_xpath + "/ancestor::div[@class='Qr7Oae']") # Find parent container
        name_field = name_container.find_element(By.XPATH, ".//input[@type='text']") # Find input within container

        name_field.clear()
        name_field.send_keys(row['Name'])

        # Address Field - Find by label text (if label is stable) and then relative textarea
        address_label_xpath = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea'
        address_container = form_element.find_element(By.XPATH, address_label_xpath + "/ancestor::div[@class='Qr7Oae']") # Find parent container
        address_field = address_container.find_element(By.XPATH, ".//textarea") # Find textarea within container
        address_field.clear()
        address_field.send_keys(row['Address'])

        # Phone Number Field - Find by label text (if label is stable) and then relative input
        phone_label_xpath = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input'
        phone_container = form_element.find_element(By.XPATH, phone_label_xpath + "/ancestor::div[@class='Qr7Oae']") # Find parent container
        phone_field = phone_container.find_element(By.XPATH, ".//input[@type='text']") # Find input within container
        phone_field.clear()
        phone_field.send_keys(str(row['Phone Number']))

        # Highschool Field - Find by label text (if label is stable) and then relative input
        highschool_label_xpath = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input'
        highschool_container = form_element.find_element(By.XPATH, highschool_label_xpath + "/ancestor::div[@class='Qr7Oae']") # Find parent container
        highschool_field = highschool_container.find_element(By.XPATH, ".//input[@type='text']") # Find input within container
        highschool_field.clear()
        highschool_field.send_keys(row['Highschool'])

        # GPA Field - Find by label text (if label is stable) and then relative input
        gpa_label_xpath = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input'
        gpa_container = form_element.find_element(By.XPATH, gpa_label_xpath + "/ancestor::div[@class='Qr7Oae']") # Find parent container
        gpa_field = gpa_container.find_element(By.XPATH, ".//input[@type='text']") # Find input within container
        gpa_field.clear()
        gpa_field.send_keys(str(row['Gpa']))

        # **Radio Button Example (Adapt based on your form's structure)**
        # course = row['Course Interested']
        # if course == 'Mechanical Engineering':
        #     course_label_xpath = ".//div[contains(text(), 'Mechanical Engineering') and @class='...']" # **REPLACE with actual label class if needed**
        #     course_radio_button = form_element.find_element(By.XPATH, course_label_xpath + "/ancestor::label[@class='...']/div[@role='radio']") # **REPLACE with actual label and radio classes**
        #     course_radio_button.click()
        # elif ... (similar for other courses)


        # Submit the form - More Robust Submit Button XPath (text-based)
        submit_button = form_element.find_element(By.XPATH, ".//span[text()='Submit']")
        submit_button.click()

        # Wait for the confirmation page to load (using class name - could also use text)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'freebirdFormviewerViewResponseConfirmation')))

        # Click "Submit another response" - More Robust XPath (text-based link)
        try:
            submit_another_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Submit another response')]"))
            )
            submit_another_button.click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//form[@id='mG61Hd']"))) # Wait for form reload
        except:
            print("No 'Submit another response' button found (possibly last entry).")

        print(f"Form submitted for row {index + 1}")

    except Exception as e:
        print(f"Error processing row {index + 1}: {e}")
        break

# Close the driver
driver.quit()
print("Automation finished.")