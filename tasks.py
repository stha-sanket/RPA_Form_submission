from RPA.Browser.Selenium import Selenium
from robocorp import browser
import pandas as pd
from robocorp.tasks import task
import time

@task
def fill_google_form_from_excel():
    """
    Fills a Google Form with data from an Excel file using Robocorp.
    """
    browser_lib = Selenium()
    df = pd.read_excel('data.xlsx')
    browser.configure(slowmo=50)
    form_url = 'https://forms.gle/qMFvYEouXZFPCtDB7'
    # Open the Google Form
    browser_lib.open_browser(form_url, browser='chrome')
    time.sleep(3)

    for index, row in df.iterrows():
        # Fill out the form fields (same as before)
        name_label_xpath = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input'
        name_container_locator = f"xpath:{name_label_xpath}/ancestor::div[@class='Qr7Oae']"
        name_field_locator = f"{name_container_locator}//input[@type='text']"
        browser_lib.clear_element_text(name_field_locator)
        browser_lib.input_text(name_field_locator, row['Name'])
        
        # Address Field
        address_label_xpath = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea'
        address_container_locator = f"xpath:{address_label_xpath}/ancestor::div[@class='Qr7Oae']"
        address_field_locator = f"{address_container_locator}//textarea"
        browser_lib.clear_element_text(address_field_locator)
        browser_lib.input_text(address_field_locator, row['Address'])
        
        # Phone Number Field
        phone_label_xpath = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input'
        phone_container_locator = f"xpath:{phone_label_xpath}/ancestor::div[@class='Qr7Oae']"
        phone_field_locator = f"{phone_container_locator}//input[@type='text']"
        browser_lib.clear_element_text(phone_field_locator)
        browser_lib.input_text(phone_field_locator, str(row['Phone Number']))
        
        # Highschool Field
        highschool_label_xpath = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input' 
        highschool_container_locator = f"xpath:{highschool_label_xpath}/ancestor::div[@class='Qr7Oae']" 
        highschool_field_locator = f"{highschool_container_locator}//input[@type='text']"
        browser_lib.clear_element_text(highschool_field_locator)
        browser_lib.input_text(highschool_field_locator, row['Highschool'])
        
        # GPA Field
        gpa_label_xpath = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input'
        gpa_container_locator = f"xpath:{gpa_label_xpath}/ancestor::div[@class='Qr7Oae']" 
        gpa_field_locator = f"{gpa_container_locator}//input[@type='text']"
        browser_lib.clear_element_text(gpa_field_locator)
        browser_lib.input_text(gpa_field_locator, str(row['Gpa']))

        # Select course radio button
        course = row['Course Interested']
        if course == 'Mechanical Engineering':
            course_radio_button_locator = '//*[@id="i34"]/div[3]/div'
            browser_lib.click_element(course_radio_button_locator)
        elif course == 'Computer Science':
            course_radio_button_locator = '//*[@id="i31"]/div[3]/div'
            browser_lib.click_element(course_radio_button_locator)
        elif course == 'Mechanical Science':
            course_radio_button_locator = '//*[@id="i37"]/div[3]/div'
            browser_lib.click_element(course_radio_button_locator)
        elif course == 'Data Science':
            course_radio_button_locator = '//*[@id="i40"]/div[3]/div'
            browser_lib.click_element(course_radio_button_locator)
        else:
            course_radio_button_locator = '//*[@id="i40"]/div[3]/div'
            browser_lib.click_element(course_radio_button_locator)
        # Submit the form
        submit_button_locator = '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div'
        browser_lib.click_element(submit_button_locator)
        time.sleep(2)  # Short wait to ensure form submission is processed

        # Wait and click the "Submit Another Response" link
        submit_another_button_locator = '//div[@class="c2gzEf"]/a'
        browser_lib.wait_until_element_is_visible(submit_another_button_locator, timeout=30)
        browser_lib.click_element(submit_another_button_locator)

        print(f"Form submitted for row {index + 1}")

    browser_lib.close_browser()
    print("Automation finished.")
