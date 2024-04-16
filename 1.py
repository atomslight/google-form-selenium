from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import subprocess
# Specify the path to your WebDriver (e.g., chromedriver.exe)
#webdriver_path = 'path/to/chromedriver.exe'

# Initialize the WebDriver (e.g., Chrome)
driver = webdriver.Chrome()
# Open the Google Form
driver.get("https://docs.google.com/forms/d/1rylDfkTsnN6-_aCBADJ6PAC77_OmDHDs67r2-UyUiR8/viewform?pli=1&pli=1&edit_requested=true")
# Initialize a variable to store the previously inserted name
previous_name = None
# Initialize a variable to store the last entered number
last_entered_number = 2400  # Initialize it with 6 to start with 7

def funct_repeat():
    time.sleep(6)
    global previous_name, last_entered_number  # Declare global variables

    # Read data from Excel file using pandas
    excel_file_path = "C:\\Users\\PC\\Documents\\malenames.xlsx"
    df = pd.read_excel(excel_file_path)

    # Define a WebDriverWait with a timeout of 10 seconds
    wait = WebDriverWait(driver, 10)

    # Iterate through each name in the column
    for name in df['Name']:
        # Check if the current name is different from the previous name
        if name != previous_name:
            # Remove the current name from the DataFrame
            df = df[df['Name'] != name]
            # Save the updated DataFrame back to the Excel file
            df.to_excel(excel_file_path, index=False)
            # Find the input element for the name field by its HTML attribute (inspect the form element)
            # For example, find input elements by name or other attributes
            StId_element = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input")))

            # Create a list to store the numbers that have been sent
            sent_numbers = []

            # Input the next number into the input field
            last_entered_number += 1
            StId_element.clear()
            StId_element.send_keys(str(last_entered_number))

            # Rest of your form filling logic goes here
            name_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input")))  # Replace with the actual field name

            # Fill in the form field with the current name
            name_field.send_keys(name)
            driver.execute_script("window.scrollBy(0, window.innerHeight);")
            #secondary_check=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[3]/div/div/div[2]/div[1]/div[3]/label/div/div[1]/div[2]"))).click()
            primary_check=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/form/div[2]/div/div[2]/div[3]/div/div/div[2]/div[1]/div[1]/label/div/div[1]/div[2]"))).click()
            school_id=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div[2]/textarea")))
            school_id.send_keys("22110417909")
            # Submit the form (you may need to adapt this to your form structure)
            #submit_button = driver.find_element_by_xpath('//input[@type="submit"]')
            #submit_button.click()
            male_check=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[5]/div/div/div[2]/div[1]/div[1]/label/div/div[1]/div[2]")))
            male_check.click()
            drop_yes=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[6]/div/div/div[2]/div[1]/div[1]/label/div/div[1]/div[2]")))
            drop_yes.click()
            #drop_no=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/form/div[2]/div/div[2]/div[6]/div/div/div[2]/div[1]/div[2]/label/div/div[1]/div[2]"))).click()
            #session_click2018=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[7]/div/div/div[2]/div[1]/div[5]/label/div/div[1]/div[2]"))).click()
            #session_click2022=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/form/div[2]/div/div[2]/div[7]/div/div/div[2]/div[1]/div[1]/label/div/div[1]/div[2]"))).click()
            ession_click2020=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[7]/div/div/div[2]/div[1]/div[3]/label/div/div[1]/div[2]"))).click()
            #session_click=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[7]/div/div/div[2]/div[1]/div[2]/label/div/div[1]/div[2]")))
            #session_click.click()
            #admin_check=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[8]/div/div/div[2]/div[1]/div[4]/label/div/div[1]/div[2]"))).click()
            #family_check=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[8]/div/div/div[2]/div[1]/div[2]/label/div/div[1]/div[2]"))).click()
            finance_check=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[8]/div/div/div[2]/div[1]/div[1]/label/div/div[1]/div[2]"))).click()
            academic_check=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[2]/div[8]/div/div/div[2]/div[1]/div[3]/label/div/div[1]/div[2]"))).click()
            # Submit the form (you may need to adapt this to your form structure)
            #submit_button = driver.find_element_by_xpath('//input[@type="submit"]')
            #submit_button.click()
            # Check if the last entered number has reached 15
            if last_entered_number > 2500:
                # Execute the new script
               # try:
                #    subprocess.run(["python", "C:\\Users\\PC\\Desktop\\Project to be made\\Project IRIS\\2020fem - Copy.py"], check=True)
                #except subprocess.CalledProcessError as e:
                 #   print("Error executing the second script:", e)
                #breakpoint
                driver.quit()
            time.sleep(5)
            submit_btn=wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/div[3]/form/div[2]/div/div[3]/div[1]/div[1]/div"))).click()
            # Open the Google Form for the next submission
            driver.get("https://docs.google.com/forms/d/1rylDfkTsnN6-_aCBADJ6PAC77_OmDHDs67r2-UyUiR8/viewform?pli=1&pli=1&edit_requested=true")

            # Update the previous_name variable with the current name
            previous_name = name
            print(previous_name)
            print(last_entered_number)
# Start the repetition process
funct_repeat()
