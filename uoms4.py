from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import csv

# Configure options for Selenium
firefox_options = Options()

# Path to the GeckoDriver
driver_path = '/media/vaishnav/DATA/Projects/shyama_maam/geckodriver'

# Create a WebDriver instance for Firefox
service = Service(driver_path)
driver = webdriver.Firefox(service=service, options=firefox_options)

# Function to login and extract data for a given registration number
def extract_mark_details(registration_no, dob):
    try:
        # Open the result page
        driver.get('https://results.uomexam.com')  # Replace with the actual URL if different

        # Wait until the registration number input field is present
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'MainContent_txtRollNo'))
        )

        # Locate the registration number input field and enter the value
        reg_no_field = driver.find_element(By.ID, 'MainContent_txtRollNo')
        reg_no_field.clear()
        reg_no_field.send_keys(registration_no)

        # Locate the date of birth input field and enter the provided DOB
        dob_field = driver.find_element(By.ID, 'MainContent_txtDoB')
        dob_field.clear()
        dob_field.send_keys(dob)

        # Try closing the date picker or obscuring elements (if any)
        try:
            date_picker_icon = driver.find_element(By.CLASS_NAME, "ui-icon-circle-triangle-e")
            if date_picker_icon.is_displayed():
                date_picker_icon.click()
                time.sleep(1)  # Wait for it to close
        except:
            pass

        # Scroll into view to ensure visibility
        submit_button = driver.find_element(By.ID, 'MainContent_btnSubmit')
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)

        # Wait until the submit button is clickable
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'MainContent_btnSubmit'))
        )

        # Use JavaScript to click the button (to bypass any overlapping issues)
        driver.execute_script("arguments[0].click();", submit_button)

        # Wait for the result page to load
        time.sleep(3)

        # Extract the result details from the result page
        exam_name = driver.find_element(By.XPATH, '//*[@id="MainContent_ddlExamBatch"]').text
        sem_name = driver.find_element(By.XPATH, '//*[@id="MainContent_ddlSemester"]').text
        student_name = driver.find_element(By.XPATH, '//*[@id="MainContent_lblStudentName"]').text
        subject_marks = driver.find_element(By.XPATH, '//*[@id="MainContent_gvStudentMarks2"]').text
        sem_marks = driver.find_element(By.XPATH, '//*[@id="MainContent_lblResultSummary"]').text

        # Extract subject marks as a list
        marks_data = []
        subject_rows = driver.find_elements(By.XPATH, '//*[@id="MainContent_gvStudentMarks2"]/tbody/tr')
        for row in subject_rows:
            columns = row.find_elements(By.TAG_NAME, 'td')
            marks_data.append([col.text for col in columns])

        # Return all extracted data
        return {
            "Student Name": student_name,
            "Exam": exam_name,
            "Semester": sem_name,
            "CGPA": sem_marks,
            "Marks": marks_data
        }

    except Exception as e:
        print(f"Error occurred for {registration_no}: {e}")
        return None

# Save data to CSV function
def save_to_csv(data, filename):
    # Define the CSV structure
    headers = ['Student Name', 'Exam', 'Semester', 'CGPA', 'Subject', 'Course Type', 'C1', 'C2', 'C3', 'Marks Secured', 'Credits', 'Grade', 'Remarks']

    # Open the CSV file for writing
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file)
        
        # Write the header
        writer.writerow(headers)
        
        # Write the data
        for entry in data:
            student_info = [entry["Student Name"], entry["Exam"], entry["Semester"], entry["CGPA"]]
            
            # First write the student information row without subject details
            writer.writerow(student_info + [''] * (len(headers) - len(student_info)))

            # Write each subject's marks with the same student information
            for marks_row in entry["Marks"]:
                writer.writerow(student_info + marks_row)


# Iterate through registration numbers and extract details
all_data = []
for reg_number in range(190001, 190003):
    registration_no = f"DP{reg_number}"
    dob = "01/01/2000"  # Use a random date of birth
    result = extract_mark_details(registration_no, dob)
    if result:
        all_data.append(result)

# Save all extracted data to a CSV file
save_to_csv(all_data, 'exam_results.csv')

# Close the WebDriver after task completion
driver.quit()

