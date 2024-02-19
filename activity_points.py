from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import pandas as pd

URL = "https://www.rajagiritech.ac.in/stud/ktu/Student/"
driver = webdriver.Chrome()
driver.get(URL)

driver.find_element(By.NAME, "Userid").send_keys('U2109053')
driver.find_element(By.NAME, "Password").send_keys('210825')
driver.find_element(By.XPATH, "//input[@type='submit']").click()
driver.find_element(By.LINK_TEXT, "Activity Point Form").click()

classCodeDropdown = Select(driver.find_element(By.NAME, "Class_Code"))
classCodes = [option.get_attribute("value") for option in classCodeDropdown.options]

activityPoints = pd.DataFrame(classCodes, columns=['SemCode'])
activityPoints['Semester'] = "Semester " + activityPoints['SemCode'].str[-3]
activityPoints['Points'] = [0] * len(activityPoints)
activityPoints = activityPoints.sort_values(by='SemCode').reset_index(drop=True)
print(activityPoints)

# S1CU, S2CU, S3CU, S4CU, S4CU, S5CU, S6CU, S7CU, S8CU = 0 
for index, classCode in enumerate(classCodes):
    classCodeDropdown = Select(driver.find_element(By.NAME, "Class_Code"))
    classCodeDropdown.select_by_value(classCode)

    activitiesDropdown = Select(driver.find_element(By.NAME, "ACode"))
    activities = [option.get_attribute("value") for option in activitiesDropdown.options]
    for activity in activities:
        activitiesDropdown = Select(driver.find_element(By.NAME, "ACode"))
        activitiesDropdown.select_by_value(activity)

        driver.find_element(By.XPATH, "//input[@value='Add Activity']").click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "B1")))
        try:
            table = driver.find_element(By.XPATH, "//table[@class='table table-striped']")
        except:
            print("No certificates added. moving on.......")
            continue

        table_html = table.get_attribute("outerHTML")
        df = pd.read_html(table_html, index_col=False)[0].iloc[:-1]
        column_names = df.iloc[0]  # Get the first row as column names
        df = df[1:]  # Exclude the first row from the DataFrame
        df.columns = column_names 
        print(df["Rating By Faculty"])
        print(df)
        activityPoints.loc[index+1, 'Points'] += df['Rating By Faculty'].sum()
        print(f"rbf = {df['Rating By Faculty'].sum()}")
        print(f"new points = {activityPoints.loc[index+1, 'Points']}")

        print("\n")
        print(activityPoints)

            
time.sleep(3)