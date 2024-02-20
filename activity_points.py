from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import pandas as pd
import numpy as np

URL = "https://www.rajagiritech.ac.in/stud/ktu/Student/"
driver = webdriver.Chrome()
driver.get(URL)

driver.find_element(By.NAME, "Userid").send_keys('U2109064')
driver.find_element(By.NAME, "Password").send_keys('211022')
driver.find_element(By.XPATH, "//input[@type='submit']").click()
driver.find_element(By.LINK_TEXT, "Activity Point Form").click()

classCodeDropdown = Select(driver.find_element(By.NAME, "Class_Code"))
classCodes = [option.get_attribute("value") for option in classCodeDropdown.options]

activityPoints = pd.DataFrame(classCodes, columns=['SemCode'])
activityPoints['Semester'] = "Semester " + activityPoints['SemCode'].str[-3]
activityPoints['Points'] = [0] * len(activityPoints)
activityPoints = activityPoints.sort_values(by='SemCode')
print(activityPoints)

combined_df = []

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
        df = pd.read_html(table_html, header=0)[0].iloc[:-1]
        df = df.dropna(how='all').reset_index(drop=True)
        df['Rating By Faculty'] = df['Rating By Faculty'].replace(np.nan, 0)
        df.index += 1
        activityPoints.loc[index, 'Points'] += df['Rating By Faculty'].sum()

        combined_df.append(df)
        print(df)
        print("\n")
        print(combined_df)
        print("\n")
        print(activityPoints)
        print("\n")


activityPoints = activityPoints.sort_values(by='SemCode').reset_index(drop=True)
activityPoints.drop('SemCode', axis=1, inplace=True)

start_row = 0
excel_file = 'activity_points.xlsx'
with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    for df in combined_df:
        df.to_excel(writer, startrow=start_row, index=False)
        start_row += len(df) + 2

    start_row+=2
    activityPoints.to_excel(writer, startrow=start_row, index=False)

    total = activityPoints['Points'].sum()
    result = f"{total} "

    df = pd.DataFrame({'Total': [result]})
    start_row += len(activityPoints) + 3
    df.to_excel(writer, index=False, startrow=start_row)

time.sleep(2)
