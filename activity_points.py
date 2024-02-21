from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from bs4 import BeautifulSoup

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
activityPoints = activityPoints.sort_values(by='SemCode')

combined_df = []
start_row = 0


excel_file = 'activity_points.xlsx'
with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    # S1CU, S2CU, S3CU, S4CU, S4CU, S5CU, S6CU, S7CU, S8CU = 0 
    for index, classCode in enumerate(classCodes):
        combined_df.clear()

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
            soup = BeautifulSoup(table_html, "html.parser")
            df = pd.read_html(str(soup), header=0)[0].iloc[:-1]
            df = df.dropna(how='all').reset_index(drop=True)
            certificate_links = [f"https://www.rajagiritech.ac.in/stud/ktu/Student/{a['href']}" for a in soup.find_all('a', href=True)]
            df['Certificate (File Size<500kb)'] = certificate_links
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

        start_row = 0
        sheet_name = activityPoints.loc[index, 'Semester']
        
        for df in combined_df:
            df.to_excel(writer,sheet_name=sheet_name, startrow=start_row, index=False)
            start_row += len(df) + 2

    activityPoints = activityPoints.sort_values(by='SemCode').reset_index(drop=True)
    activityPoints.drop('SemCode', axis=1, inplace=True)

    start_row=0
    activityPoints.to_excel(writer,sheet_name='Statistics', startrow=start_row, index=False)

    total = activityPoints['Points'].sum()
    result = f"{total} "

    df = pd.DataFrame({'Total': [result]})
    start_row += len(activityPoints) + 3
    df.to_excel(writer,sheet_name='Statistics', index=False, startrow=start_row)

time.sleep(2)
wb = load_workbook(filename=excel_file)
for sheetname in wb.sheetnames:
    ws = wb[sheetname]
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='left')

wb.save(excel_file)
