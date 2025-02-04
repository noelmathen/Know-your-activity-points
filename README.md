### Activity Points Calculator for RSMS (Rajagiri Student Management System)

This Python script automates the process of fetching and calculating activity points for students using the Rajagiri Student Management System (RSMS). It utilizes Selenium to log in to the RSMS portal, extract data about uploaded certificates, approval status, and overall activity points, and then organizes this information into an Excel spreadsheet.

#### Features:

- **Login Automation**: Seamlessly logs in to the RSMS portal using the provided user ID and password.
  
- **Data Extraction**: Retrieves data about uploaded certificates, including their approval status and other details, from different semesters.
  
- **Activity Points Calculation**: Calculates the activity points for each semester and the overall total points.

- **Excel Output**: Organizes the extracted data into an Excel spreadsheet with separate sheets for each semester and a summary sheet for statistics.

- **Online Execution**: The script can be run online using the Colab notebook provided (`activity_points.ipynb`). It allows users to execute the script on any device, including mobile phones, without requiring local Python installation.

#### Requirements:

- Python 3.x
- Selenium
- pandas
- openpyxl
- BeautifulSoup

#### Usage:

1. Install the required dependencies:

   ```bash
   pip install selenium pandas openpyxl beautifulsoup4
   ```

2. Clone the repository:

   ```bash
   git clone https://github.com/your-username/your-repo.git
   ```

3. Navigate to the project directory:

   ```bash
   cd activity-points-calculator
   ```

4. Run the script:

   ```bash
   python activity_points_calculator.py
   ```

5. Enter your RSMS user ID and password when prompted.

#### Online Execution:

- The script can be executed online using the provided Colab notebook: [activity_points.ipynb](https://colab.research.google.com/github/noelmathen/Know-your-activity-points/blob/main/activity_points.ipynb).
  
- Users can access and run the notebook on any device, including mobile phones, by opening it in Google Colab.

#### Excel Output:

The generated Excel file (`activity_points.xlsx`) will contain the following sheets:

1. **Semester-wise Sheets**: Each semester's activity data, including certificates, approval status, and ratings by faculty members.
   
2. **Statistics Sheet**: Summary of activity points for each semester and the overall total points.

#### Note:

- Ensure that the Chrome WebDriver is installed and its path is set in the system environment variables.
  
- Before running the script, make sure to customize the `URL`, `Userid`, and `Password` variables according to your RSMS credentials and RSMS URL.

This script provides a convenient way for Rajagiri students to track their activity points and certificate statuses efficiently. Feel free to contribute to the project or suggest improvements!
