from selenium import webdriver
from bs4 import BeautifulSoup
import os
import time
import zipfile
import pandas as pd
from selenium.webdriver.common.by import By
import re
import win32com.client as win32
from PIL import ImageGrab
import sys
import shutil




created_folders = set()

def create_folder(folder_name):
    """Create a folder if it does not exist."""
    if folder_name not in created_folders:
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            created_folders.add(folder_name)

def create_workbook():
    """
    Initializes a new Excel workbook using the COM interface.
    Returns:
        workbook: A new Excel workbook object.
    """
    os.system('taskill /f /in EXCEL.exe')
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True 

    workbook = excel.Workbooks.Add()
    workbook.Activate()
    time.sleep(2)
    
    return workbook

def write_data_to_sheet(filtered_data, workbook):
    """
    Writes the given DataFrame to a new sheet in the specified Excel workbook
    with improved formatting.
    Returns:
        data_sheet: The newly created and formatted Excel sheet.
    """
    data_sheet = workbook.Sheets.Add()
    data_sheet.Name = 'Data'

    # Set headers with bold text and center alignment
    header_range = data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(1, len(filtered_data.columns)))
    header_range.Value = [filtered_data.columns.tolist()]
    header_range.Font.Bold = True
    header_range.HorizontalAlignment = -4108  # Center alignment

    # Write the data values
    for row_num, row in enumerate(filtered_data.values, 2):
        for col_num, value in enumerate(row, 1):
            data_sheet.Cells(row_num, col_num).Value = value

    # AutoFit columns for better readability
    data_sheet.Columns.AutoFit()

    # Apply table formatting (banded rows, borders)
    last_row = len(filtered_data) + 1
    last_col = len(filtered_data.columns)
    table_range = data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(last_row, last_col))
    
    table_range.Borders.LineStyle = 1  # Add borders
    table_range.Borders.Weight = 2  # Medium weight

    # Apply light gray banding for readability
    table_range.Interior.Color = 0xF0F0F0  # Light Gray Background

    return data_sheet



def create_pivot(data_sheet, workbook):
    """
    Creates a Pivot Table in a new sheet named "Pivot" using data from the given worksheet.
    Returns:
        pivot_sheet: The newly created worksheet containing the pivot table.
    """
    pivot_sheet = workbook.Sheets.Add()
    pivot_sheet.Name = 'Pivot'

    # Define the data range for the pivot table
    last_row = data_sheet.Cells(data_sheet.Rows.Count, 1).End(-4162).Row
    last_col = data_sheet.Cells(1, data_sheet.Columns.Count).End(-4159).Column
    data_range = data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(last_row, last_col))

    pivot_cache = workbook.PivotCaches().Create(1, data_range)
    pivot_table = pivot_cache.CreatePivotTable(pivot_sheet.Cells(1, 1), "EmployeePivot")

    # Rows (Business Unit, Department)
    pivot_table.PivotFields("Business Unit").Orientation = 1  # Row
    pivot_table.PivotFields("Department").Orientation = 1  # Row

    # Columns (Gender, Ethnicity)
    pivot_table.PivotFields("Gender").Orientation = 2  # Filter
    pivot_table.PivotFields("Ethnicity").Orientation = 2  # Filter

    # Values (Average Annual Salary)
    pivot_field = pivot_table.PivotFields("Annual Salary")
    pivot_field.Orientation = 4  # Data
    pivot_field.Function = -4106  # Average function
    pivot_field.NumberFormat = '#,##0'  # Number formatting
    pivot_field.Position = 1

    # Refresh the pivot table to apply changes
    pivot_table.RefreshTable()

    return pivot_sheet


def screenshot_table(pivot_sheet, workbook):
    '''
    Takes a screenshot of the pivot table and saves it as an image.'''
    workbook.Activate()
    time.sleep(2)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    excel.ActiveWindow.WindowState = 3  # 1 corresponds to maximizing the window

    time.sleep(1)

    # Get the range of the pivot table dynamically
    pivot_table_range = pivot_sheet.UsedRange
    start_row = pivot_table_range.Row
    end_row = start_row + pivot_table_range.Rows.Count - 1
    end_col = pivot_table_range.Columns.Count
    picture_range = pivot_sheet.Range(pivot_sheet.Cells(start_row, 1), pivot_sheet.Cells(end_row, end_col))

    picture_range.CopyPicture(Appearance=1, Format=2)
    time.sleep(2)

    # Crop the region of the pivot table based on its coordinates
    ImageGrab.grabclipboard().save("c:/Akshaya_scripts/python_projects/PivotTable_Data_Insights.png")

    print("Screenshot saved successfully.")

    workbook.SaveAs(r"C:\Akshaya_scripts\python_projects\test\Employee_Data_Summary.xlsx")
    workbook.Close(SaveChanges = True)
    excel.Quit()


def get_latest_zip(downloads_folder, timeout=10):
    """Find the most recent ZIP file in the Downloads folder, waiting for it to complete downloading."""
    print(f"Checking {downloads_folder} for the latest ZIP file...")

    end_time = time.time() + timeout
    zip_file_path = None

    while time.time() < end_time:
        time.sleep(10)  
        zip_files = [os.path.join(downloads_folder, f) for f in os.listdir(downloads_folder) if f.endswith(".zip")]
        if zip_files:
            time.sleep(10) 
            zip_file_path = max(zip_files, key=os.path.getctime)  # Get the most recently created ZIP
            if os.path.exists(zip_file_path) and not zip_file_path.endswith(".crdownload"):  
                zip_filename = os.path.basename(zip_file_path)  
                print(f"Latest ZIP file found: {zip_filename}")
                return zip_file_path, zip_filename
            else:
                print("Waiting for ZIP file to finish downloading...")
                time.sleep(10)  
        
    print("No ZIP file found in Downloads after timeout.")
    return None, None


def extract_zip_file(zip_file_path, extract_to):
    """Extracts a ZIP file to a given location and finds the first .xlsx file inside."""
    print(f"Extracting ZIP file: {os.path.basename(zip_file_path)} to {extract_to}")

    try:
        os.makedirs(extract_to, exist_ok=True)

        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
            extracted_files = zip_ref.namelist()

        print(f"Extracted files: {extracted_files}")
        return extracted_files

    except Exception as e:
        print("An error occurred while extracting:", e)
        return None


def get_xslx(extracted_files, extract_to):
    '''Find the first .xlsx file in the extracted files.'''
    for file in extracted_files:
        if file.endswith(".xlsx"):
            xlsx_file_path = os.path.join(extract_to, file)

            print(f"Found Excel file: {xlsx_file_path}")
            return xlsx_file_path
    
    print("No .xlsx file found in the ZIP.")
    return None

def clean_data(df):
    """
    Data Cleaning Steps:
    1. Remove duplicate records to ensure data consistency.
    2. Replace null values with "N/A" where applicable.
    3. Standardize Job Title spelling (case-insensitive corrections).
    4. Convert Salary to a numeric format by removing "$" and ",", ensuring missing values are treated as NaN.
    5. Convert Bonus % to numeric by stripping "%", while handling null and non-string values gracefully.
    6. Localize timestamp fields to UTC for consistency across time zones.
    Returns:
        Cleaned DataFrame
    """

    df = df.drop_duplicates()
    df.loc[:, ~df.columns.isin(["Exit Date", "Hire Date", "Annual Salary"])] = df.loc[:, ~df.columns.isin(["Exit Date", "Hire Date", "Annual Salary"])].fillna('N/A')

    df["Job Title"] = df["Job Title"].replace("Sr. Manger", "Sr. Manager", regex=True)

    df["Annual Salary"] = df["Annual Salary"].astype(str).str.replace(r'[\$,]', '', regex=True)
    df["Annual Salary"] = pd.to_numeric(df["Annual Salary"], errors='coerce')

    df["Bonus %"] = df["Bonus %"].replace('%', '').astype(float).multiply(100)

    
    df["Hire Date"] = df["Hire Date"].dt.tz_localize('UTC')

    return df

def filter_data(df):
    '''Filter data to retain information of employees still at the company and under the Age of 60 y/o.'''
    filtered_data = df[(df['Age'] < 60) & (df['Exit Date'].isna())]
    dropped_df = filtered_data.drop(columns="Exit Date")
    return dropped_df

def send_mail():
    """
    Sends an email with the Excel file and screenshot attached. 
    Validates file existence, creates an email with HTML body, 
    attaches the files, and sends the email using Outlook.
    """

    excel_file_path = "c:/Akshaya_scripts/python_projects/test/Employee_Data_Summary.xlsx"
    screenshot_path = "c:/Akshaya_scripts/python_projects/PivotTable_Data_Insights.png"

    if not os.path.exists(excel_file_path) or not os.path.exists(screenshot_path):
        print("One or both files do not exist.")
    else:
        outlook = win32.gencache.EnsureDispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        mail = outlook.CreateItem(0)   # Create a new email item
        
        mail.To = 'adam.kapadia@aristocrat.com'
        mail.Subject = 'Pivot Table Analysis Report - Employee Dataset'

        body = """
        <html>
        <body>
            <p>Hey Adam,</p>

            <p>I hope this email finds you well.</p>

            <p>Attached, you will find the Excel file containing the clean data, pivot table analysis and the screenshot of the pivot table for your reference.</p>

            <p>The script used to automate this process is available at the following GitHub link: <a href="https://github.com/akshayasiddi/EmployeeDataScraper.git">GitHub Repository</a>.</p>

            <h3>Overview:</h3>
                <p>The data provides a comprehensive breakdown of average annual salaries by gender, ethnicity, and job category. It also highlights the salary trends within specific departments.</p>
                
                <h3>Key Insights:</h3>
                <ul>
                <li><strong>Gender Pay Gap:</strong> 
                    <ul>
                    <li>On average, males earn $8,308 more than females across all departments and ethnicities, indicating a clear gender pay gap.</li>
                    </ul>
                </li>
                
                <li><strong>Ethnic Salary Discrepancies:</strong> 
                    <ul>
                    <li><strong>Asian employees:</strong>Asian males have the highest average salary at $121,587 (females: $116,544, males: $121,587).</li>
                    <li><strong>Black employees:</strong>Black male employees earn significantly less than their female counterparts by about $21,275 (females: $120,730, males: $99,455).</li>
                    <li><strong>Caucasian employees:</strong>Salaries for Caucasian employees are slightly higher for males than females, with a $5,378 difference  ($108,547) and male salaries ($113,925).</li>
                    <li><strong>Latino employees:</strong>Latino employees' salaries are nearly identical across genders, with a small difference (females: $111,718, males: $111,157).</li>
                    </ul>
                </li>
                
                <li><strong>Department-Specific Insights:</strong>
                    <ul>
                    <li><strong>Corporate:</strong> Balanced distribution of salaries across genders.</li>
                    <li><strong>Manufacturing:</strong> The lowest salaries are found in this department.</li>
                    <li><strong>Research & Development:</strong> Females earn significantly lower than males in this department.</li>
                    <li><strong>Speciality Products:</strong> The highest salaries are in this department, with a balanced gender distribution.</li>
                    </ul>
                </li>
                </ul>

            <p>Thank you for your time, and I look forward to your feedback!</p>

            <p>Best regards,</p>
            <p><strong>Akshaya</strong></p>
        </body>
        </html>
        """
        
        mail.HTMLBody = body
        
        mail.Attachments.Add(excel_file_path)
        mail.Attachments.Add(screenshot_path)
        mail.SentOnBehalfOfName = "siddisa@mail.uc.edu"
        mail.Send()

        print(f"Email sent successfully with attachments.")


def send_error_email(error_message):
    outlook = win32.gencache.EnsureDispatch('Outlook.Application')
    namespace = outlook.GetNamespace("MAPI")
    mail = outlook.CreateItem(0)  
    
    mail.To = 'siddisa@mail.uc.edu'  
    mail.Subject = 'Error in Pivot Table Analysis Report - Employee Dataset'

    body = f"""
    <html>
    <body>

        <p>There was an error in the script execution. Below is the error message:</p>

        <p><strong>Error Message:</strong></p>
        <p>{error_message}</p>

    </body>
    </html>
    """
    
    mail.HTMLBody = body
    mail.SentOnBehalfOfName = "siddisa@mail.uc.edu"
    mail.Send()
    print(f"Error email sent successfully.")

def scrape_page(driver, root_folder):
    """Navigates to the page, downloads the ZIP file, and extracts the data."""
    main_url = "https://www.thespreadsheetguru.com/sample-data/"
    driver.get(main_url)
    time.sleep(5)

    # Locate the download button and click it 
    ele = driver.find_element(By.XPATH, '/html/body/section/div/div[1]/article/div[6]/div/div/a')
    driver.execute_script("arguments[0].click();", ele)  # Execute JavaScript click

    # Define the Downloads folder path
    downloads_folder = os.path.expanduser("~/Downloads")

    zip_file_path, zip_filename = get_latest_zip(downloads_folder)

    extracted_files = extract_zip_file(zip_file_path, root_folder)

    driver.quit() 

    xslx_file = get_xslx(extracted_files, root_folder)

    # Read the Excel file using Pandas
    df = pd.read_excel(xslx_file) 
    cleaned_data = clean_data(df)
    filtered_data = filter_data(cleaned_data)

    workbook = create_workbook()
    data_sheet = write_data_to_sheet(filtered_data, workbook)
    pivot_sheet = create_pivot(data_sheet, workbook)
    screenshot_table(pivot_sheet, workbook)

    # Send success mail
    send_mail()


# Main Program
if __name__ == '__main__':
    root_folder = "c:/Akshaya_scripts/python_projects"

    driver = webdriver.Chrome()

    # Create the root folder for the entire scraping operation
    create_folder(root_folder)

    retries = 2  # Total number of retries
    attempt = 0
    while attempt <= retries:
        try:
            attempt += 1

            # Run the automation
            scrape_page(driver, root_folder)
            break

        except  Exception as e:
            if attempt == retries + 1:
                print(f"Failed after {retries+1} attempts: {str(e)}") 
                send_error_email(str(e))
                break 
            else:
                print(f"Attempt {attempt} failed. Retrying...")  # Retry message
                time.sleep(1)  # Optional sleep before retrying
