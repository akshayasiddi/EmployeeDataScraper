from selenium import webdriver
from bs4 import BeautifulSoup
import os
import time
import zipfile
import pandas as pd
from selenium.webdriver.common.by import By
import re
import win32com.client as win32
import pyautogui


created_folders = set()

def create_folder(folder_name):
    """Create a folder if it does not exist."""
    if folder_name not in created_folders:
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            created_folders.add(folder_name)

def create_workbook():
    # Initialize COM interface to Excel
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True  # Set to False if you want to run in the background

    # Create a new workbook
    workbook = excel.Workbooks.Add()
    
    return workbook

def write_data_to_sheet(filtered_data, workbook):

    data_sheet = workbook.Sheets.Add()
    data_sheet.Name = 'Data'

    # Write the DataFrame to the "Data" sheet, including all rows and columns
    for i, column in enumerate(filtered_data.columns, 1):
        data_sheet.Cells(1, i).Value = column  # Column headers

    # Write data rows starting from the second row
    for row_num, row in enumerate(filtered_data.values, 2):
        for col_num, value in enumerate(row, 1):
            data_sheet.Cells(row_num, col_num).Value = value  # Data values

    # Optional: Adjust column width for better readability
    for i in range(1, len(filtered_data.columns) + 1):
        data_sheet.Columns(i).AutoFit()

    return data_sheet



def create_pivot(data_sheet, filtered_data, workbook):
    # Create a "Pivot" sheet for the pivot table
    pivot_sheet = workbook.Sheets.Add()
    pivot_sheet.Name = 'Pivot'

    data_range = data_sheet.Range("A1").CurrentRegion  # Get the range that includes all data

    # Add a Pivot Table
    pivot_table_range = pivot_sheet.Range("A1")
    pivot_table = pivot_sheet.PivotTableWizard(
    SourceType=1,  # Use range as source
    SourceData=data_range,  # Define the range for pivot table data
    TableDestination=pivot_table_range  # Set the destination of the pivot table
    )

    # Set the row, column, and data fields after the pivot table is created
    pivot_table.PivotFields("Department").Orientation = 1  # Row field: Department
    pivot_table.PivotFields("Job Title").Orientation = 2  # Column field: Job Title
    pivot_table.PivotFields("Annual Salary").Orientation = 4  # Data field: Annual Salary (Sum)
    pivot_table.PivotFields("Annual Salary").Function = 4  # -4157 is the constant for Sum aggregation

    # Refresh the pivot table to apply changes
    pivot_table.RefreshTable()
    # Step 1: Ensure only the pivot table is visible
    # Adjust the range to show only the pivot table (remove any blank spaces or unnecessary rows/columns)
    pivot_sheet.PageSetup.PrintArea = pivot_sheet.Range("A1").CurrentRegion.Address

    # Optional: Resize the columns/rows to fit the pivot table nicely
    pivot_sheet.Cells.EntireColumn.AutoFit()
    pivot_sheet.Cells.EntireRow.AutoFit()

    # Step 2: Bring the Excel window to the foreground
    workbook.WindowState = 2  # Maximize Excel window
    workbook.Activate()  # Bring the Excel window to the foreground

    # Step 3: Allow time for Excel to adjust and render everything properly
    time.sleep(2)  # Sleep to ensure that Excel has finished rendering

    # Step 4: Take the screenshot of the entire screen (containing the pivot table)
    screenshot = pyautogui.screenshot()

    # Step 5: Save the screenshot to a file
    screenshot.save("c:/Akshaya_scripts/python_projects/pivot_table_screenshot.png")


def get_latest_zip(downloads_folder, timeout=10):
    """Find the most recent ZIP file in the Downloads folder, waiting for it to complete downloading."""
    print(f"Checking {downloads_folder} for the latest ZIP file...")

    end_time = time.time() + timeout
    zip_file_path = None

    while time.time() < end_time:
        time.sleep(10)  # Wait and check again
        zip_files = [os.path.join(downloads_folder, f) for f in os.listdir(downloads_folder) if f.endswith(".zip")]
        if zip_files:
            time.sleep(10)  # Wait and check again
            zip_file_path = max(zip_files, key=os.path.getctime)  # Get the most recently created ZIP
            if os.path.exists(zip_file_path) and not zip_file_path.endswith(".crdownload"):  # Ensure it's fully downloaded
                zip_filename = os.path.basename(zip_file_path)  # Extract just the file name
                print(f"Latest ZIP file found: {zip_filename}")
                return zip_file_path, zip_filename
            else:
                print("Waiting for ZIP file to finish downloading...")
                time.sleep(10)  # Wait and check again
        
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
    # Find the first .xlsx file in the extracted files
        for file in extracted_files:
            if file.endswith(".xlsx"):
                xlsx_file_path = os.path.join(extract_to, file)

                print(f"Found Excel file: {xlsx_file_path}")
                return xlsx_file_path
        
        print("No .xlsx file found in the ZIP.")
        return None

def clean_data(df):

    # Remove Duplicates
    df = df.drop_duplicates()

    # Handle null values by replacing them with 'N/A'
    df.loc[:, ~df.columns.isin(["Exit Date", "Hire Date", "Annual Salary"])] = df.loc[:, ~df.columns.isin(["Exit Date", "Hire Date", "Annual Salary"])].fillna('N/A')

    # Fix Job Title spelling (case-insensitive)
    df["Job Title"] = df["Job Title"].replace("Sr. Manger", "Sr. Manager", regex=True)

    # Convert Date Format (DD-MM-YYYY → YYYY-MM-DD), handling invalid dates
    #df["Hire Date"] = pd.to_datetime(df["Hire Date"], format="%d-%m-%Y", errors='coerce').dt.strftime("%Y-%m-%d")

    # Convert Salary to Numeric (Removing '$' and ','), handling nulls
    df["Annual Salary"] = df["Annual Salary"].apply(lambda x: int(re.sub(r'[\$,]', '', str(x))) if pd.notnull(x) else None)
    df["Annual Salary"] = pd.to_numeric(df["Annual Salary"], errors='coerce')  # Coerce errors into NaN


    # Convert Bonus % to Numeric (Removing '%'), handling nulls and non-string values
    df["Bonus %"] = df["Bonus %"].replace('%', '').astype(float).multiply(100)

    return df

def filter_data(df):

    filtered_data = df[(df['Age'] < 60) & (df['Exit Date'].notnull())]
    return filtered_data

def scrape_page(driver, root_folder):
    """Navigates to the page, downloads the ZIP file, and extracts the data."""
    main_url = "https://www.thespreadsheetguru.com/sample-data/"
    driver.get(main_url)
    time.sleep(5)

    try:
        # Locate the download button and click it 
        ele = driver.find_element(By.XPATH, '/html/body/section/div/div[1]/article/div[6]/div/div/a')
        driver.execute_script("arguments[0].click();", ele)  # Execute JavaScript click

        # Define the Downloads folder path
        downloads_folder = os.path.expanduser("~/Downloads")

        # Wait for the latest ZIP file to appear and get its name
        zip_file_path, zip_filename = get_latest_zip(downloads_folder)

        # Extract ZIP file and process contents
        extracted_files = extract_zip_file(zip_file_path, root_folder)

        xslx_file = get_xslx(extracted_files, root_folder)

        # Read the Excel file using Pandas
        df = pd.read_excel(xslx_file) 
        cleaned_data = clean_data(df)
        filtered_data = filter_data(cleaned_data)
        print(filtered_data)


        datetime_columns = filtered_data.select_dtypes(include=['datetime']).columns

        for col in datetime_columns:
            # Localize each datetime column to UTC
            if filtered_data[col].dt.tz is None:
                filtered_data[col] = filtered_data[col].dt.tz_localize('UTC')

        workbook = create_workbook()
        data_sheet = write_data_to_sheet(filtered_data, workbook)
        create_pivot(data_sheet, filtered_data, workbook)
        workbook.SaveAs(r'c:\Akshaya_scripts\python_projects\test\test.xlsx')



    except Exception as e:
        print("An error occurred:", e)

# Main Program
if __name__ == '__main__':
    root_folder = "c:/Akshaya_scripts/python_projects"

    driver = webdriver.Chrome()

    # Create the root folder for the entire scraping operation
    create_folder(root_folder)

    # Run the scraping function
    scrape_page(driver, root_folder)

    driver.quit()
