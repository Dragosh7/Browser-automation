import os
from pathlib import Path

import requests
import time
from robocorp import browser
from robocorp.tasks import task
from RPA.Excel.Files import Files as Excel
from datetime import datetime


FILE_NAME = "challenge.xlsx"
EXCEL_URL = f"https://rpachallenge.com/assets/downloadFiles/{FILE_NAME}"
OUTPUT_DIR = Path(os.getenv("ROBOT_ARTIFACTS", "output"))

EXCEL_NAME = "price_tracking.xlsx"
excel_file_path = OUTPUT_DIR / EXCEL_NAME



@task
def solve_challenge():
    """
    Main task which solves the RPA challenge!

    Downloads the source data Excel file and uses Playwright to fill the entries inside
    rpachallenge.com.
    """
    browser.configure(
        browser_engine="chromium", 
        screenshot="only-on-failure", 
        headless=True 
    )
    try:
        # Reads a table from an Excel file hosted online.
        excel_file = download_file(
            EXCEL_URL, target_dir=OUTPUT_DIR, target_filename=FILE_NAME
        )
        excel = Excel()
        excel.open_workbook(excel_file)
        rows = excel.read_worksheet_as_table("Sheet1", header=True)

        # Surf the automation challenge website and fill in information from the table
        #  extracted above.
        page = browser.goto("https://rpachallenge.com/")
        page.click("button:text('Start')")
        for row in rows:
            fill_and_submit_form(row, page=page)
        element = page.locator("css=div.congratulations")
        browser.screenshot(element)
    finally:
        # A place for teardown and cleanups. (Playwright handles browser closing)
        print("Automation finished!")


def download_file(url: str, *, target_dir: Path, target_filename: str) -> Path:
    """
    Downloads a file from the given URL into a custom folder & name.

    Args:
        url: The target URL from which we'll download the file.
        target_dir: The destination directory in which we'll place the file.
        target_filename: The local file name inside which the content gets saved.

    Returns:
        Path: A Path object pointing to the downloaded file.
    """
    # Obtain the content of the file hosted online.
    response = requests.get(url)
    response.raise_for_status()  # this will raise an exception if the request fails
    # Write the content of the request response to the target file.
    target_dir.mkdir(exist_ok=True)
    local_file = target_dir / target_filename
    local_file.write_bytes(response.content)
    return local_file


def fill_and_submit_form(row: dict, *, page: browser.Page):
    """
    Fills a single form with the information of a single row from the table.

    Args:
        row: One row from the generated table out of the input Excel file.
        page: The page object over which the browser interactions are done.
    """
    field_data_map = {
        "labelFirstName": "First Name",
        "labelLastName": "Last Name",
        "labelCompanyName": "Company Name",
        "labelRole": "Role in Company",
        "labelAddress": "Address",
        "labelEmail": "Email",
        "labelPhone": "Phone Number",
    }
    for field, key in field_data_map.items():
        page.fill(f"//input[@ng-reflect-name='{field}']", str(row[key]))
    page.click("input:text('Submit')")

def record_price(title, price):
    excel = Excel()

    if excel_file_path.exists():
        excel.open_workbook(str(excel_file_path))

        # Find the next available column starting from B1 that matches today's date
        column = 2
        current_date = datetime.now().strftime("%d-%m-%Y")
        while True:
            cell_value = excel.get_cell_value(row=1, column=column)
            if cell_value is None or cell_value == current_date:
                break
            column += 1

        # Set today's date in the header if it's not already there
        if excel.get_cell_value(row=1, column=column) != current_date:
            excel.set_cell_value(row=1, column=column, value=current_date)

        # Find the next available row starting from A2 that matches the product title
        row = 2
        while True:
            cell_value = excel.get_cell_value(row=row, column=1)
            if cell_value is None or cell_value == title:
                break
            row += 1

        # Set the product title in column A if it's not already there
        if excel.get_cell_value(row=row, column=1) != title:
            excel.set_cell_value(row=row, column=1, value=title)

        # Set the price in the correct column
        excel.set_cell_value(row=row, column=column, value=price)

        # Save and close the workbook
        excel.save_workbook(str(excel_file_path))
        excel.close_workbook()

        print(f"Data has been written to {excel_file_path}.")
    else:
        print(f"Excel file not found at {excel_file_path}. Check the path.")

def update_lowest_price():
    excel = Excel()
    
    if excel_file_path.exists():
        excel.open_workbook(str(excel_file_path))

        # Iterate through rows to calculate the lowest price for each product
        row = 2  # Start from the second row since the first row has headers
        while True:
            product_name = excel.get_cell_value(row=row, column=1)
            if not product_name:  # Stop if the product name is empty
                break

            # Initialize lowest price for the current row
            lowest_price = float('inf')
            
            # Iterate over the price columns starting from column 3
            col = 3
            while True:
                cell_value = excel.get_cell_value(row=row, column=col)
                curr_date = excel.get_cell_value(row=1,column=col)
                if curr_date is None:  # Stop when no more dates available
                    break
                
                if cell_value is not None:
                    if "Lei" in cell_value:
                        try:
                            # Remove non-numeric characters and convert to float
                            price = float(cell_value.replace("Lei", "").replace(".", "").replace(",", ".").strip())
                            if price < lowest_price:
                                lowest_price = price
                        except (ValueError, TypeError):
                            pass  # Ignore cells that cannot be converted to a float

                col += 1

                # Update the "Lowest Price" column with the lowest price found for this row
            if lowest_price != float('inf'):  # Only update if a valid lowest price is found
                excel.set_cell_value(row=row, column=2, value=f"{lowest_price:,.2f} Lei")

            row += 1

        # Save and close the workbook
        excel.save_workbook(str(excel_file_path))
        excel.close_workbook()

        print("Lowest prices have been updated for each product.")
    else:
        print(f"Excel file not found at {excel_file_path}. Check the path.")


@task
def daily_price():
    browser.configure(
        browser_engine="chromium", 
        screenshot="only-on-failure", 
        headless=False 
    )

    page = browser.goto("https://www.emag.ro/")

    try:
        page.click('button:has-text("Accept toate")')
    except:
        pass 

    page.wait_for_timeout(2000)

    search_term = "iphone 15 pink 256GB"
    search_words = search_term.lower().split()

    page.wait_for_selector("#searchboxTrigger")
    src = page.locator("#searchboxTrigger")
    src.fill(search_term)
    page.keyboard.press('Enter')

    while (1):
        page.wait_for_selector("#card_grid")
        results = page.locator("#card_grid .card-item")
        product_found = False
        price = "-"
        title='-'
        for i in range(results.count()):
            card = results.nth(i).locator('h2.card-v2-title-wrapper a')
            title = card.inner_text()
            print(title)

            if all(word in title.lower() for word in search_words):
            #if search_term.lower() in title.lower():
                print("available")
                product_found=True
                price = results.nth(i).locator("p.product-new-price").inner_text()
                print(price)
                break

        if product_found:
            print(f"Product found")
            print(f"Waiting for the next day's price ...")
            record_price(title,price)
            update_lowest_price()
            break
        else:
            print("No matching results found.")
            print(f"Waiting for the next day ...")
            time.sleep(3600)
            page.goto("https://www.emag.ro/")
