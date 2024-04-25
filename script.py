from selenium import webdriver
from selenium.common import exceptions as seleniumException
from selenium.webdriver.common.by import By
import openpyxl
from time import sleep


def get_case_links():
    base_url = "https://stateofgreen.com/en/solution-providers/"
    driver.get(base_url)

    # if there is a "more" button then click it
    while True:
        if driver.find_element(By.CLASS_NAME, "button-down").text != "":
            sleep(0.3)
            downBtn = driver.find_element(By.CLASS_NAME, "button-down")
            downBtn.click()
        else:
            break

    partner_items = driver.find_elements(By.CLASS_NAME, "partners-item")
    for item in partner_items:
        case_links.append(item.get_attribute("href"))


def save_to_excel():
    file_name = "stateofgreen-data.xlsx"
    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active

    # Headers
    col_names = [
        "company name",
        "company description",
        "specialisations",
        "specialisation colors",
        "CVR",
        "contact name",
        "contact email",
        "contact phone",
        "company website",
        "provider link",
    ]

    sheet[f"A1"] = col_names[0].capitalize()
    sheet[f"B1"] = col_names[1].capitalize()
    sheet[f"C1"] = col_names[2].capitalize()
    sheet[f"D1"] = col_names[3].capitalize()
    sheet[f"E1"] = col_names[4].capitalize()
    sheet[f"F1"] = col_names[5].capitalize()
    sheet[f"G1"] = col_names[6].capitalize()
    sheet[f"H1"] = col_names[7].capitalize()
    sheet[f"I1"] = col_names[8].capitalize()
    sheet[f"J1"] = col_names[9].capitalize()

    data_start_row = 2
    for index, case in enumerate(cases):
        sheet[f"A{index+data_start_row}"] = case["company_name"]
        sheet[f"B{index+data_start_row}"] = case["company_description"]
        sheet[f"C{index+data_start_row}"] = case["specialisations"]
        sheet[f"D{index+data_start_row}"] = case["colors"]
        sheet[f"E{index+data_start_row}"] = case["cvr"]
        sheet[f"F{index+data_start_row}"] = case["contact_name"]

        if len(case["contact_email"]) > 1:
            sheet[f"G{index+data_start_row}"].value = case["contact_email"]
            sheet[f"G{index+data_start_row}"].hyperlink = (
                f"mailto:{case['contact_email']}"
            )
            sheet[f"G{index+data_start_row}"].style = "Hyperlink"

        sheet[f"H{index+data_start_row}"] = case["contact_phone"]

        if len(case["website"]) > 1:
            sheet[f"I{index+data_start_row}"].value = case["website"]
            sheet[f"I{index+data_start_row}"].hyperlink = case["website"]
            sheet[f"I{index+data_start_row}"].style = "Hyperlink"

        sheet[f"J{index+data_start_row}"].value = case["link"]
        sheet[f"J{index+data_start_row}"].hyperlink = case["link"]
        sheet[f"J{index+data_start_row}"].style = "Hyperlink"

    wb.save(filename=file_name)
    wb.close()


def get_company_name():
    company_name = driver.find_element(By.TAG_NAME, "h1").text
    return company_name


def get_company_description():
    try:
        entry_elem = driver.find_element(By.CLASS_NAME, "entry")
        description = entry_elem.find_element(By.TAG_NAME, "p").text
        return description
    except seleniumException.NoSuchElementException:
        return ""


def get_company_website():
    contact_elem = driver.find_element(By.CLASS_NAME, "partner-contact")
    try:
        website_link = contact_elem.find_element(
            By.CLASS_NAME, "js-visit-website"
        ).get_attribute("href")
        if "https:" not in website_link:
            return ""
    except seleniumException.NoSuchElementException:
        return ""
    return website_link


def get_contact_name():
    try:
        contact_elem = driver.find_element(By.CLASS_NAME, "partner-contact")
        contact_divs = contact_elem.find_elements(By.TAG_NAME, "div")
        company_contact = contact_divs[5]
        company_contact_parts = company_contact.text.split("\n")
        return company_contact_parts[1]
    except seleniumException.NoSuchElementException:
        return ""


def get_contact_phone():
    try:
        contact_elem = driver.find_element(By.CLASS_NAME, "partner-contact")
        contact_divs = contact_elem.find_elements(By.TAG_NAME, "div")
        company_contact = contact_divs[5]
        company_contact_parts = company_contact.text.split("\n")
        return company_contact_parts[3]
    except (IndexError, seleniumException.NoSuchElementException):
        return ""


def get_contact_email():
    try:
        contact_elem = driver.find_element(By.CLASS_NAME, "partner-contact")
        contact_divs = contact_elem.find_elements(By.TAG_NAME, "div")
        company_contact = contact_divs[5]
        company_contact_parts = company_contact.text.split("\n")
        return company_contact_parts[2]
    except seleniumException.NoSuchElementException:
        return ""


def get_specialisations():
    try:
        spec_elem = driver.find_element(By.CLASS_NAME, "partner-types")
    except seleniumException.NoSuchElementException:
        return ""

    try:
        expand_spec_btn = spec_elem.find_element(By.CLASS_NAME, "plus")
        expand_spec_btn.click()
        expand_spec_btn.click()
    except (
        seleniumException.NoSuchElementException,
        seleniumException.ElementNotInteractableException,
    ):
        None

    spec_items = spec_elem.find_elements(By.CLASS_NAME, "type-item")
    items_arr = []
    for item in spec_items:
        if item.text != "":
            items_arr.append(item.text.strip())

    return " / ".join(items_arr)


def get_color():
    try:
        spec_elem = driver.find_element(By.CLASS_NAME, "partner-types")
        spec_items = spec_elem.find_elements(By.CLASS_NAME, "type-item")
    except seleniumException.NoSuchElementException:
        return ""

    color_arr = []
    for item in spec_items:
        if item.text == "" or item.text == "NONE":
            continue
        item_class = item.get_attribute("class")
        split_item_class = item_class.split(" ")
        try:
            color_class = split_item_class[1].strip()
        except IndexError:
            continue
        if color_class not in color_arr and color_class != "other":
            color_arr.append(color_class)

    return " / ".join(color_arr)


def get_company_cvr():
    contact_elem = driver.find_element(By.CLASS_NAME, "partner-contact")
    contact_divs = contact_elem.find_elements(By.TAG_NAME, "div")
    company_contact = contact_divs[0]
    company_contact_parts = company_contact.text.split("\n")
    for part in company_contact_parts:
        if "CVR" in part:
            cvr_num = part.split("CVR:")[1].strip()
            return cvr_num


# Global variables and configuration
case_links = []
cases = []
driver = webdriver.Firefox()


def main():
    print("Collecting website links")
    get_case_links()
    print("Successfully collected case links")
    print(f"Found {len(case_links)} cases")

    for index, case_link in enumerate(case_links):
        # Progress visualization
        percentage = index / len(case_links) * 100
        formatted_percentage = "{:.2f}".format(percentage)
        print(f"{formatted_percentage}% Complete")

        # Website data configuration
        company_data = {}
        driver.get(case_link)
        sleep(1)

        # Save the website data in variables
        company_data["company_name"] = get_company_name()
        company_data["company_description"] = get_company_description()
        company_data["specialisations"] = get_specialisations()
        company_data["colors"] = get_color()
        company_data["link"] = case_link
        company_data["website"] = get_company_website()
        company_data["cvr"] = get_company_cvr()
        company_data["contact_name"] = get_contact_name()
        company_data["contact_email"] = get_contact_email()
        company_data["contact_phone"] = get_contact_phone()

        # Bundle the website data variables together with the data from the previous website data
        cases.append(company_data)

    driver.quit()
    save_to_excel()
    print("Done!")


main()
