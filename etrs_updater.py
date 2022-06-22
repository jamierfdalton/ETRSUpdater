""" Autoupdater for Engineering Timing Release Plan and eBOM Export

Gathers data from Upchain eBOM exports, google sheets,
custom Upchain email reports and consolidates them into
a single excel file, maintaining the formula that already exist
in the master file. Requires BOM Exports from Upchain on a daily
basis run through the BOM Analyser VBA document and the custom
Upchain reporting that is recieved by email daily.

"""

from datetime import date, timedelta
import datetime
import os
import glob
import logging
import pandas as pd
import openpyxl
import gspread

BASE_PATH = r"S:\PDM Files\P1 - Mustang\\"
TARGET_PATH = fr"{BASE_PATH}\ETRS\ETRS Master\ETRS v4 Master.xlsx"

logging.basicConfig(format='%(asctime)s:%(levelname)s:%(message)s - ',
                    encoding='utf-8',
                    datefmt='%Y-%m-%d %H:%M:%S',
                    level=logging.INFO,
                    handlers=[
                        logging.FileHandler("etrs_updater.log"),
                        logging.StreamHandler()
                        ]
                    )

def connect_to_google_sheet(sheet_id):
    """ Connects to Google Sheets and returns values as a list of lists """
    logging.info("Connecting to Google Sheets...")
    google_connect = gspread.service_account()
    sheet = google_connect.open_by_key(sheet_id).sheet1
    logging.info("Connected")
    return sheet


def write_to_finance_update_csv(sheet_key, filename):
    """ Retrieves values from Google Sheet specified in the sheet_key

    sheet_key is a string that can be found in the URL of the target sheet you
    are connecting to. This function retrieves values from the Google Sheet
    and saves them in a CSV labelled with today's date.
    """

    logging.info("Retrieving values from Google Sheets")
    google_sheet = connect_to_google_sheet(sheet_key)
    data_frame = pd.DataFrame(google_sheet.get_all_values())

    logging.info("Writing values to CSV at %s", filename)
    data_frame.to_csv(filename, index=False, header=False)


def load_data_file(source_path):
    """ Loads data from CSV or XLSX into a dataframe

    source_path should be the file path of either a CSV or an Excel. If the
    file is a Formatted BOM, the sheet name will be correctly labelled in the
    ETRS, otherwise it will follow standard Excel naming conventions
    """

    if source_path[-4:] == "csv ":
        logging.info("CSV found, reading csv")
        data = pd.read_csv(source_path)
    elif source_path[-24:-14] == "BOM Export":
        logging.info("BOM Export found, reading excel")
        data = pd.read_excel(source_path, sheet_name="Formatted BOM")
    elif source_path[-4:] == "xlsx":
        logging.info("Excel doc found, reading excel")
        data = pd.read_excel(source_path)
    else:
        logging.critical("Document not found")

    return data


def excel_archiver():
    """ Moves any old excel documents to the Archive folder in the ETRS folder
    """
existing_file_list = glob.glob(fr"{BASE_PATH}\ETRS\*.xlsx")
datetime_timestamp = datetime.datetime.now()
string_timestamp = datetime_timestamp.strftime("%H-%M-%S")

for i in existing_file_list:
    path, file = os.path.split(i)
    existing_filename, extension = os.path.splitext(file)
    archive_path = fr"{BASE_PATH}ETRS\Archive\{existing_filename} -- {string_timestamp}{extension}"
    os.rename(i,archive_path)


def write_to_etrs():
    """ Collects the various data sources and writes them to an XLSX file

    If you have the accompanying ETRS Master file at targetPath, this export
    will conform to the requirements of that sheet to automate the creation of
    a new ETRS file and archive the one. Ideally this process happens
    on a daily basis.
    """
    # TODO Gen's feedback -
    # f(strings) (done),
    # create a base path variable (done),
    # loop through this somehow?
    bom_export_path = r"\BOM\BOM Exports\BOM Export "
    today = date.today()
    monday = today - datetime.timedelta(days=today.weekday())
    monday_bom_format = str(monday.strftime('%Y%m%d'))
    today_bom_format = str(today.strftime('%Y%m%d'))
    yesterday_bom_format = str((today - timedelta(days=1)).strftime('%Y%m%d'))
    weekend_bom_format = str((today - timedelta(days=3)).strftime('%Y%m%d'))
     # Trailing space is important workflow_path!
    workflow_path = r"\BOM\Upchain Custom Reports\EBOM Reports\eBOM Workflow Report "

    finance_source = fr"{BASE_PATH}ETRS\DataFiles\Finance {today}.csv "
    today_bom_source = fr"{BASE_PATH}{bom_export_path}{today_bom_format}.xlsx"
    yesterday_bom_source = fr"{BASE_PATH}{bom_export_path}{yesterday_bom_format}.xlsx"
    weekend_bom_source = fr"{BASE_PATH}{bom_export_path}{weekend_bom_format}.xlsx"
    monday_bom_source = fr"{BASE_PATH}{bom_export_path}{monday_bom_format}.xlsx"
    workflow_source = fr"{BASE_PATH}{workflow_path}{today_bom_format}.xlsx"

    logging.info("Loading ETRS Workbook %s", TARGET_PATH)
    book = openpyxl.load_workbook(TARGET_PATH)

    with pd.ExcelWriter(TARGET_PATH, engine='openpyxl', mode='a', # pylint: disable=abstract-class-instantiated
                        if_sheet_exists="replace") as writer:
        logging.info("Loading Finance Export Data")
        finance_data = load_data_file(finance_source)
        logging.info("Loading today's BOM Export Data")
        today_bom_data = load_data_file(today_bom_source)

        logging.info("Loading yesterday's BOM Export Data")

        if os.path.exists(yesterday_bom_source):
            yesterday_bom_data = load_data_file(yesterday_bom_source)
        elif os.path.exists(weekend_bom_source):
            logging.info("Yesterday's BOM Export not found. Skipping the weekend")
            yesterday_bom_data = load_data_file(weekend_bom_source)
        else:
            logging.info("Couldn't find the BOM data files for the dates requested")

        logging.info("Loading Monday's BOM Export Data")
        monday_bom_data = load_data_file(monday_bom_source)

        logging.info("Loading Workflow Data")
        workflow_data = load_data_file(workflow_source)

# Gen's feedback - loop this and avoid the explict calls
        sheet_name_1 = "BOM Export"
        sheet_name_2 = "Purchasing Lead Times"
        sheet_name_3 = "Workflow"
        sheet_name_4 = "Yesterday BOM Export"
        sheet_name_5 = "Monday's BOM Export"
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}

        logging.info("Writing Finance Data to ETRS")
        finance_data.to_excel(writer, sheet_name=sheet_name_2)
        logging.info("Writing today's BOM Export Data to ETRS")
        today_bom_data.to_excel(writer, sheet_name=sheet_name_1)
        logging.info("Writing yesterday's BOM Export Data to ETRS")
        yesterday_bom_data.to_excel(writer, sheet_name=sheet_name_4)
        logging.info("Writing Monday's BOM Export Data to ETRS")
        monday_bom_data.to_excel(writer, sheet_name=sheet_name_5)
        logging.info("Writing Workflow Export Data to ETRS")
        workflow_data.to_excel(writer, sheet_name=sheet_name_3)

        logging.info("Saving Master...")
        book.save(fr"{BASE_PATH}\ETRS\\ETRS " + str(date.today()) + ".xlsx")
        logging.info(r"Master Saved!")

def main():
    """ Wrapper function for running the major elements of the script in order
    """

    logging.info("\n\n")
    logging.info("Updating ETRS...")

    write_to_finance_update_csv(
        "1OZemQa88tV9a4_-21oaAQnt5mbAo1Y7WLTXCDM7jIoE",
        fr"{BASE_PATH}\ETRS\DataFiles\\Finance " + str(date.today()) + ".csv"
        )
    excel_archiver()
    write_to_etrs()

    logging.info("Update Successful")

if __name__ == "__main__":
    main()
