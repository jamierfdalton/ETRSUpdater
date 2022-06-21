""" Autoupdater for Engineering Timing Release Plan and eBOM Export

Gathers data from Upchain eBOM exports and consolidates them into
a single excel file. Requires BOM Exports from Upchain on a daily
basis run through the BOM Analyser VBA document and the custom
Upchain reporting that is recieved by email daily.

"""

from datetime import date, timedelta
import datetime
import os
import glob
import pandas as pd
import openpyxl
import gspread

BASE_PATH = r"S:\PDM Files\P1 - Mustang\\"
TARGET_PATH = fr"{BASE_PATH}\ETRS\ETRS Master\ETRS v4 Master.xlsx"


def connect_to_google_sheet(sheet_id):
    """ Connects to Google Sheets and returns values as a list of lists """
    print("Connecting to Google Sheets...")
    g_c = gspread.service_account()
    sheet = g_c.open_by_key(sheet_id).sheet1
    print("Connected")
    return sheet


def write_to_finance_update_csv(sheet_key, filename):
    """ Retrieves values from Google Sheet specified in the sheetKey

    sheetKey is a string that can be found in the URL of the target sheet you
    are connecting to. This function retrieves values from the Google Sheet
    and saves them in a CSV labelled with today's date.
    """

    print("Retrieving values from Google Sheets...")
    google_sheet = connect_to_google_sheet(sheet_key)
    data_frame = pd.DataFrame(google_sheet.get_all_values())

    print("Writing values to CSV at " + filename)
    data_frame.to_csv(filename, index=False, header=False)


def load_data_file(source_path):
    """ Loads data from CSV or XLSX into a dataframe

    SourcePath should be the file path of either a CSV or an Excel. If the
    file is a Formatted BOM, the sheet name will be correctly labelled in the
    ETRS, otherwise it will follow standard Excel naming conventions
    """

    if source_path[-4:] == "csv ":
        print("CSV found, reading csv")
        data = pd.read_csv(source_path)
    elif source_path[-24:-14] == "BOM Export":
        print("BOM Export found, reading excel")
        data = pd.read_excel(source_path, sheet_name="Formatted BOM")
    elif source_path[-4:] == "xlsx":
        print("Excel doc found, reading excel")
        data = pd.read_excel(source_path)
    else:
        print("Document not found")

    return data


def excel_archiver():
    """ Moves any old excel documents to the Archive folder in the ETRS folder
    """

    existing_file_list = glob.glob(fr"{BASE_PATH}\ETRS\*.xlsx")
    print("Saving Export...")

    for i in existing_file_list:
        archive_base_path = fr"{BASE_PATH}\ETRS\Archive\\"
        archive_file_name = i[32:]

        try:
            os.rename(i, archive_base_path + archive_file_name)
        except FileExistsError:

            os.rename(i, archive_base_path + "New " + archive_file_name)

    print("Export Saved!")


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
    workflow_path = r"\BOM\Upchain Custom Reports\EBOM Reports\eBOM Workflow Report " # Trailing space is important here!

    finance_source = fr"{BASE_PATH}ETRS\DataFiles\Finance {today}.csv "
    today_bom_source = fr"{BASE_PATH}{bom_export_path}{today_bom_format}.xlsx"
    yesterday_bom_source = fr"{BASE_PATH}{bom_export_path}{yesterday_bom_format}.xlsx"
    weekend_bom_source = fr"{BASE_PATH}{bom_export_path}{weekend_bom_format}.xlsx"
    monday_bom_source = fr"{BASE_PATH}{bom_export_path}{monday_bom_format}.xlsx"
    workflow_source = fr"{BASE_PATH}{workflow_path}{today_bom_format}.xlsx"

    print("Loading ETRS Workbook " + TARGET_PATH)
    book = openpyxl.load_workbook(TARGET_PATH)

    with pd.ExcelWriter(TARGET_PATH, engine='openpyxl', mode='a', # pylint: disable=abstract-class-instantiated
                        if_sheet_exists="replace") as writer:
        print("Loading Finance Export Data")
        finance_data = load_data_file(finance_source)
        print("Loading today's BOM Export Data")
        today_bom_data = load_data_file(today_bom_source)

        print("Loading yesterday's BOM Export Data")

        if os.path.exists(yesterday_bom_source):
            yesterday_bom_data = load_data_file(yesterday_bom_source)
        elif os.path.exists(weekend_bom_source):
            print("Yesterday's BOM Export not found. Skipping the weekend")
            yesterday_bom_data = load_data_file(weekend_bom_source)
        else:
            print("Couldn't find the BOM data files for the dates requested")

        print("Loading Monday's BOM Export Data")
        monday_bom_data = load_data_file(monday_bom_source)

        print("Loading Workflow Data")
        workflow_data = load_data_file(workflow_source)

# Gen's feedback - loop this and avoid the explict calls
        sheet_name_1 = "BOM Export"
        sheet_name_2 = "Purchasing Lead Times"
        sheet_name_3 = "Workflow"
        sheet_name_4 = "Yesterday BOM Export"
        sheet_name_5 = "Monday's BOM Export"
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}

        print("Writing Finance Data to ETRS")
        finance_data.to_excel(writer, sheet_name=sheet_name_2)
        print("Writing today's BOM Export Data to ETRS")
        today_bom_data.to_excel(writer, sheet_name=sheet_name_1)
        print("Writing yesterday's BOM Export Data to ETRS")
        yesterday_bom_data.to_excel(writer, sheet_name=sheet_name_4)
        print("Writing Monday's BOM Export Data to ETRS")
        monday_bom_data.to_excel(writer, sheet_name=sheet_name_5)
        print("Writing Workflow Export Data to ETRS")
        workflow_data.to_excel(writer, sheet_name=sheet_name_3)

        print("Saving Master...")
        book.save(fr"{BASE_PATH}\ETRS\\ETRS " + str(date.today()) + ".xlsx")
        print(r"Master Saved!")

def main():
    """ Wrapper function for running the major elements of the script in order
    """

    print("Updating ETRS...")

    write_to_finance_update_csv(
        "1OZemQa88tV9a4_-21oaAQnt5mbAo1Y7WLTXCDM7jIoE",
        fr"{BASE_PATH}\ETRS\DataFiles\\Finance " + str(date.today()) + ".csv"
        )
    excel_archiver()
    write_to_etrs()

    print("Update Complete")

if __name__ == "__main__":
    main()
