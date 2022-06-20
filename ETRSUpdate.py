""" Automated updates for Engineering Timing Release Plan and BOM Validation

Gathers data from Upchain eBOM exports and consolidates them into
a single excel file. Requires BOM Exports from Upchain on a daily basis
run through the BOM Analyser VBA document and the custom Upchain reporting that
is recieved by email daily.

"""

import gspread
from datetime import date, timedelta
import datetime
import pandas as pd
import openpyxl
import os
import glob

global basePath
basePath = "S:\PDM Files\P1 - Mustang\\"
global targetPath
targetPath = fr"{basePath}\ETRS\ETRS Master\ETRS v4 Master.xlsx"


def connect_to_google_sheet(sheetId):
    """ Connects to Google Sheets and returns values as a list of lists """
    print("Connecting to Google Sheets...")
    gc = gspread.service_account()
    sheet = gc.open_by_key(sheetId).sheet1
    print("Connected")
    return sheet


def write_to_finance_update_csv(sheetKey, filename):
    """ Retrieves values from Google Sheet specified in the sheetKey

    sheetKey is a string that can be found in the URL of the target sheet you
    are connecting to. This function retrieves values from the Google Sheet
    and saves them in a CSV labelled with today's date.
    """

    print("Retrieving values from Google Sheets...")
    sh = connect_to_google_sheet(sheetKey)
    df = pd.DataFrame(sh.get_all_values())

    print("Writing values to CSV at " + filename)
    df.to_csv(filename, index=False, header=False)


def load_data_file(sourcePath):
    """ Loads data from CSV or XLSX into a dataframe

    SourcePath should be the file path of either a CSV or an Excel. If the
    file is a Formatted BOM, the sheet name will be correctly labelled in the
    ETRS, otherwise it will follow standard Excel naming conventions
    """

    #TODO change this to cases rather than if/elif
    if sourcePath[-4:] == "csv ":
        print("CSV found, reading csv")
        data = pd.read_csv(sourcePath)
    elif sourcePath[-24:-14] == "BOM Export":
        print("BOM Export found, reading excel")
        data = pd.read_excel(sourcePath, sheet_name="Formatted BOM")
    elif sourcePath[-4:] == "xlsx":
        print("Excel doc found, reading excel")
        data = pd.read_excel(sourcePath)
    else:
        print("Document not found")

    return data


def excel_archiver():
    """ Moves any old excel documents to the Archive folder in the ETRS folder
    """

    existingFileList = glob.glob(f"{basePath}\ETRS\*.xlsx")
    print("Saving Export...")

    for i in existingFileList:
        archiveFilename = f"{basePath}\ETRS\Archive\\" + i[31:]
        os.rename(i, archiveFilename)

    print("Export Saved!")


def write_to_ETRS():
    """ Collects the various data sources and writes them to an XLSX file

    If you have the accompanying ETRS Master file at targetPath, this export
    will conform to the requirements of that sheet to automate the creation of
    a new ETRS file and archive the one. Ideally this process happens on a daily
    basis.
    """
    # Gen's feedback - f(strings) (done), create a base path variable (done), loop through this somehow?
    BOMExportPath = "\BOM\BOM Exports\BOM Export "
    today = date.today()
    monday = today - datetime.timedelta(days=today.weekday())
    mondayBOMFormat = str(monday.strftime('%Y%m%d'))
    todayBOMFormat = str(today.strftime('%Y%m%d'))
    yesterdayBOMFormat = str((today - timedelta(days=1)).strftime('%Y%m%d'))
    weekendBOMFormat = str((today - timedelta(days=3)).strftime('%Y%m%d'))
    workflowPath = r"\BOM\Upchain Custom Reports\EBOM Reports\eBOM Workflow Report "

    financeSource = fr"{basePath}ETRS\DataFiles\Finance {today}.csv "
    todayBOMSource = fr"{basePath}{BOMExportPath}{todayBOMFormat}.xlsx"
    yesterdayBOMSource = fr"{basePath}{BOMExportPath}{yesterdayBOMFormat}.xlsx"
    weekendBOMSource = fr"{basePath}{BOMExportPath}{weekendBOMFormat}.xlsx"
    mondayBOMSource = fr"{basePath}{BOMExportPath}{mondayBOMFormat}.xlsx"
    workflowSource = fr"{basePath}{workflowPath}{todayBOMFormat}.xlsx"
    savePath = fr"{basePath}\ETRS\ETRS Exports\\ETRS " + str(date.today())

    print("Loading ETRS Workbook " + targetPath)
    book = openpyxl.load_workbook(targetPath)

    with pd.ExcelWriter(targetPath, engine='openpyxl', mode='a',
                        if_sheet_exists="replace") as writer:
        print("Loading Finance Export Data")
        financeData = load_data_file(financeSource)
        print("Loading today's BOM Export Data")
        todayBOMData = load_data_file(todayBOMSource)

        print("Loading yesterday's BOM Export Data")

        if os.path.exists(yesterdayBOMSource):
            yesterdayBOMData = load_data_file(yesterdayBOMSource)
        elif os.path.exists(weekendBOMSource):
            print("Yesterday's BOM Export Data not found. Skipping the weekend")
            yesterdayBOMData = load_data_file(weekendBOMSource)
        else:
            print("Couldn't find the BOM data files for the dates requested")

        print("Loading Monday's BOM Export Data")
        mondayBOMData = load_data_file(mondayBOMSource)

        print("Loading Workflow Data")
        workflowData = load_data_file(workflowSource)

# Gen's feedback - loop this and avoid the explict calls
        sheetName1 = "BOM Export"
        sheetName2 = "Purchasing Lead Times"
        sheetName3 = "Workflow"
        sheetName5 = "Yesterday BOM Export"
        sheetName6 = "Monday's BOM Export"
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}

        print("Writing Finance Data to ETRS")
        financeData.to_excel(writer, sheet_name=sheetName2)
        print("Writing today's BOM Export Data to ETRS")
        todayBOMData.to_excel(writer, sheet_name=sheetName1)
        print("Writing yesterday's BOM Export Data to ETRS")
        yesterdayBOMData.to_excel(writer, sheet_name=sheetName5)
        print("Writing Monday's BOM Export Data to ETRS")
        mondayBOMData.to_excel(writer, sheet_name=sheetName6)
        print("Writing Workflow Export Data to ETRS")
        workflowData.to_excel(writer, sheet_name=sheetName3)

        print("Saving Master...")
        book.save(f"{basePath}\ETRS\\ETRS " + str(date.today())+ ".xlsx")
        print("Master Saved!")


def main():
    """ Wrapper function for running the major elements of the script in order
    """

    print("Updating ETRS...")

    write_to_finance_update_csv(
        "1OZemQa88tV9a4_-21oaAQnt5mbAo1Y7WLTXCDM7jIoE",
        f"{basePath}\ETRS\DataFiles\\Finance " + str(date.today()) + ".csv"
        )
    write_to_ETRS()
    excel_archiver()

    print("Update Complete")

if __name__ == "__main__":
    main()
