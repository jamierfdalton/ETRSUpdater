import gspread
from datetime import date, timedelta
import pandas as pd
import openpyxl
import os
import glob

global basePath
basePath = r"S:\PDM Files\P1 - Mustang\"
global targetPath
targetPath = fr"{basePath}\ETRS\ETRS Master\ETRS v3 Master.xlsx"

print("Jamie is Great Master")


def connect_to_google_sheet(sheetId):
    # Connects to the Google Sheet specified in the sheetID
    # and returns the sheet
    print("Connecting to Google Sheets...")
    gc = gspread.service_account()
    sheet = gc.open_by_key(sheetId).sheet1
    print("Connected")
    return sheet


def write_to_finance_update_csv():
    # Gets all values from the open google sheet and stores them as a dataframe
    # Then writes that dataframe to CSV containing today's date in the filename
    # (Needs to save so we have a snapshot of each day)
    print("Retrieving values from Google Sheets...")
    sh = connect_to_google_sheet("1OZemQa88tV9a4_-21oaAQnt5mbAo1Y7WLTXCDM7jIoE")
    df = pd.DataFrame(sh.get_all_values())

    filename = f"{basePath}\ETRS\DataFiles\\Finance " + str(date.today()) + ".csv"
    print("Writing values to CSV at " + filename)
    df.to_csv(filename, index=False, header=False)


def load_data_file(sourcePath):
    # load data file using pandas
    if sourcePath[-3:] == "csv":
        print("CSV found, reading csv")
        data = pd.read_csv(sourcePath)
    elif sourcePath[10:] == "BOM Export":
        print("BOM Export found, reading excel")
        data = pd.read_excel(sourcePath, sheet_name="Formatted BOM")
    elif sourcePath[-4:] == "xlsx":
        print("Excel doc found, reading excel")
        data = pd.read_excel(sourcePath)
    return data


def write_to_ETRS():
    # Gen's feedback - f(strings), create a base path variable,
    BOMExportPath = "\BOM\BOM Exports\BOM Export "
    today = date.today()
    todayBOMFormat = str(today.strftime('%Y%m%d'))
    yesterdayBOMFormat = str(today - timedelta(days=1).strftime('%Y%m%d'))
    weekendBOMFormat = str(today - timedelta(days=3).strftime('%Y%m%d'))

    existingFileList = glob.glob(f"{basePath}\ETRS\*.xlsx")
    financeSource = f"{basePath}\ETRS\DataFiles\\Finance {today}.csv "
    todayBOMSource = f"{basePath}{BOMExportPath}{todayBOMFormat}.xlsx"
    yesterdayBOMSource = f"{basePath}{BOMExportPath}{yesterdayBOMFormat}.xlsx"
    weekendBOMSource = f"{basePath}{BOMExportPath}{weekendBOMFormat}.xlsx"
    workflowSource = fr"{basePath}\BOM\Upchain Custom Reports\EBOM Reports\eBOM Workflow Report {todayBOMFormat}.xlsx"
    savePath = f"{basePath}\ETRS\ETRS Exports\\ETRS " + str(date.today())

    print("Loading ETRS Workbook " + targetPath)
    book = openpyxl.load_workbook(targetPath)

    with pd.ExcelWriter(targetPath, engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
        print("Loading Finance Export Data")
        financeData = load_data_file(financeSource)
        print("Loading today's BOM Export Data")
        todayBOMData = load_data_file(todayBOMSource)

        print("Loading yesterday's BOM Export Data")
        try:
            yesterdayBOMData = load_data_file(yesterdayBOMSource)
        except:
            try:
                print("Yesterday's BOM Export Data not found. Skipping the weekend")
                yesterdayBOMData = load_data_file(weekendBOMSource)
            except:
                print("ERROR Yesterday's BOM and last Friday's BOM not found.")
                # Add a user input date picker here?

        print("Loading Workflow Data")
        workflowData = load_data_file(workflowSource)

# Gen's feedback - loop this and avoid the explict calls
        sheetName1 = "BOM Export"
        sheetName2 = "Purchasing Lead Times"
        sheetName3 = "Workflow"
        sheetName4 = "Raw BOM"
        sheetName5 = "Old BOM Export"
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}

        print("Writing Finance Data to ETRS")
        financeData.to_excel(writer, sheet_name=sheetName2)
        print("Writing today's BOM Export Data to ETRS")
        todayBOMData.to_excel(writer, sheet_name=sheetName1)
        print("Writing yesterday's BOM Export Data to ETRS")
        todayBOMData.to_excel(writer, sheet_name=sheetName5)
        print("Writing Raw BOM Export Data to ETRS")
        yesterdayBOMData.to_excel(writer, sheet_name=sheetName4)
        print("Writing Workflow Export Data to ETRS")
        workflowData.to_excel(writer, sheet_name=sheetName3)
        print("Saving Master...")
        book.save(f"{basePath}\ETRS\ETRS Master\ETRS v3 Master.xlsx")
        print("Master Saved!")
        print("Saving Export...")

        # Archive existing .xlsx files in the main directory.
        for i in existingFileList:
            archiveFilename = f"{basePath}\ETRS\Archive\\" + i[31:]
            try:
                os.rename(i, archiveFilename)
            except:
                pass

        book.save(f"{basePath}\ETRS\\ETRS " + str(date.today()) + ".xlsx")
        print("Export Saved!")


def main():
    print("Updating ETRS...")
    # write_to_finance_update_csv()
    # write_to_ETRS()
    print("Update Complete")

if __name__ == "__main__":
    main()
