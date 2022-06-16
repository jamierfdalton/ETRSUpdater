import gspread
from datetime import date, timedelta
import pandas as pd
import openpyxl
import os
import glob

global targetPath
targetPath = "S:\PDM Files\P1 - Mustang\ETRS\ETRS Master\ETRS v3 Master.xlsx"


def connect_to_google_sheet(sheetId):
    # Connects to the Google Sheet specified in the sheetID and returns the sheet
    print("Connecting to Google Sheets...")
    gc = gspread.service_account()
    sheet = gc.open_by_key(sheetId).sheet1
    print("Connected")
    return sheet

def write_to_finance_update_csv():
    # Gets all values from the open google sheet and stores them as a dataframe
    # Then writes that dataframe to a CSV containing today's date in the filename
    # (Needs to save so we have a snapshot of each day)
    print("Retrieving values from Google Sheets...")
    sh = connect_to_google_sheet("1OZemQa88tV9a4_-21oaAQnt5mbAo1Y7WLTXCDM7jIoE")
    df = pd.DataFrame(sh.get_all_values())

    filename = r"S:\PDM Files\P1 - Mustang\ETRS\DataFiles\\Finance " + str(date.today()) +".csv"
    print("Writing values to CSV at " + filename)
    df.to_csv(filename, index=False, header=False)

def load_finance_update_csv():
    # Use pandas to read today's finance CSV
    sourcePath = r"S:\PDM Files\P1 - Mustang\ETRS\DataFiles\\Finance " + str(date.today()) +".csv"
    print("Loading Finance source CSV " + sourcePath)
    financeData = pd.read_csv(sourcePath)
    return financeData

def load_BOM_Export_data(date):
    # Use pandas to read today's BOM Export
    print("Loading BOM Export "+ date +".xlsx from shared drive")
    sourcePath = r"S:\PDM Files\P1 - Mustang\BOM\BOM Exports\BOM Export " + date +".xlsx"
    BOMExportData = pd.read_excel(sourcePath, sheet_name="Formatted BOM")
    return BOMExportData

def load_Raw_BOM_Export_data():
    # Use pandas to read today's BOM Export
    print("Loading BOM Export "+ str(date.today().strftime('%Y%m%d')) +".xlsx from shared drive")
    sourcePath = r"S:\PDM Files\P1 - Mustang\BOM\BOM Exports\BOM Export " + str(date.today().strftime('%Y%m%d')) +".xlsx"
    RawBOMExportData = pd.read_excel(sourcePath, sheet_name="Raw BOM")
    return RawBOMExportData

def load_Workflow_Export_data():
    # Use pandas to read today's BOM Workflow Export
    print("eBOM Workflow Report "+ str(date.today().strftime('%Y%m%d')) +".xlsx")
    sourcePath = r"S:\PDM Files\P1 - Mustang\BOM\Upchain Custom Reports\EBOM Reports\\eBOM Workflow Report " + date.today().strftime('%Y%m%d') + ".xlsx"
    workflowExportData = pd.read_excel(sourcePath, sheet_name="Report Data")
    return workflowExportData

def write_to_ETRS():
    today = str(date.today().strftime('%Y%m%d'))
    yesterday = str((date.today() - timedelta(days=1)).strftime("%Y%m%d"))
    weekendDelta = str((date.today() - timedelta(days=3)).strftime("%Y%m%d"))
    savePath = r"S:\PDM Files\P1 - Mustang\ETRS\ETRS Exports\\ETRS " + str(date.today())
    existingFileList = glob.glob(r"S:\PDM Files\P1 - Mustang\ETRS\*.xlsx")

    print("Loading ETRS Workbook " + targetPath)
    book = openpyxl.load_workbook(targetPath)

    with pd.ExcelWriter(targetPath, engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
        print("Loading Finance Export Data")
        financeData = load_finance_update_csv()
        print("Loading today's BOM Export Data")
        todayBOMData = load_BOM_Export_data(today)

        print("Loading yesterday's BOM Export Data")
        try:
            yesterdayBOMData = load_BOM_Export_data(yesterday)
        except:
            try:
                print("Yesterday's BOM Export Data not found. Skipping the weekend")
                yesterdayBOMData = load_BOM_Export_data(weekendDelta)
            except:
                print("ERROR Yesterday's BOM and last Friday's BOM not found.")
                # Add a user input date picker here?

        print("Loading Workflow Data")
        workflowData = load_Workflow_Export_data()

        sheetName1 = "BOM Export"
        sheetName2 = "Purchasing Lead Times"
        sheetName3 = "Workflow"
        sheetName4 = "Raw BOM"
        sheetName5 = "Old BOM Export"
        writer.book = book
        writer.sheets = {ws.title:ws for ws in book.worksheets}

        print("Writing Finance Data to ETRS")
        financeData.to_excel(writer, sheet_name=sheetName2)
        print("Writing today's BOM Export Data to ETRS")
        todayBOMData.to_excel(writer,sheet_name=sheetName1)
        print("Writing yesterday's BOM Export Data to ETRS")
        todayBOMData.to_excel(writer,sheet_name=sheetName5)
        print("Writing Raw BOM Export Data to ETRS")
        yesterdayBOMData.to_excel(writer,sheet_name=sheetName4)
        print("Writing Workflow Export Data to ETRS")
        workflowData.to_excel(writer,sheet_name=sheetName3)
        print("Saving Master...")
        book.save(r"S:\PDM Files\P1 - Mustang\ETRS\ETRS Master\ETRS v3 Master.xlsx")
        print("Master Saved!")
        print("Saving Export...")

        # Archive existing .xlsx files in the main directory.
        for i in existingFileList:
            archiveFilename = r"S:\PDM Files\P1 - Mustang\ETRS\Archive\\" + i[31:]
            try:
                os.rename(i,archiveFilename)
            except:
                pass

        book.save(r"S:\PDM Files\P1 - Mustang\ETRS\\ETRS " + str(date.today()) + ".xlsx")
        print("Export Saved!")

def main():
    print("Updating ETRS...")
    write_to_finance_update_csv()
    write_to_ETRS()
    print("Update Complete")

if __name__ == "__main__":
    main()
