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
import numpy as np
import openpyxl
import xlwings as xl
import gspread

BASE_PATH = r"S:\PDM Files\P1 - Mustang\\"
TARGET_PATH = fr"{BASE_PATH}\ETRS\ETRS Master\ETRS v7 Master.xlsx"

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
    ETRS, otherwise it will follow default naming conventions
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
    """ Moves any old excel documents to the Archive folder from the ETRS folder
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

    bom_export_path = r"\BOM\BOM Exports\BOM Export "
    today = date.today()
    monday = today - datetime.timedelta(days=today.weekday())
    monday_bom_format = str(monday.strftime('%Y%m%d'))
    today_bom_format = str(today.strftime('%Y%m%d'))
    yesterday_bom_format = str((today - timedelta(days=1)).strftime('%Y%m%d'))
    weekend_bom_format = str((today - timedelta(days=3)).strftime('%Y%m%d'))

    # Trailing space is important workflow_path!

    # workflow_path = r"\BOM\Upchain Custom Reports\EBOM Reports\eBOM Workflow Report "
    purchasing_source = fr"{BASE_PATH}ETRS\DataFiles\Finance {today}.csv "
    today_bom_source = fr"{BASE_PATH}{bom_export_path}{today_bom_format}.xlsx"
    yesterday_bom_source = fr"{BASE_PATH}{bom_export_path}{yesterday_bom_format}.xlsx"
    weekend_bom_source = fr"{BASE_PATH}{bom_export_path}{weekend_bom_format}.xlsx"
    monday_bom_source = fr"{BASE_PATH}{bom_export_path}{monday_bom_format}.xlsx"
    # workflow_source = fr"{BASE_PATH}{workflow_path}{today_bom_format}.xlsx"


    logging.info("Loading ETRS Workbook %s", TARGET_PATH)
    book = openpyxl.load_workbook(TARGET_PATH)

    with pd.ExcelWriter(TARGET_PATH, engine='openpyxl', mode='a', # pylint: disable=abstract-class-instantiated
                        if_sheet_exists="replace") as writer:

        all_data_exports = {
            "today_bom_export" : ["BOM Export", today_bom_source],
            "purchasing_export" : ["Purchasing Lead Times", purchasing_source],
            # "workflow_export" : ["Workflow",workflow_source],
            "yesterday_bom_export" : ["Yesterday BOM Export",yesterday_bom_source],
            "monday_bom_export" : ["Monday's BOM Export",monday_bom_source]
        }

        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}

        for j in all_data_exports:
            if os.path.exists(all_data_exports[j][1]):
                logging.info("Loading  %s from source" , all_data_exports[j][0])
                loaded_data = load_data_file(all_data_exports[j][1])
                logging.info("Writing %s to ETRS" , all_data_exports[j][0])
                loaded_data.to_excel(writer, sheet_name=all_data_exports[j][0])

            elif os.path.exists(weekend_bom_source):
                logging.info("Loading  %s from source" , "Last Friday's BOM Export")
                loaded_data = load_data_file(weekend_bom_source)
                logging.info("Writing %s to ETRS" , "Last Friday's BOM Export")

                # Last Friday's BOM Export has to be named
                # Yesterday BOM Export for the excel formula to work in Master File
                loaded_data.to_excel(writer, sheet_name="Yesterday BOM Export")

            else:
                logging.critical("Source files not found!")

        logging.info("Saving Master...")
        # DEBUGGING PATH
        # book.save(fr"{BASE_PATH}\ETRS\\ETRS Master\\ETRS " + str(date.today()) + ".xlsx")
        # REAL PATH
        book.save(fr"{BASE_PATH}\ETRS\\ETRS " + str(date.today()) + ".xlsx")
        logging.info(r"Master Saved!")

def refresh_excel_values(path):
    app = xl.App(visible=False)
    book = app.books.open(path)
    book.save()
    app.kill()


def tableify_etrs():
    """ Manipulates the ETRS in order to make it easier to report stats on"""

    today = date.today()
    bom_export_path = r"\BOM\BOM Exports\BOM Export "
    today_bom_format = str(today.strftime('%Y%m%d'))
    workflow_path = r"\BOM\Upchain Custom Reports\EBOM Reports\eBOM Workflow Report "
    todays_etrs = fr"{BASE_PATH}\ETRS\\ETRS Master\\ETRS " + str(date.today()) + ".xlsx"
    workflow_file = fr"{BASE_PATH}{workflow_path}{today_bom_format}.xlsx"
    today_bom_source = fr"{BASE_PATH}{bom_export_path}{today_bom_format}.xlsx"

    refresh_excel_values(todays_etrs) # refresh ETRS values
    


    etrs_df = pd.read_excel(r"S:\PDM Files\P1 - Mustang\ETRS\ETRS Master\ETRS 2022-06-29(1).xlsx", sheet_name="BTRS", skiprows = 4)
    workflow_df = pd.read_excel(r"S:\PDM Files\P1 - Mustang\BOM\Upchain Custom Reports\EBOM Reports\eBOM Workflow Report " + str(today.strftime('%Y%m%d') + ".xlsx"))
   # bom_df = pd.read_excel(today_bom_source, sheet_name = "Raw BOM")


    etrs_df.rename(columns={ 
        etrs_df.columns[22]: "Reqs Item Name",
        etrs_df.columns[23]: "Reqs Item Description",
        etrs_df.columns[24]: "Reqs Quantity",
        etrs_df.columns[25]: "Reqs Type",
        etrs_df.columns[26]: "Reqs Mass (grams)",
        etrs_df.columns[27]: "Reqs Revision Note",
        etrs_df.columns[28]: "Reqs Material",
        etrs_df.columns[29]: "Reqs Finish",
        etrs_df.columns[30]: "Reqs Safety Critical",
        etrs_df.columns[31]: "Reqs 3D Model",
        etrs_df.columns[32]: "Reqs Fixings and Torque Value Required",
        etrs_df.columns[33]: "Reqs Fixings and Torque Value Complete",
        etrs_df.columns[34]: "Reqs 2D Drawings Required",
        etrs_df.columns[35]: "Reqs 2d Drawings Complete",
        etrs_df.columns[40]: "Reqs Supplier Nomination Status",
        etrs_df.columns[41]: "Reqs Manufacturer",
        etrs_df.columns[42]: " - ",
        etrs_df.columns[43]: "Reqs Development Lead Time",
        etrs_df.columns[44]: "Reqs Tool Lead Time",
        etrs_df.columns[45]: "Reqs Production Lead Time",
        etrs_df.columns[47]: "Reqs PPAP Complete",
        etrs_df.columns[48]: "Reqs PPAP Timing",
        etrs_df.columns[49]: "Reqs SIP Timing",
        etrs_df.columns[51]: "Item Number"
    }, inplace=True)

    selected_columns_df = etrs_df.filter([
        "Item Number",
        "Treepath", 
        "Revision Note", 
        "Function Group", 
        "Status",
        "Reqs Item Name",
        "Reqs Item Description",
        "Reqs Quantity",
        "Reqs Type",
        "Reqs Mass (grams)",
        "Reqs Revision Note",
        "Reqs Material",
        "Reqs Finish",
        "Reqs Safety Critical",
        "Reqs 3D Model",
        "Reqs Fixings and Torque Value Required",
        "Reqs Fixings and Torque Value Complete",
        "Reqs 2D Drawings Required",
        "Reqs 2d Drawings Complete",
        "Reqs Supplier Nomination Status",
        "Reqs Manufacturer",
        "Reqs Development Lead Time",
        "Reqs Tool Lead Time",
        "Reqs Production Lead Time",
        "Reqs PPAP Complete",
        "Reqs PPAP Timing",
        "Reqs SIP Timing",
        "New Part from Yesterday",
        "New Part from Monday",
        ])

    merged_df = pd.merge(selected_columns_df, workflow_df, how="outer", on="Item Number")
       
    gateway_conditions = [
        (merged_df["Workflow"] == "PDM Release") & (merged_df["Revision Note_x"].str.contains('G2', na = False)),
        (merged_df["Workflow"].str.contains("G2", na = False)),
        (merged_df["Workflow"].str.contains("G3", na = False)),
        (merged_df["Workflow"].str.contains("G1", na = False)),
        (merged_df["Revision Note_x"].str.contains("G3", na = False)),
        (merged_df["Revision Note_x"].str.contains("G2", na = False)),
        (merged_df["Revision Note_x"].str.contains("G1", na = False)),
        (merged_df["Revision Note_y"].str.contains("G1", na = False)),
        (merged_df["Revision Note_y"].str.contains("Initial", na = False)),
    ]

    merged_df["Gateway"] = np.select(gateway_conditions, 
                                    ["G2 - Release", "G2 - Release", "G3 - Release", 
                                     "G1 - Release", "G3 - Release", "G2 - Release",
                                     "G1 - Release", "G1 - Release", "G1 - Release"],
                                     default = "Unreleased")
   

    merged_df["Length of Item Name"] = merged_df["Item Name"].astype(str).map(len)

    fixing_conditions = [
        (merged_df["Item Name"].str.contains("PHANTOM", na = False)),
        (merged_df["Length of Item Name"] > 10),
        # TODO Add item name ends in 8s
        ]


    merged_df["Fixing"] = np.select(fixing_conditions, ["No", "Yes"])

    merged_df.to_csv(fr"{BASE_PATH}\ETRS\ETRS Master\output.csv")

    prepping_flat_df = merged_df.filter(["Item Name", "Item Description", "Quantity", "Part Type", "Item Number"])
    prepping_flat_df.set_index("Item Number")

    print(prepping_flat_df.info())
    print(prepping_flat_df.head())
    test_df = prepping_flat_df.loc[(prepping_flat_df['Part Type'] == "Purchased Item") | (prepping_flat_df['Part Type'] == "Purchased ElectroMechanical Part") | (prepping_flat_df['Part Type'] == "Purchased Electrical Part") | (prepping_flat_df['Part Type'] == "Purchased Mechanical Part")]
    
    
    
    output = test_df.pivot_table(index = ["Item Name", "Item Description", "Part Type", "Item Number"], values="Quantity")
    
    print(output.info())
    print(output["Part Type"])
    # output.to_csv(fr"{BASE_PATH}\ETRS\ETRS Master\output2.csv")


def main():
    """ Wrapper function for running the major elements of the script in order
    """

    logging.info("\n\n")
    logging.info("Updating ETRS...")

    
    
    # write_to_finance_update_csv(
    #     "1OZemQa88tV9a4_-21oaAQnt5mbAo1Y7WLTXCDM7jIoE",
    #     fr"{BASE_PATH}\ETRS\DataFiles\\Finance " + str(date.today()) + ".csv"
    #     )
    # excel_archiver()
    # write_to_etrs()
    tableify_etrs()
 
    logging.info("Update Successful")

if __name__ == "__main__":
    main()