import os
import time
from datetime import datetime

import config
import pandas as pd
import s3fs
import xlwings as xw
from appscript.reference import CommandError

from ahl_share_action_modelling.getters.base_getters import get_df_from_excel_s3_path

#
# PIPELINE A :
# Initial Transformation of ATNi Data, Aggregating into a Single Table.


def load_s3_data(bucket_path: str) -> dict:  # Reading S3 Data
    """Loads a dictionary of dataframes from all .xlsx files in the specified directory
    args:
        bucket_path(str): path to s3 directory
    returns:
        dict: dataframes of loaded files where keys are corresponding file names
    """
    fs = s3fs.S3FileSystem()
    if not bucket_path.endswith("/"):
        bucket_path += "/"
    files = fs.glob(f"{bucket_path}*.xlsx")

    loaded_data = {}
    for file_path in files:
        file_name = file_path.split("/")[-1]
        full_s3_path = f"s3://{file_path}" if not file_path.startswith("s3://") else file_path
        loaded_data[file_name] = get_df_from_excel_s3_path(full_s3_path)
        print(f"Ready to analyze: {file_name}")

    return loaded_data


def transform_atni_data(all_data: dict) -> pd.DataFrame:  # Transforming ATNi Data
    """Transforms raw ATNi data into a dataframe of nutritional differences per person, per day
    args:
        all_data(dict): ATNi summary sheets, where keys represent the food category/filename (e.g. savoury snacks.xlsx)
    returns:
        pd.DataFrame: pd.DataFrame: Aggregated category-level results and totals for all nutrients and percentiles.
    """

    summary_list = []

    drink_sheets = [  # Identifying 'drinks' requiring lower wastage
        "USA Bottled Water.xlsx",
        "USA Carbonates.xlsx",
        "USA Concentrates.xlsx",
        "USA Juice.xlsx",
        "USA Other Hot Drinks.xlsx",
        "USA RTD coffee.xlsx",
        "USA RTD tea.xlsx",
        "USA Sports Drinks.xlsx",
    ]

    for filename, df in all_data.items():
        res = {"ATNi_Data": filename}

        sales_pc = float(df.iloc[9, 1])  # Sales value

        percentages = ["25%", "50%", "75%"]  # HSR Percentiles
        wastage = (
            0.93 if filename in drink_sheets else 0.79  # Average Drink/Food waste. SENSITIVITY ANALYSIS: changed to 1
        )
        diff_pp_pd_ww = (
            (sales_pc * config.KG_CONV) / config.DAYS_PER_YEAR
        ) * wastage  # Calc: total sales per person, per day, with wastage

        for i, pct in enumerate(percentages):
            offset = 5 + (i * 9)  # Ignoring white-spaces between the data we need

            val_kcal = ((df.iloc[offset, 8] * config.KCAL_CONV) - (df.iloc[4, 1] * config.KCAL_CONV)) * diff_pp_pd_ww
            res[f"{pct}_KCAL"] = (
                val_kcal * config.COMPENSATION if val_kcal < 0 else val_kcal
            )  # Kcal. Calc: nutrient difference considering sales per person, per day, with wastage.

            val_satfat = (df.iloc[offset + 1, 8] - df.iloc[5, 1]) * diff_pp_pd_ww
            res[f"{pct}_SatFat"] = (
                val_satfat * config.COMPENSATION if val_satfat < 0 else val_satfat
            )  # SatFat. Calc: nutrient difference considering sales per person, per day, with wastage.

            val_fibre = (df.iloc[offset + 2, 8] - df.iloc[6, 1]) * diff_pp_pd_ww
            res[f"{pct}_Fibre"] = (
                val_fibre * config.COMPENSATION if val_fibre < 0 else val_fibre
            )  # Fibre. Calc: nutrient difference considering sales per person, per day, with wastage.

            val_salt = (
                (df.iloc[offset + 3, 8] * config.SALT_CONV) - (df.iloc[7, 1] * config.SALT_CONV)
            ) * diff_pp_pd_ww
            res[f"{pct}_Salt"] = (
                val_salt * config.COMPENSATION if val_salt < 0 else val_salt
            )  # Salt. Calc: nutrient difference considering sales per person, per day, with wastage.

        summary_list.append(res)

    final_report = pd.DataFrame(summary_list).set_index("ATNi_Data")  # Aggregating Data into a Single Table
    final_report.loc["TOTAL"] = final_report.sum(  # Including Totals Row
        numeric_only=True, axis=0
    )

    s3_destination = "Placeholder_URL"
    final_report.to_excel(s3_destination, index=True)

    print(f"Report generated successfully: {s3_destination}")
    return final_report

    #
    # PIPELINE B
    # Running PRIME Model with Transformed ATNi Data


def get_time() -> str:
    """Get current datetime.
    args:
        None
    returns:
        str: current date and time
    """
    return datetime.now().strftime("%B %d, %Y, %H:%M:%S")


def inject_prime_data(ws: xw.Sheet, nutrients: dict) -> None:
    """Inputs the calculated nutrient differences into the PRIME Excel sheets.
    args:
        ws: the xlwings sheet for 'Baseline & Counterfactual"
        nutrients: a dictionary containing flat values for 'kcal' 'fibre' 'salt' and 'satfat'
    returns:
        None: function modified the Excel sheet in-place
    """
    print("Adding sample data...")
    # Baseline & Counterfactual data entry
    for row in range(4, 35):
        if row == 19:
            continue
        ws.range(f"R{row}").value = (ws.range(f"C{row}").value) + nutrients["kcal"]
        ws.range(f"Y{row}").value = (ws.range(f"J{row}").value) + nutrients["fibre"]
        ws.range(f"AA{row}").value = ((ws.range(f"N{row}").value) * 0.0025) + nutrients["salt"]

    # Saturated Fat specific table
    for row in range(39, 70):
        if row == 54:
            continue
        energy_row = row - 35  # The SatFat table starts at row 39, which corresponds to row 4 in the energy table

        ws.range(f"T{row}").value = (
            ((ws.range(f"M{row}").value + nutrients["satfat"]) * 9) / ws.range(f"R{energy_row}").value * 100
        )  # noqa: E501
        # Updated to reflect % change in SatFat rather than absolute change, as per PRIME model structure


def run_prime_simulation(wb: xw.Book) -> None:
    """Triggers the MC Macro and monitors for completion.
    args:
        wb: the xlwings workbook for 'PRIME Model' including the macro
    returns:
        None: function runs the macro in the Excel workbook
    """
    macro = "NewMCAllResults5000"
    start_time = time.time()  # Start the stopwatch
    timeout_seconds = 20 * 60  # 20 minutes converted to seconds
    print(f"Running macro {get_time()}...")

    try:
        wb.macro(macro)()
    except CommandError:
        print("Timeout circumvented. Macro still processing...")

    finished = False
    while not finished:
        elapsed = time.time() - start_time

        if elapsed > timeout_seconds:
            print(f"!!! TIMEOUT: Macro exceeded 20 mins. Saving/Uploading current results {get_time()}")
            break

        time.sleep(20)  # Wait 20 seconds before checking again

        try:
            status = wb.sheets["Results"].range("BS26").value
            if status:
                finished = True
        except CommandError:
            print("Excel is busy calculating... retrying in 20s")

    print(f"Macro finished {get_time()}...")


def run_single_simulation(app: xw.App, master_path: str, nutrients: dict) -> xw.Book:
    """ "
    opens the Excel master, injects nutrient data and runs the macro
    args:
        app: the active xlwings excel
        master_path: the local file path to master_model_temp
        nutrients: dictionary with 'kcal', 'satfat', 'fibre', and 'salt'
    returns:
        xw.Book: the opened workbook after the macro and calculations finish
    """
    print("Opening workbook...")
    wb = app.books.open(master_path, update_links=False, read_only=False)
    ws = wb.sheets["Baseline & Counterfactual"]

    # Run simulation steps
    inject_prime_data(ws, nutrients)
    run_prime_simulation(wb)

    print("Calculating results...")
    app.calculate()
    time.sleep(20)
    return wb


def save_and_upload_results(wb: xw.Book, category: str, pct: str, output_path: str) -> None:
    """Saves the local workbook, uploads to S3, and cleans up.
    args:
        wb (xw.Book): The active xlwings workbook object containing the simulation results.
        category (str): The name of the food category (e.g., 'savoury snacks.xlsx').
        pct (str): The HSR percentile string (e.g., '50%').
        output_path (str): The S3 destination directory path.
    returns:
        None: function performs saving/uploading/deleting operations
    """
    fs = s3fs.S3FileSystem()
    clean_name = str(category).replace(" ", "_").replace(".xlsx", "")

    local_filename = os.path.join(os.getcwd(), f"{clean_name}_{pct.replace('%', 'pc')}.xlsm")
    print(f"Saving results to {local_filename}...")
    wb.save(local_filename)
    fs.put(local_filename, f"{output_path}{os.path.basename(local_filename)}")

    wb.close()
    if os.path.exists(local_filename):
        os.remove(local_filename)


if __name__ == "__main__":
    fs = s3fs.S3FileSystem()
    S3_PATH = "Placeholder_URL"
    all_shares = load_s3_data(S3_PATH)
    report = transform_atni_data(all_shares)

    # S3 Paths
    S3_MODEL = "Placeholder_URL"
    S3_OUTPUT_PATH = "Placeholder_URL"
    S3_FINAL_RESULTS_DIR = "Placeholder_URL"

    # Local File Names
    BASE_DIR = os.getcwd()
    LOCAL_MASTER = os.path.join(BASE_DIR, "master_model_temp.xlsm")

    FINAL_CSV_NAME = "PRIME_Results_Summary.csv"
    S3_FINAL_DEST = f"{S3_FINAL_RESULTS_DIR}{FINAL_CSV_NAME}"

    # Simulation Parameters
    PERCENTILES = ["25%", "50%", "75%"]
    fs.download(S3_MODEL, LOCAL_MASTER)  # Downloads S3 Model

    all_results_data = []

    # Start Excel ONCE before the loops
    app = xw.App(
        visible=False,
    )

    app.api.timeout = 3600  # to prevent timeout

    try:
        # Nested Loop: Cycles through Categories then Percentiles
        for category in report.index:
            for pct in PERCENTILES:
                print(f"\n--- Processing: {category} | {pct} | {get_time()} ---")

                nutrients = {
                    "kcal": report.loc[category, f"{pct}_KCAL"],
                    "satfat": report.loc[category, f"{pct}_SatFat"],
                    "fibre": report.loc[category, f"{pct}_Fibre"],
                    "salt": report.loc[category, f"{pct}_Salt"],
                }

                print("Opening workbook...")
                wb = run_single_simulation(app=app, master_path=LOCAL_MASTER, nutrients=nutrients)
                status = wb.sheets["Results"].range("BS26").value

                if not status:
                    print(f"!!! TIMEOUT DETECTED for {category}. Skipping extraction.")
                    data_row = {
                        "Category": category,
                        "Percentile": pct,
                        "Is_Timeout": True,
                        "Timestamp": get_time(),
                    }
                    for m in [
                        "Average Deaths Averted",
                        "Lower Bound Deaths Averted",
                        "Upper Bound Deaths Averted",
                        "Average Deaths Averted <75",
                        "Lower Bound Deaths Averted <75",
                        "Upper Bound Deaths Averted <75",
                        "Average LYS",
                        "Lower Bound LYS",
                        "Upper Bound LYS",
                        "Average Subjective VLYS $:",
                        "Lower Bound Subjective VLYS $:",
                        "Upper Bound Subjective VLYS $",
                        "Average Economic Value of LYS $:",
                        "Lower Bound Economic Value of LYS $:",
                        "Upper Bound Economic Value of LYS $:",
                        "Disease Breakdown: CVD Average Deaths Averted",
                        "Disease Breakdown: CVD Lower Bound Deaths Averted",
                        "Disease Breakdown: CVD Upper Bound Deaths Averted",
                        "Disease Breakdown: Stroke Average Deaths Averted",
                        "Disease Breakdown: Stroke Lower Bound Deaths Averted",
                        "Disease Breakdown: Stroke Upper Bound Deaths Averted",
                        "Disease Breakdown: Hypertensive Disease Average Deaths Averted",
                        "Disease Breakdown: Hypertensive Disease Lower Bound Deaths Averted",
                        "Disease Breakdown: Hypertensive Disease Upper Bound Deaths Averted",
                        "Disease Breakdown: Diabetes Average Deaths Averted",
                        "Disease Breakdown: Diabetes Lower Bound Deaths Averted",
                        "Disease Breakdown: Diabetes Upper Bound Deaths Averted",
                        "Risk Factor Breakdown: Obesity Average Deaths Averted",
                        "Risk Factor Breakdown: Obesity Lower Bound Deaths Averted",
                        "Risk Factor Breakdown: Obesity Upper Bound Deaths Averted",
                    ]:
                        data_row[m] = None
                else:
                    print("Extracting results...")
                    results_sheet = wb.sheets["Results"]
                    mc_results_sheet = wb.sheets["MC_Results"]  # Updated to capture key deaths averted metrics

                    # Capture all metrics
                    data_row = {
                        "Category": category,
                        "Percentile": pct,
                        "Average Deaths Averted": mc_results_sheet.range("C4").value,
                        "Lower Bound Deaths Averted": mc_results_sheet.range("B4").value,
                        "Upper Bound Deaths Averted": mc_results_sheet.range("D4").value,
                        "Average Deaths Averted <75": mc_results_sheet.range("C5").value,
                        "Lower Bound Deaths Averted <75": mc_results_sheet.range("B5").value,
                        "Upper Bound Deaths Averted <75": mc_results_sheet.range("D5").value,
                        "Average LYS": results_sheet.range("BS21").value,
                        "Lower Bound LYS": results_sheet.range("BT21").value,
                        "Upper Bound LYS": results_sheet.range("BU21").value,
                        "Average Subjective VLYS $": results_sheet.range("BS22").value,
                        "Lower Bound Subjective VLYS $": results_sheet.range("BT22").value,
                        "Upper Bound Subjective VLYS $": results_sheet.range("BU22").value,
                        "Average Economic Value LYS $": results_sheet.range("BS26").value,
                        "Lower Bound Economic Value LYS $": results_sheet.range("BT26").value,
                        "Upper Bound Economic Value LYS $": results_sheet.range("BU26").value,
                        "Disease Breakdown: CVD Average Deaths Averted": mc_results_sheet.range("H4").value,
                        "Disease Breakdown: CVD Lower Bound Deaths Averted": mc_results_sheet.range("G4").value,
                        "Disease Breakdown: CVD Upper Bound Deaths Averted": mc_results_sheet.range("I4").value,
                        "Disease Breakdown: Stroke Average Deaths Averted": mc_results_sheet.range("H6").value,
                        "Disease Breakdown: Stroke Lower Bound Deaths Averted": mc_results_sheet.range("G6").value,
                        "Disease Breakdown: Stroke Upper Bound Deaths Averted": mc_results_sheet.range("I6").value,
                        "Disease Breakdown: Hypertensive Disease Average Deaths Averted": mc_results_sheet.range(
                            "H11"
                        ).value,
                        "Disease Breakdown: Hypertensive Disease Lower Bound Deaths Averted": mc_results_sheet.range(
                            "G11"
                        ).value,
                        "Disease Breakdown: Hypertensive Disease Upper Bound Deaths Averted": mc_results_sheet.range(
                            "I11"
                        ).value,
                        "Disease Breakdown: Diabetes Average Deaths Averted": mc_results_sheet.range("H13").value,
                        "Disease Breakdown: Diabetes Lower Bound Deaths Averted": mc_results_sheet.range("G13").value,
                        "Disease Breakdown: Diabetes Upper Bound Deaths Averted": mc_results_sheet.range("I13").value,
                        "Risk Factor Breakdown: Obesity Average Deaths Averted": mc_results_sheet.range("M14").value,
                        "Risk Factor Breakdown: Obesity Lower Bound Deaths Averted": mc_results_sheet.range(
                            "L14"
                        ).value,
                        "Risk Factor Breakdown: Obesity Upper Bound Deaths Averted": mc_results_sheet.range(
                            "N14"
                        ).value,
                        "Timestamp": get_time(),
                    }

                all_results_data.append(data_row)

                # Save/Upload the file and close the workbook
                save_and_upload_results(wb=wb, category=category, pct=pct, output_path=S3_OUTPUT_PATH)

                # Incremental CSV backup to S3
                pd.DataFrame(all_results_data).to_csv(S3_FINAL_DEST.replace(".csv", "_TEMP_BACKUP.csv"), index=False)  # noqa: E501
                print("TEMP backup saved to S3.")

                print(f"Summary table updated on S3: {S3_FINAL_DEST}")
                final_df = pd.DataFrame(all_results_data)
                final_df.to_csv(S3_FINAL_DEST, index=False)
                print(f"Successfully saved to: {S3_FINAL_DEST}")

    except Exception as e:
        print(f"Saving Temp Backup incase of crash: {e}")
        if all_results_data:
            TEMP_dest = S3_FINAL_DEST.replace(".csv", "_TEMP_BACKUP.csv")
            with fs.open(TEMP_dest, "w") as f:
                pd.DataFrame(all_results_data).to_csv(f, index=False)

    finally:
        # Final Cleanup
        print("Shutting down Excel and removing local master...")
        app.quit()
        if os.path.exists(LOCAL_MASTER):
            os.remove(LOCAL_MASTER)
        print(f"All processing complete at {get_time()}")
