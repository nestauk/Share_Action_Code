import os

import openpyxl
import pandas as pd
import s3fs
from scipy.stats import spearmanr

from ahl_share_action_modelling.getters.base_getters import get_df_from_excel_s3_path


# Function to load Company Data
def loading_company_data() -> pd.DataFrame:
    """Loads dataframe of ATNi's Company Sales Data, from a hardcoded s3 path
    args:
        none
    returns:
        pd.DataFrame of Company Sales Data
    """
    file_path = "Placeholder_URL"
    file_name = "Placeholder_Name"
    df = get_df_from_excel_s3_path(file_path, sheet_name="Sheet1")
    print(f"Ready: {file_name}")
    return df


# Function to load Product Data
def loading_product_data() -> pd.DataFrame:
    """Loads dataframe of ATNis Product Data, from a hardcoded s3 path
    args:
        none
    returns:
        pd.DataFrame of Product Sales Data
    """
    file_path = "Placeholder_UR"
    file_name = "Placeholder_Name"
    df = get_df_from_excel_s3_path(file_path)
    print(f"Ready:{file_name}")
    return df


if __name__ == "__main__":
    company_df = loading_company_data()
    product_df = loading_product_data()

    # Updating column names to be more readable
    product_df = product_df.rename(
        columns={
            "emi": "Food/Drink Category",
            "manuf": "Company Name",
            "hsr": "Healthy Star Rating",
        }
    )

    ### TESTING CORRELATION - HSR score against: Kcal, SatFat, Fibre, Sodium.
    # The most strongly correlated is used in the ranking below, where HSR scores are equal

    print("\n" + "=" * 60)
    print("CORRELATION ANALYSIS: Healthy Star Rating vs Nutrients")
    print("=" * 60)

    target_var = "Healthy Star Rating"
    nutrients = ["kj", "fibre", "sod_mg", "satfat"]

    # Running Spearman's Correlation
    for nutrient in nutrients:
        clean_df = product_df[[target_var, nutrient]].dropna()
        if len(clean_df) > 30:
            corr_coeff, p_value = spearmanr(clean_df[target_var], clean_df[nutrient])

            print(f"\n* {target_var} vs {nutrient.upper()}:")
            print(f"  Correlation Coefficient (r) : {corr_coeff:.4f}")
            print(f"  p-Value                     : {p_value:.4e}")

            # Interpreting significance
            if p_value < 0.05:
                print("  Significance                : Statistically Significant (p < 0.05)")
            else:
                print("  Significance                : NOT Statistically Significant (p >= 0.05)")
        else:
            print(f"\n* {target_var} vs {nutrient.upper()}: Not enough complete data rows to compute correlation.")

    print("\n" + "=" * 60 + "\n")

    ##
    # # Ranking product data. Per category, products ranked from high-low HSR. Where equal, products ranked low-high kj
    ranked_product_data = product_df.sort_values(
        by=["Food/Drink Category", "Healthy Star Rating", "kj"],
        ascending=[True, False, True],
    )

    ###
    # # Sales Weighting Step 1: Estimating Product Sales
    # Product Sales = Company Sales / Number of products in a company

    # Splitting products by company name/category
    number_of_products_per_company_category = ranked_product_data.copy()
    number_of_products_per_company_category["product_count"] = number_of_products_per_company_category.groupby(
        ["Company Name", "Food/Drink Category"]
    )["atni_id"].transform("nunique")

    # Taking sales data from company df, merging into a table
    product_table_terms = ["Company Name", "Food/Drink Category"]
    company_table_terms = [
        "company",
        "category",
    ]

    estimated_company_sales = number_of_products_per_company_category.merge(
        company_df[["company", "category", "company_sales_cat (in USD Million)"]],
        left_on=product_table_terms,
        right_on=company_table_terms,
        how="left",
    )

    # Calculating Sales per product in USD Million: (company sales / number of products per company in a category)
    estimated_company_sales["Estimated Product Sales (USD Million)"] = (
        estimated_company_sales["company_sales_cat (in USD Million)"] / estimated_company_sales["product_count"]
    )

    ###
    # Sales Weighting Step 2: Scenario-by-Scenario
    # Calculation = Product information * product sales / sum (product sales within category / scenario)

    # Calculates group sizes/item ranks per category, for calculating the scenarios
    category_counts = estimated_company_sales.groupby("Food/Drink Category")["atni_id"].transform("nunique")
    row_position_in_category = estimated_company_sales.groupby("Food/Drink Category").cumcount()

    # Creating 3 scenarios: Top 25%, tOP 50%, tOP 75%
    top_25_group = estimated_company_sales[row_position_in_category <= (category_counts * 0.25)].copy()
    top_50_group = estimated_company_sales[row_position_in_category <= (category_counts * 0.50)].copy()
    top_75_group = estimated_company_sales[row_position_in_category <= (category_counts * 0.75)].copy()

    # Testing the scenarios contain the correct proportion of products
    print(f"Total Rows: {len(estimated_company_sales)}")
    print(f"Top 25% contains: {len(top_25_group)} rows")
    print(f"Top 50% contains: {len(top_50_group)} rows")
    print(f"Top 75% contains: {len(top_75_group)} rows")

    ###
    # Top 25% Scenario
    nutrition_cols = ["fibre", "kj", "sod_mg", "satfat"]
    top_25_dataset = top_25_group.copy()

    # Calculates the total sales for each category
    category_sales_sums = top_25_dataset.groupby("Food/Drink Category")[
        "Estimated Product Sales (USD Million)"
    ].transform("sum")  # noqa: E501

    # Applies the weighting calculation to every nutrient column
    category_weighted_df = top_25_dataset.copy()
    for col in nutrition_cols:
        category_weighted_df[col] = (
            category_weighted_df[col] * category_weighted_df["Estimated Product Sales (USD Million)"]
        ) / category_sales_sums

    # Group by Category and SUM the weighted values to get the 'sales weighted average'
    final_category_summary = category_weighted_df.groupby("Food/Drink Category")[nutrition_cols].sum().reset_index()

    print("--- Top 25% Scenario Broken Down by Category ---")
    print(final_category_summary)

    # Uploading Scenario to S3
    fs = s3fs.S3FileSystem()
    s3_input_dir = "s3://ahl-shareaction-private/Input/"

    category_to_file = {
        "Baked Goods": "USA Baked Goods.xlsx",
        "Bottled Water": "USA Bottled Water.xlsx",
        "Breakfast Cereals": "USA Breakfast Cereals.xlsx",
        "Carbonates": "USA Carbonates.xlsx",
        "Concentrates": "USA Concentrates.xlsx",
        "Confectionery": "USA Confectionary.xlsx",
        "Ice Cream": "USA Ice Cream.xlsx",
        "Juice": "USA Juice.xlsx",
        "Other Hot Drinks": "USA Other Hot Drinks.xlsx",
        "Processed Fruit and Vegetables": "USA Processed Fruit and Vegetables.xlsx",
        "Rice, Pasta and Noodles": "USA Rice, Pasta and Noodles.xlsx",
        "RTD Coffee": "USA RTD Coffee.xlsx",
        "RTD Tea": "USA RTD Tea.xlsx",
        "Sauces, Dips and Condiments": "USA Sauces, Dips and Condiments.xlsx",
        "Savoury Snacks": "USA Savoury Snacks.xlsx",
        "Sports Drinks": "USA Sports Drinks.xlsx",
        "Sweet Biscuits, Snack Bars and Fruit Snacks": "USA Sweet Biscuits, Snack Bars and Fruit Snacks.xlsx",
        "Sweet Spreads": "USA Sweet Spreads.xlsx",
    }

    # Load sales-weighted score into each category Excel (used in the model)
    for _index, row in final_category_summary.iterrows():
        file_name = category_to_file.get(row["Food/Drink Category"])
        if not file_name:
            continue
        full_s3_path = f"{s3_input_dir}{file_name}"
        local_temp_path = f"/tmp/{file_name}" if os.name != "nt" else f"C:\\TEMP\\{file_name}"

        print(f"\n[1/4] Downloading Excel {file_name}")
        try:
            fs.download(full_s3_path, local_temp_path)

            # Adding data to correct cells
            wb = openpyxl.load_workbook(local_temp_path, keep_vba=False)
            ws = wb.active
            ws["H4"] = "Top 25%"  # Updating Sheet for clarity
            ws["H3"] = ""  # Removing Worksheet Text as Not Needed
            ws["H12"] = ""
            ws["H21"] = ""
            for row_num in range(30, 47):
                ws[f"H{row_num}"] = ""

            print("[2/4] Injecting data...")
            ws["I7"] = row["kj"]
            ws["I8"] = row["satfat"]
            ws["I9"] = row["fibre"]
            ws["I10"] = row["sod_mg"]

            # Save/Close
            wb.save(local_temp_path)
            wb.close()

            print("[3/4] Uploading to S3...")
            fs.upload(local_temp_path, full_s3_path)
            print(f"[4/4] Successfully completed: {file_name}")

            # Removing Temp
            if os.path.exists(local_temp_path):
                os.remove(local_temp_path)

        except Exception as e:
            print(f"CRITICAL ERROR on {file_name}: {e}")
            if os.path.exists(local_temp_path):
                os.remove(local_temp_path)

    ###
    # Top 50% Scenario
    nutrition_cols = ["fibre", "kj", "sod_mg", "satfat"]
    top_50_dataset = top_50_group.copy()

    # Calculate the total sales for each category within this top 50% group
    category_sales_sums = top_50_dataset.groupby("Food/Drink Category")[
        "Estimated Product Sales (USD Million)"
    ].transform("sum")  # noqa: E501

    # Apply the weights using each category's specific total sales
    category_weighted_df = top_50_dataset.copy()
    for col in nutrition_cols:
        category_weighted_df[col] = (
            category_weighted_df[col] * category_weighted_df["Estimated Product Sales (USD Million)"]
        ) / category_sales_sums

    # Group by Category and SUM the weighted values to get the 'sales weighted average'
    final_category_summary = category_weighted_df.groupby("Food/Drink Category")[nutrition_cols].sum().reset_index()

    print("--- Top 50% Scenario Broken Down by Category ---")
    print(final_category_summary)

    # Uploading Scenario to S3
    fs = s3fs.S3FileSystem()
    s3_input_dir = "s3://ahl-shareaction-private/Input/"

    category_to_file = {
        "Baked Goods": "USA Baked Goods.xlsx",
        "Bottled Water": "USA Bottled Water.xlsx",
        "Breakfast Cereals": "USA Breakfast Cereals.xlsx",
        "Carbonates": "USA Carbonates.xlsx",
        "Concentrates": "USA Concentrates.xlsx",
        "Confectionery": "USA Confectionary.xlsx",
        "Ice Cream": "USA Ice Cream.xlsx",
        "Juice": "USA Juice.xlsx",
        "Other Hot Drinks": "USA Other Hot Drinks.xlsx",
        "Processed Fruit and Vegetables": "USA Processed Fruit and Vegetables.xlsx",
        "Rice, Pasta and Noodles": "USA Rice, Pasta and Noodles.xlsx",
        "RTD Coffee": "USA RTD Coffee.xlsx",
        "RTD Tea": "USA RTD Tea.xlsx",
        "Sauces, Dips and Condiments": "USA Sauces, Dips and Condiments.xlsx",
        "Savoury Snacks": "USA Savoury Snacks.xlsx",
        "Sports Drinks": "USA Sports Drinks.xlsx",
        "Sweet Biscuits, Snack Bars and Fruit Snacks": "USA Sweet Biscuits, Snack Bars and Fruit Snacks.xlsx",
        "Sweet Spreads": "USA Sweet Spreads.xlsx",
    }

    # Load sales-weighted score into each category Excel (used in the model)
    for _index, row in final_category_summary.iterrows():
        file_name = category_to_file.get(row["Food/Drink Category"])
        if not file_name:
            continue
        full_s3_path = f"{s3_input_dir}{file_name}"

        local_temp_path = f"/tmp/{file_name}" if os.name != "nt" else f"C:\\TEMP\\{file_name}"

        print(f"\n[1/4] Downloading Excel {file_name}")
        try:
            fs.download(full_s3_path, local_temp_path)

            # Adding data to correct cells
            wb = openpyxl.load_workbook(local_temp_path, keep_vba=False)
            ws = wb.active
            ws["H13"] = "Top 50%"

            print("[2/4] Injecting data...")
            ws["I16"] = row["kj"]
            ws["I17"] = row["satfat"]
            ws["I18"] = row["fibre"]
            ws["I19"] = row["sod_mg"]

            # Save/Close
            wb.save(local_temp_path)
            wb.close()

            print("[3/4] Uploading to S3...")
            fs.upload(local_temp_path, full_s3_path)
            print(f"[4/4] Successfully completed: {file_name}")

            # Removing temp file
            if os.path.exists(local_temp_path):
                os.remove(local_temp_path)

        except Exception as e:
            print(f"CRITICAL ERROR on {file_name}: {e}")
            if os.path.exists(local_temp_path):
                os.remove(local_temp_path)

    ###
    # Top 75% Scenario
    nutrition_cols = ["fibre", "kj", "sod_mg", "satfat"]
    top_75_dataset = top_75_group.copy()

    # Calculate the total sales for each category within this top 75% group
    category_sales_sums = top_75_dataset.groupby("Food/Drink Category")[
        "Estimated Product Sales(USD Million)"
    ].transform("sum")  # noqa: E501

    # Apply the weights using each category's specific total sales
    category_weighted_df = top_75_dataset.copy()
    for col in nutrition_cols:
        category_weighted_df[col] = (
            category_weighted_df[col] * category_weighted_df["Estimated Product Sales (USD Million)"]
        ) / category_sales_sums

    # Group by Category and SUM the weighted values to get the 'sales weighted average''
    final_category_summary = category_weighted_df.groupby("Food/Drink Category")[nutrition_cols].sum().reset_index()

    print("--- Top 75% Scenario Broken Down by Category ---")
    print(final_category_summary)

    # Uploading Scenario-data to AWS
    fs = s3fs.S3FileSystem()
    s3_input_dir = "s3://ahl-shareaction-private/Input/"

    # Matching column names to the correct file name in AWS
    category_to_file = {
        "Baked Goods": "USA Baked Goods.xlsx",
        "Bottled Water": "USA Bottled Water.xlsx",
        "Breakfast Cereals": "USA Breakfast Cereals.xlsx",
        "Carbonates": "USA Carbonates.xlsx",
        "Concentrates": "USA Concentrates.xlsx",
        "Confectionery": "USA Confectionary.xlsx",
        "Ice Cream": "USA Ice Cream.xlsx",
        "Juice": "USA Juice.xlsx",
        "Other Hot Drinks": "USA Other Hot Drinks.xlsx",
        "Processed Fruit and Vegetables": "USA Processed Fruit and Vegetables.xlsx",
        "Rice, Pasta and Noodles": "USA Rice, Pasta and Noodles.xlsx",
        "RTD Coffee": "USA RTD Coffee.xlsx",
        "RTD Tea": "USA RTD Tea.xlsx",
        "Sauces, Dips and Condiments": "USA Sauces, Dips and Condiments.xlsx",
        "Savoury Snacks": "USA Savoury Snacks.xlsx",
        "Sports Drinks": "USA Sports Drinks.xlsx",
        "Sweet Biscuits, Snack Bars and Fruit Snacks": "USA Sweet Biscuits, Snack Bars and Fruit Snacks.xlsx",
        "Sweet Spreads": "USA Sweet Spreads.xlsx",
    }

    # Load sales-weighted score into each category Excel (used in the model)
    for _index, row in final_category_summary.iterrows():
        file_name = category_to_file.get(row["Food/Drink Category"])
        if not file_name:
            continue
        full_s3_path = f"{s3_input_dir}{file_name}"

        local_temp_path = f"/tmp/{file_name}" if os.name != "nt" else f"C:\\TEMP\\{file_name}"

        print(f"\n[1/4] Downloading Excel: {file_name}")
        try:
            fs.download(full_s3_path, local_temp_path)

            # Opening file, adding data to correct cells, and saving to AWS
            wb = openpyxl.load_workbook(local_temp_path, keep_vba=False)
            ws = wb.active
            ws["H22"] = "Top 75%"

            print("[2/4] Injecting data...")
            ws["I25"] = row["kj"]
            ws["I26"] = row["satfat"]
            ws["I27"] = row["fibre"]
            ws["I28"] = row["sod_mg"]

            wb.save(local_temp_path)
            wb.close()

            print("[3/4] Uploading file to S3...")
            fs.upload(local_temp_path, full_s3_path)
            print(f"[4/4] Successfully completed: {file_name}")

            # Removing temp file
            if os.path.exists(local_temp_path):
                os.remove(local_temp_path)

        except Exception as e:
            print(f"CRITICAL ERROR on {file_name}: {e}")
            if os.path.exists(local_temp_path):
                os.remove(local_temp_path)
