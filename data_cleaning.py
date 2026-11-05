
# ============================================================
# RETAILPULSE - ONLINE RETAIL II DATA CLEANING PROJECT
# Input : online_retail_II.xlsx
# Output: Cleaned CSV, Excel workbook, cleaning summary
# ============================================================

from pathlib import Path
import pandas as pd

# ------------------------------------------------------------
# 1. CONFIGURATION
# ------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "online_retail_II.xlsx"
OUTPUT_DIR = BASE_DIR / "03_Cleaned_Data"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

CSV_FILE = OUTPUT_DIR / "RetailPulse_Cleaned_Data.csv"
EXCEL_FILE = OUTPUT_DIR / "RetailPulse_Cleaned_Data.xlsx"
SUMMARY_FILE = OUTPUT_DIR / "RetailPulse_Cleaning_Summary.xlsx"

SHEETS_TO_READ = [
    "Year 2009-2010",
    "Year 2010-2011"
]

# Excel allows a maximum of 1,048,576 rows per worksheet.
# Keep the header row within this limit.
EXCEL_DATA_ROWS_PER_SHEET = 1_000_000


# ------------------------------------------------------------
# 2. READ AND COMBINE EXCEL SHEETS
# ------------------------------------------------------------

def load_data(file_path):
    if not file_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {file_path}\n"
            "Keep online_retail_II.xlsx in the same folder "
            "as this Python script, or update INPUT_FILE."
        )

    print("\nReading Excel workbook...")

    excel_file = pd.ExcelFile(file_path)
    available_sheets = excel_file.sheet_names

    dataframes = []

    for sheet in SHEETS_TO_READ:
        if sheet not in available_sheets:
            raise ValueError(
                f"Sheet '{sheet}' was not found. "
                f"Available sheets: {available_sheets}"
            )

        df = pd.read_excel(
            excel_file,
            sheet_name=sheet
        )

        df.columns = df.columns.str.strip()
        df["Source_Sheet"] = sheet

        print(f"{sheet}: {len(df):,} rows loaded")

        dataframes.append(df)

    combined_df = pd.concat(
        dataframes,
        ignore_index=True
    )

    print(f"\nCombined rows: {len(combined_df):,}")

    return combined_df


# ------------------------------------------------------------
# 3. CLEAN COLUMN NAMES AND TEXT VALUES
# ------------------------------------------------------------

def clean_text_and_columns(df):
    print("\nCleaning column names and text...")

    df.columns = (
        df.columns
        .str.strip()
    )

    # Standardize the customer ID column name
    df = df.rename(
        columns={"Customer ID": "Customer_ID"}
    )

    # Clean text fields without converting missing values
    # into the literal string "nan".
    text_columns = [
        "Invoice",
        "StockCode",
        "Description",
        "Country",
        "Source_Sheet"
    ]

    for col in text_columns:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype("string")
                .str.strip()
            )

            df[col] = df[col].replace(
                "", pd.NA
            )

    return df


# ------------------------------------------------------------
# 4. CONVERT DATA TYPES
# ------------------------------------------------------------

def convert_data_types(df):
    print("\nConverting data types...")

    df["Invoice"] = df["Invoice"].astype("string")
    df["StockCode"] = df["StockCode"].astype("string")

    df["Description"] = df["Description"].astype("string")
    df["Country"] = df["Country"].astype("string")

    df["InvoiceDate"] = pd.to_datetime(
        df["InvoiceDate"],
        errors="coerce"
    )

    df["Quantity"] = pd.to_numeric(
        df["Quantity"],
        errors="coerce"
    )

    df["Price"] = pd.to_numeric(
        df["Price"],
        errors="coerce"
    )

    df["Customer_ID"] = pd.to_numeric(
        df["Customer_ID"],
        errors="coerce"
    )

    # Customer IDs are identifiers, not quantities.
    # Keep missing IDs as missing.
    df["Customer_ID"] = df["Customer_ID"].astype(
        "Int64"
    )

    return df


# ------------------------------------------------------------
# 5. REMOVE DUPLICATES AND INVALID RECORDS
# ------------------------------------------------------------

def remove_duplicates_and_invalid_rows(df):
    print("\nRemoving duplicates and invalid records...")

    original_rows = len(df)

    # Exact duplicate check across all source and derived
    # input columns currently available.
    duplicate_count = int(df.duplicated().sum())

    df = df.drop_duplicates().copy()

    # Remove records without essential transaction fields.
    # Customer_ID and Description are allowed to be missing.
    required_columns = [
        "Invoice",
        "StockCode",
        "InvoiceDate",
        "Quantity",
        "Price"
    ]

    before_required = len(df)

    df = df.dropna(
        subset=required_columns
    ).copy()

    missing_required_removed = (
        before_required - len(df)
    )

    # Quantity must be non-zero for a normal sales or return
    # line. Negative quantities are retained as returns.
    zero_quantity_count = int(
        (df["Quantity"] == 0).sum()
    )

    df = df[df["Quantity"] != 0].copy()

    # Negative prices are treated as invalid unit prices.
    # Zero prices are retained because free items may exist.
    negative_price_count = int(
        (df["Price"] < 0).sum()
    )

    df = df[df["Price"] >= 0].copy()

    # Add a review flag for missing customer IDs.
    # Do not remove these rows because they still contain
    # potentially useful sales information.
    missing_customer_count = int(
        df["Customer_ID"].isna().sum()
    )

    # Fill missing descriptions and country labels so that
    # they remain visible in Power BI.
    df["Description"] = df["Description"].fillna(
        "Unknown Product"
    )

    df["Country"] = df["Country"].fillna(
        "Unknown"
    )

    print(f"Initial rows: {original_rows:,}")
    print(f"Duplicate rows removed: {duplicate_count:,}")
    print(
        "Missing required fields removed: "
        f"{missing_required_removed:,}"
    )
    print(f"Zero-quantity rows removed: {zero_quantity_count:,}")
    print(f"Negative-price rows removed: {negative_price_count:,}")
    print(
        "Rows with missing customer ID retained: "
        f"{missing_customer_count:,}"
    )
    print(f"Rows after cleaning: {len(df):,}")

    cleaning_counts = {
        "Initial Rows": original_rows,
        "Duplicate Rows Removed": duplicate_count,
        "Missing Required Fields Removed": missing_required_removed,
        "Zero Quantity Rows Removed": zero_quantity_count,
        "Negative Price Rows Removed": negative_price_count,
        "Rows with Missing Customer ID": missing_customer_count,
        "Final Cleaned Rows": len(df)
    }

    return df, cleaning_counts


# ------------------------------------------------------------
# 6. CREATE ANALYTICS COLUMNS
# ------------------------------------------------------------

def create_analytics_columns(df):
    print("\nCreating analytics columns...")

    # Transaction line revenue.
    # Negative quantities create negative revenue lines,
    # which can represent returns or credit adjustments.
    df["Revenue"] = (
        df["Quantity"] * df["Price"]
    ).round(2)

    # Date attributes for Power BI reporting.
    df["Year"] = df["InvoiceDate"].dt.year
    df["Month_Number"] = df["InvoiceDate"].dt.month

    df["Month_Name"] = (
        df["InvoiceDate"].dt.strftime("%B")
    )

    df["Quarter"] = (
        "Q" + df["InvoiceDate"].dt.quarter.astype(str)
    )

    df["Year_Month"] = (
        df["InvoiceDate"].dt.strftime("%Y-%m")
    )

    # Invoice IDs beginning with C are commonly used for
    # cancellations/credit transactions in this dataset.
    df["Is_Cancelled"] = (
        df["Invoice"].str.upper().str.startswith("C", na=False)
    )

    # Useful for separate sales and returns analysis.
    df["Transaction_Type"] = "Sale"

    df.loc[
        df["Quantity"] < 0,
        "Transaction_Type"
    ] = "Return"

    df.loc[
        df["Is_Cancelled"],
        "Transaction_Type"
    ] = "Cancellation/Credit"

    # Mark transactions where customer information is absent.
    df["Customer_ID_Missing"] = (
        df["Customer_ID"].isna()
    )

    # Sort data by transaction date.
    df = df.sort_values(
        by=["InvoiceDate", "Invoice", "StockCode"],
        kind="stable"
    ).reset_index(drop=True)

    return df


# ------------------------------------------------------------
# 7. CREATE DATA QUALITY SUMMARY
# ------------------------------------------------------------

def create_quality_summary(df, cleaning_counts):
    print("\nCreating data quality summary...")

    missing_summary = pd.DataFrame({
        "Column": df.columns,
        "Missing_Count": [
            int(df[col].isna().sum())
            for col in df.columns
        ],
        "Missing_Percentage": [
            round(
                df[col].isna().mean() * 100, 2
            )
            for col in df.columns
        ]
    })

    cleaning_summary = pd.DataFrame({
        "Metric": list(cleaning_counts.keys()),
        "Value": list(cleaning_counts.values())
    })

    country_summary = (
        df.groupby("Country", dropna=False)
        .agg(
            Transaction_Lines=("Invoice", "size"),
            Total_Quantity=("Quantity", "sum"),
            Total_Revenue=("Revenue", "sum")
        )
        .reset_index()
        .sort_values(
            "Total_Revenue",
            ascending=False
        )
    )

    monthly_summary = (
        df.groupby("Year_Month", dropna=False)
        .agg(
            Transaction_Lines=("Invoice", "size"),
            Total_Quantity=("Quantity", "sum"),
            Total_Revenue=("Revenue", "sum")
        )
        .reset_index()
        .sort_values("Year_Month")
    )

    return (
        missing_summary,
        cleaning_summary,
        country_summary,
        monthly_summary
    )


# ------------------------------------------------------------
# 8. SAVE CLEANED DATA TO CSV
# ------------------------------------------------------------

def save_csv(df, file_path):
    print("\nSaving cleaned CSV...")

    df.to_csv(
        file_path,
        index=False,
        encoding="utf-8-sig",
        date_format="%Y-%m-%d %H:%M:%S"
    )

    print(f"CSV saved: {file_path}")


# ------------------------------------------------------------
# 9. SAVE CLEANED DATA AND SUMMARIES TO EXCEL
# ------------------------------------------------------------

def save_excel(
    df,
    file_path,
    missing_summary,
    cleaning_summary,
    country_summary,
    monthly_summary
):
    print("\nSaving cleaned Excel workbook...")

    with pd.ExcelWriter(
        file_path,
        engine="openpyxl"
    ) as writer:

        # Split the cleaned data into worksheets to respect
        # Excel's maximum worksheet row limit.
        total_rows = len(df)
        number_of_parts = max(
            1,
            (
                total_rows
                + EXCEL_DATA_ROWS_PER_SHEET - 1
            ) // EXCEL_DATA_ROWS_PER_SHEET
        )

        for part in range(number_of_parts):
            start = part * EXCEL_DATA_ROWS_PER_SHEET
            end = min(
                start + EXCEL_DATA_ROWS_PER_SHEET,
                total_rows
            )

            chunk = df.iloc[start:end]

            sheet_name = (
                "Cleaned_Data"
                if number_of_parts == 1
                else f"Cleaned_Data_{part + 1}"
            )

            chunk.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            print(
                f"Excel sheet {sheet_name}: "
                f"{len(chunk):,} rows"
            )

        missing_summary.to_excel(
            writer,
            sheet_name="Missing_Values",
            index=False
        )

        cleaning_summary.to_excel(
            writer,
            sheet_name="Cleaning_Summary",
            index=False
        )

        country_summary.to_excel(
            writer,
            sheet_name="Country_Summary",
            index=False
        )

        monthly_summary.to_excel(
            writer,
            sheet_name="Monthly_Summary",
            index=False
        )

    print(f"Excel saved: {file_path}")


# ------------------------------------------------------------
# 10. MAIN EXECUTION
# ------------------------------------------------------------

def main():
    try:
        # Load workbook
        df = load_data(INPUT_FILE)

        # Clean text and column names
        df = clean_text_and_columns(df)

        # Convert data types
        df = convert_data_types(df)

        # Remove duplicates and invalid records
        df, cleaning_counts = (
            remove_duplicates_and_invalid_rows(df)
        )

        # Create analytics fields
        df = create_analytics_columns(df)

        # Create quality and business summaries
        (
            missing_summary,
            cleaning_summary,
            country_summary,
            monthly_summary
        ) = create_quality_summary(
            df,
            cleaning_counts
        )

        # Export full cleaned dataset to CSV
        save_csv(df, CSV_FILE)

        # Export cleaned data and summaries to Excel
        save_excel(
            df,
            EXCEL_FILE,
            missing_summary,
            cleaning_summary,
            country_summary,
            monthly_summary
        )

        # Also save the summary report separately
        with pd.ExcelWriter(
            SUMMARY_FILE,
            engine="openpyxl"
        ) as writer:
            cleaning_summary.to_excel(
                writer,
                sheet_name="Cleaning_Summary",
                index=False
            )

            missing_summary.to_excel(
                writer,
                sheet_name="Missing_Values",
                index=False
            )

            country_summary.to_excel(
                writer,
                sheet_name="Country_Summary",
                index=False
            )

            monthly_summary.to_excel(
                writer,
                sheet_name="Monthly_Summary",
                index=False
            )

        print("\n" + "=" * 60)
        print("RETAILPULSE DATA CLEANING COMPLETED")
        print("=" * 60)
        print(f"Final rows: {len(df):,}")
        print(f"Final columns: {len(df.columns)}")
        print(f"\nCleaned CSV: {CSV_FILE}")
        print(f"Cleaned Excel: {EXCEL_FILE}")
        print(f"Summary report: {SUMMARY_FILE}")

        print("\nSample cleaned data:")
        print(df.head(5).to_string(index=False))

    except Exception as error:
        print("\nDATA CLEANING FAILED")
        print(f"Error: {error}")
        raise


if __name__ == "__main__":
    main()
