
# ============================================================
# RETAILPULSE PROJECT
# PAGE 10: DEMAND FORECAST
# Python + Power BI
# ============================================================

import pandas as pd
import numpy as np
from pathlib import Path

# ============================================================
# STEP 1: FILE PATHS AND SETTINGS
# ============================================================

# Existing cleaned CSV file
INPUT_FILE = Path(
    r"C:\RETAILPULSE\03_Cleaned_Data\RetailPulse_Cleaned_Data.csv"
)

# Folder to save all forecast datasets
OUTPUT_FOLDER = Path(
    r"C:\RETAILPULSE\Demand_Forecast_Output"
)

# Automatically create output folder
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# Forecast settings
TOP_N_PRODUCTS = 20
TEST_PERCENTAGE = 0.20
LOOKBACK_DAYS = 7
FUTURE_DAYS = 30

print("=" * 60)
print("RETAILPULSE - DEMAND FORECAST")
print("=" * 60)

# ============================================================
# STEP 2: CHECK INPUT FILE
# ============================================================

print("\nSTEP 1: Checking cleaned dataset...")

if not INPUT_FILE.exists():

    raise FileNotFoundError(
        "\nCleaned dataset not found!\n"
        f"Expected location: {INPUT_FILE}\n"
        "Please verify the file name and folder."
    )

print("Input file found successfully!")

# ============================================================
# STEP 3: LOAD CLEANED CSV DATA
# ============================================================

print("\nSTEP 2: Loading cleaned CSV data...")

try:
    df = pd.read_csv(
        INPUT_FILE,
        encoding="utf-8-sig",
        low_memory=False
    )

except UnicodeDecodeError:
    df = pd.read_csv(
        INPUT_FILE,
        encoding="latin1",
        low_memory=False
    )

print("Dataset loaded successfully!")
print("Total records:", len(df))

# ============================================================
# STEP 4: STANDARDIZE COLUMN NAMES
# ============================================================

df.columns = (
    df.columns.astype(str)
    .str.strip()
)

print("\nAvailable dataset columns:")
print(df.columns.tolist())

# Create a lookup for different column-name formats
def normalize_column(name):

    return (
        str(name)
        .strip()
        .lower()
        .replace("_", "")
        .replace(" ", "")
        .replace("-", "")
    )


column_lookup = {
    normalize_column(col): col
    for col in df.columns
}


def find_column(possible_names):

    for name in possible_names:

        key = normalize_column(name)

        if key in column_lookup:
            return column_lookup[key]

    return None


# Detect required columns
DATE_COL = find_column([
    "InvoiceDate",
    "Invoice Date",
    "Invoice_Date",
    "Date",
    "TransactionDate"
])

PRODUCT_COL = find_column([
    "Description",
    "Product",
    "ProductName",
    "Product Name",
    "Item"
])

QUANTITY_COL = find_column([
    "Quantity",
    "Qty",
    "Units",
    "UnitsSold"
])

# Optional cancellation column
CANCEL_COL = find_column([
    "Is_Cancelled",
    "Is Cancelled",
    "Cancelled",
    "IsCancelled"
])

# ============================================================
# STEP 5: VALIDATE REQUIRED COLUMNS
# ============================================================

print("\nSTEP 3: Validating columns...")

if DATE_COL is None:

    raise ValueError(
        "Invoice date column was not found.\n"
        f"Available columns: {df.columns.tolist()}"
    )

if PRODUCT_COL is None:

    raise ValueError(
        "Product or Description column was not found.\n"
        f"Available columns: {df.columns.tolist()}"
    )

if QUANTITY_COL is None:

    raise ValueError(
        "Quantity column was not found.\n"
        f"Available columns: {df.columns.tolist()}"
    )

print("Date column:", DATE_COL)
print("Product column:", PRODUCT_COL)
print("Quantity column:", QUANTITY_COL)
print("Cancellation column:", CANCEL_COL)

# ============================================================
# STEP 6: DATA PREPARATION
# ============================================================

print("\nSTEP 4: Preparing data...")

# Convert date column
# dayfirst=True supports common UK retail date formats
df[DATE_COL] = pd.to_datetime(
    df[DATE_COL],
    errors="coerce",
    dayfirst=True
)

# Convert quantity to numeric
df[QUANTITY_COL] = pd.to_numeric(
    df[QUANTITY_COL],
    errors="coerce"
)

# Remove missing records
df = df.dropna(
    subset=[
        DATE_COL,
        PRODUCT_COL,
        QUANTITY_COL
    ]
).copy()

# Clean product names
df[PRODUCT_COL] = (
    df[PRODUCT_COL]
    .astype(str)
    .str.strip()
)

# Remove empty product names
df = df[
    df[PRODUCT_COL] != ""
].copy()

# Exclude cancelled transactions if available
if CANCEL_COL is not None:

    cancelled = (
        df[CANCEL_COL]
        .astype(str)
        .str.strip()
        .str.lower()
        .isin([
            "true",
            "1",
            "yes",
            "y"
        ])
    )

    df = df.loc[~cancelled].copy()

# Exclude returns and zero quantities
df = df[
    df[QUANTITY_COL] > 0
].copy()

if df.empty:

    raise ValueError(
        "No valid positive sales transactions found."
    )

# Create standard columns
df["Date"] = df[DATE_COL].dt.normalize()

df["Product"] = df[PRODUCT_COL]

df["Quantity"] = df[QUANTITY_COL]

print("Data preparation completed!")
print("Valid records:", len(df))

# ============================================================
# STEP 7: SELECT TOP 20 PRODUCTS
# ============================================================

print("\nSTEP 5: Selecting top-selling products...")

product_sales = (
    df.groupby("Product")["Quantity"]
    .sum()
    .sort_values(ascending=False)
)

top_products = (
    product_sales
    .head(TOP_N_PRODUCTS)
    .index
    .tolist()
)

df = df[
    df["Product"].isin(top_products)
].copy()

print("Selected products:", len(top_products))

# ============================================================
# STEP 8: CALCULATE DAILY PRODUCT DEMAND
# ============================================================

print("\nSTEP 6: Calculating daily product demand...")

daily_demand = (
    df.groupby(["Date", "Product"])["Quantity"]
    .sum()
    .reset_index()
)

daily_demand.rename(
    columns={
        "Quantity": "Actual_Demand"
    },
    inplace=True
)

# Create continuous date range
start_date = daily_demand["Date"].min()

end_date = daily_demand["Date"].max()

all_dates = pd.date_range(
    start=start_date,
    end=end_date,
    freq="D"
)

all_products = sorted(top_products)

# Create date-product combinations
full_index = pd.MultiIndex.from_product(
    [all_dates, all_products],
    names=["Date", "Product"]
)

# Fill missing product-date combinations with zero
daily_demand = (
    daily_demand
    .set_index(["Date", "Product"])
    .reindex(full_index, fill_value=0)
    .reset_index()
)

print("Daily demand calculated!")
print("Start date:", start_date.date())
print("End date:", end_date.date())
print("Total historical days:", len(all_dates))

# ============================================================
# STEP 9: TRAIN AND TEST SPLIT
# ============================================================

print("\nSTEP 7: Preparing historical testing...")

if len(all_dates) < 10:

    raise ValueError(
        "At least 10 days of historical data are required."
    )

split_index = int(
    len(all_dates) * (1 - TEST_PERCENTAGE)
)

split_index = max(
    LOOKBACK_DAYS,
    min(split_index, len(all_dates) - 1)
)

train_dates = all_dates[:split_index]

test_dates = all_dates[split_index:]

print("Training days:", len(train_dates))
print("Testing days:", len(test_dates))

# ============================================================
# STEP 10: HISTORICAL DEMAND FORECAST
# ============================================================

print("\nSTEP 8: Generating test forecasts...")

test_rows = []

metric_rows = []

for product in all_products:

    product_data = (
        daily_demand[
            daily_demand["Product"] == product
        ]
        .set_index("Date")["Actual_Demand"]
    )

    history = list(
        product_data.loc[train_dates].astype(float)
    )

    actual_values = []

    forecast_values = []

    for current_date in test_dates:

        # Average demand from previous 7 days
        forecast = float(
            np.mean(history[-LOOKBACK_DAYS:])
        )

        actual = float(
            product_data.loc[current_date]
        )

        error = actual - forecast

        test_rows.append({

            "Date": current_date,

            "Product": product,

            "Actual_Demand": actual,

            "Forecast_Demand": forecast,

            "Forecast_Error": error

        })

        actual_values.append(actual)

        forecast_values.append(forecast)

        # Update history with actual demand
        history.append(actual)

    # Calculate MAE
    absolute_errors = np.abs(
        np.array(actual_values)
        - np.array(forecast_values)
    )

    mae = float(
        np.mean(absolute_errors)
    )

    total_actual = float(
        np.sum(actual_values)
    )

    total_error = float(
        np.sum(absolute_errors)
    )

    # Calculate WAPE
    if total_actual > 0:

        wape = total_error / total_actual

        accuracy = max(
            0,
            1 - wape
        )

    else:

        wape = np.nan

        accuracy = np.nan

    metric_rows.append({

        "Product": product,

        "MAE": mae,

        "WAPE": wape,

        "Forecast_Accuracy": accuracy,

        "Test_Days": len(test_dates)

    })

# Create historical test dataset
test_df = pd.DataFrame(test_rows)

print("Historical test forecast completed!")

# ============================================================
# STEP 11: FUTURE 30-DAY DEMAND FORECAST
# ============================================================

print("\nSTEP 9: Generating future demand forecast...")

future_rows = []

future_dates = pd.date_range(
    start=end_date + pd.Timedelta(days=1),
    periods=FUTURE_DAYS,
    freq="D"
)

for product in all_products:

    product_data = (
        daily_demand[
            daily_demand["Product"] == product
        ]
        .set_index("Date")["Actual_Demand"]
    )

    history = list(
        product_data.astype(float)
    )

    # Baseline forecast using the last 7 days
    forecast = float(
        np.mean(history[-LOOKBACK_DAYS:])
    )

    for future_date in future_dates:

        future_rows.append({

            "Date": future_date,

            "Product": product,

            "Forecast_Demand": max(
                0,
                forecast
            )

        })

future_df = pd.DataFrame(future_rows)

print("Future forecast completed!")

# ============================================================
# STEP 12: FORECAST METRICS
# ============================================================

print("\nSTEP 10: Preparing forecast metrics...")

metrics_df = pd.DataFrame(metric_rows)

# ============================================================
# STEP 13: SAVE THREE CSV DATASETS
# ============================================================

print("\nSTEP 11: Saving CSV datasets...")

test_file = (
    OUTPUT_FOLDER /
    "Forecast_Test_Results.csv"
)

future_file = (
    OUTPUT_FOLDER /
    "Future_Demand_Forecast.csv"
)

metrics_file = (
    OUTPUT_FOLDER /
    "Forecast_Metrics.csv"
)

# Save test results
test_df.to_csv(
    test_file,
    index=False,
    encoding="utf-8-sig",
    date_format="%Y-%m-%d"
)

# Save future demand
future_df.to_csv(
    future_file,
    index=False,
    encoding="utf-8-sig",
    date_format="%Y-%m-%d"
)

# Save forecast metrics
metrics_df.to_csv(
    metrics_file,
    index=False,
    encoding="utf-8-sig"
)

# ============================================================
# STEP 14: DISPLAY FINAL RESULTS
# ============================================================

print("\n" + "=" * 60)

print("RETAILPULSE DEMAND FORECAST COMPLETED SUCCESSFULLY!")

print("=" * 60)

print("\nYOUR THREE DATASETS:")

print("\n1. Forecast_Test_Results.csv")

print("Rows:", len(test_df))

print("\n2. Future_Demand_Forecast.csv")

print("Rows:", len(future_df))

print("\n3. Forecast_Metrics.csv")

print("Rows:", len(metrics_df))

print("\nSAVED LOCATION:")

print(OUTPUT_FOLDER)

print("\nFILE CHECK:")

for file_path in [
    test_file,
    future_file,
    metrics_file
]:

    print(
        file_path.name,
        "->",
        "SAVED" if file_path.exists() else "NOT FOUND"
    )

print("\nYour original cleaned dataset is unchanged.")

print("\n" + "=" * 60)
