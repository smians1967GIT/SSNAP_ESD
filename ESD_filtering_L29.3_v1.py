import pandas as pd
import numpy as np
import os

# --- Configuration ---
file_path = r"C:\SSNAP_dashboard\SSNAP_Dashboard_New_Metrics_2025\Datasets\JanMar2025-FullResultsPortfolioESD.xls"
sheet_name = "L4. Outcome measures"
metric_id = "L29.3"
output_csv = r"C:\SSNAP_dashboard\SSNAP_Dashboard_New_Metrics_2025\Datasets\L29_3_Outcome_Measures_2025Q1_FIXED.csv"

# --- Step 1: File and Sheet Validation ---
if not os.path.exists(file_path):
    raise FileNotFoundError(f"❌ Excel file not found at: {file_path}")

try:
    xls = pd.ExcelFile(file_path)
except Exception as e:
    raise IOError(f"❌ Failed to load Excel file: {e}")

if sheet_name not in xls.sheet_names:
    raise ValueError(f"❌ Sheet '{sheet_name}' not found in workbook. Available sheets: {xls.sheet_names}")

# --- Step 2: Load Sheet and Extract Metadata ---
try:
    df = xls.parse(sheet_name, header=None)
except Exception as e:
    raise RuntimeError(f"❌ Could not parse sheet '{sheet_name}': {e}")

try:
    metadata = df.iloc[0:4, 4:].T
    metadata.columns = ['Team Type', 'Region', 'Trust', 'Team']
    metadata = metadata.dropna(subset=['Team']).reset_index(drop=True)
except Exception as e:
    raise RuntimeError(f"❌ Failed to extract metadata: {e}")

# --- Step 3: Extract Metric Row and Values ---
try:
    metric_row = df[df.iloc[:, 1] == metric_id]
    if metric_row.empty:
        raise ValueError(f"❌ Metric ID '{metric_id}' not found in sheet '{sheet_name}'")
    
    metric_label = metric_row.iloc[0, 0]
    metric_values = metric_row.iloc[0, 4:4 + len(metadata)]
except Exception as e:
    raise RuntimeError(f"❌ Failed to extract metric '{metric_id}': {e}")

# --- Step 4: Construct Records ---
records = []
for i in range(min(len(metadata), len(metric_values))):
    try:
        team = metadata.iloc[i]
        value = metric_values.iloc[i]

        clean_value = (
            np.nan if str(value).strip() in ["", " ", "Too few to report", ".", "N/A", "nan"] else value
        )

        records.append({
            "Quarter": "2025-Q1",
            "Domain": sheet_name,
            "Team Type": team["Team Type"],
            "Region": team["Region"],
            "Trust": team["Trust"],
            "Team": team["Team"],
            "Metric ID": metric_id,
            "Metric Label": metric_label,
            "Value": clean_value
        })
    except Exception as e:
        print(f"⚠️ Skipped row {i} due to error: {e}")
        continue

# --- Step 5: Export to CSV ---
try:
    df_cleaned = pd.DataFrame(records)
    df_cleaned.to_csv(output_csv, index=False)
    print(f"✅ Exported cleaned data to: {output_csv}")
except Exception as e:
    raise IOError(f"❌ Failed to export CSV: {e}")
