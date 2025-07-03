import pandas as pd
import numpy as np
import os
import gradio as gr

# --- Configuration ---
FILE_PATH = r"C:\SSNAP_dashboard\SSNAP_Dashboard_New_Metrics_2025\Datasets\\JanMar2025-FullResultsPortfolioESD.xls"
SHEET_NAME = "L4. Outcome measures"
QUARTER = "2025-Q1"
EXPORT_DIR = r"C:\SSNAP_dashboard\SSNAP_Dashboard_New_Metrics_2025\Datasets"

def extract_metric(metric_id):
    if not os.path.exists(FILE_PATH):
        return f"❌ Excel file not found at: {FILE_PATH}", None

    try:
        xls = pd.ExcelFile(FILE_PATH)
    except Exception as e:
        return f"❌ Failed to load Excel file: {e}", None

    if SHEET_NAME not in xls.sheet_names:
        return f"❌ Sheet '{SHEET_NAME}' not found in workbook. Available sheets: {xls.sheet_names}", None

    try:
        df = xls.parse(SHEET_NAME, header=None)
    except Exception as e:
        return f"❌ Could not parse sheet '{SHEET_NAME}': {e}", None

    try:
        metadata = df.iloc[0:4, 4:].T
        metadata.columns = ['Team Type', 'Region', 'Trust', 'Team']
        metadata = metadata.dropna(subset=['Team']).reset_index(drop=True)
    except Exception as e:
        return f"❌ Failed to extract metadata: {e}", None

    try:
        metric_row = df[df.iloc[:, 1] == metric_id]
        if metric_row.empty:
            return f"❌ Metric ID '{metric_id}' not found in sheet", None

        metric_label = metric_row.iloc[0, 0]
        metric_values = metric_row.iloc[0, 4:4 + len(metadata)]
    except Exception as e:
        return f"❌ Failed to extract metric '{metric_id}': {e}", None

    records = []
    for i in range(min(len(metadata), len(metric_values))):
        try:
            team = metadata.iloc[i]
            value = metric_values.iloc[i]
            clean_value = (
                np.nan if str(value).strip() in ["", " ", "Too few to report", ".", "N/A", "nan"] else value
            )
            records.append({
                "Quarter": QUARTER,
                "Domain": SHEET_NAME,
                "Team Type": team["Team Type"],
                "Region": team["Region"],
                "Trust": team["Trust"],
                "Team": team["Team"],
                "Metric ID": metric_id,
                "Metric Label": metric_label,
                "Value": clean_value
            })
        except Exception as e:
            continue

    try:
        df_cleaned = pd.DataFrame(records)
        output_csv = os.path.join(EXPORT_DIR, f"{metric_id}_Outcome_Measures_{QUARTER}_FIXED.csv")
        df_cleaned.to_csv(output_csv, index=False)
        return f"✅ Exported cleaned data to: {output_csv}", output_csv
    except Exception as e:
        return f"❌ Failed to export CSV: {e}", None

# --- Gradio Interface ---
def gradio_interface(metric_id):
    message, file_path = extract_metric(metric_id)
    if file_path and os.path.exists(file_path):
        return message, file_path
    return message, None

demo = gr.Interface(
    fn=gradio_interface,
    inputs=gr.Textbox(label="Enter Metric ID (e.g., L32.3)"),
    outputs=[
        gr.Textbox(label="Status"),
        gr.File(label="Download Cleaned CSV")
    ],
    title="SSNAP Metric Extractor",
    description="Enter a metric ID from the Outcome Measures sheet to extract and clean data for export."
)

if __name__ == "__main__":
    demo.launch(allowed_paths=[r"C:\SSNAP_dashboard\SSNAP_Dashboard_New_Metrics_2025\Datasets"])

