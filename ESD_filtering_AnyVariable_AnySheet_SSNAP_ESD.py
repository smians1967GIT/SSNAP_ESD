import pandas as pd
import numpy as np
import os
import gradio as gr
import tempfile

def extract_metrics(file_obj, sheet_name, metric_ids, quarter):
    if file_obj is None:
        return "❌ No file uploaded", None

    try:
        xls = pd.ExcelFile(file_obj.name)
    except Exception as e:
        return f"❌ Failed to load Excel file: {e}", None

    if sheet_name not in xls.sheet_names:
        return f"❌ Sheet '{sheet_name}' not found. Available sheets: {xls.sheet_names}", None

    try:
        df = xls.parse(sheet_name, header=None)
        metadata = df.iloc[0:4, 4:].T
        metadata.columns = ['Team Type', 'Region', 'Trust', 'Team']
        metadata = metadata.dropna(subset=['Team']).reset_index(drop=True)
    except Exception as e:
        return f"❌ Failed to parse or extract metadata: {e}", None

    all_records = []
    metric_list = [m.strip() for m in metric_ids.split(",")]

    for metric_id in metric_list:
        try:
            metric_row = df[df.iloc[:, 1] == metric_id]
            if metric_row.empty:
                continue
            metric_label = metric_row.iloc[0, 0]
            metric_values = metric_row.iloc[0, 4:4 + len(metadata)]

            for i in range(len(metadata)):
                value = metric_values.iloc[i]
                team = metadata.iloc[i]
                clean_value = (
                    np.nan if str(value).strip() in ["", " ", ".", "N/A", "Too few to report", "nan"] else value
                )
                all_records.append({
                    "Quarter": quarter,
                    "Domain": sheet_name,
                    "Team Type": team["Team Type"],
                    "Region": team["Region"],
                    "Trust": team["Trust"],
                    "Team": team["Team"],
                    "Metric ID": metric_id,
                    "Metric Label": metric_label,
                    "Value": clean_value
                })
        except Exception:
            continue

    if not all_records:
        return f"❌ No matching metrics found in sheet '{sheet_name}' for: {metric_ids}", None

    try:
        df_cleaned = pd.DataFrame(all_records)
        filename = f"{sheet_name.replace(' ', '_')}_Metrics_{quarter}.csv"
        output_path = os.path.join(tempfile.gettempdir(), filename)
        df_cleaned.to_csv(output_path, index=False)
        return f"✅ Exported cleaned data to: {filename}", output_path
    except Exception as e:
        return f"❌ Failed to export cleaned CSV: {e}", None

# --- Gradio App Interface ---
demo = gr.Interface(
    fn=extract_metrics,
    inputs=[
        gr.File(label="Upload Excel File (.xlsx)", file_types=[".xlsx"]),
        gr.Textbox(label="Enter Sheet Name (e.g., L4. Outcome measures)"),
        gr.Textbox(label="Enter Metric ID(s) (e.g., L27.3 or L27.3, L32.3)"),
        gr.Textbox(label="Enter Quarter (e.g., 2025-Q1)")
    ],
    outputs=[
        gr.Textbox(label="Status"),
        gr.File(label="Download Cleaned CSV")
    ],
    title="SSNAP Community Metric Extractor",
    description="Upload a SSNAP .xlsx report and extract cleaned metrics from any sheet. Supports multiple metric IDs and custom quarters."
)

if __name__ == "__main__":
    demo.launch()
