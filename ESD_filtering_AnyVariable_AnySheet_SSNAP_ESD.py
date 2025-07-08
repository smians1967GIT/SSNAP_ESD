import pandas as pd
import numpy as np
import os
import gradio as gr
import tempfile

def get_sheet_names(file_obj):
    if file_obj is None:
        return [], None
    try:
        xls = pd.ExcelFile(file_obj.name)
        return xls.sheet_names, file_obj
    except Exception:
        return [], None

def extract_metrics(file_obj, sheet_name, metric_ids, quarter):
    if file_obj is None or sheet_name is None:
        return "❌ Please upload a file and select a sheet", None

    try:
        xls = pd.ExcelFile(file_obj.name)
        df = xls.parse(sheet_name, header=None)
        metadata = df.iloc[0:4, 4:].T
        metadata.columns = ['Team Type', 'Region', 'Trust', 'Team']
        metadata = metadata.dropna(subset=['Team']).reset_index(drop=True)
    except Exception as e:
        return f"❌ Failed to load or parse sheet: {e}", None

    metric_list = [m.strip() for m in metric_ids.split(",")]
    all_records = []

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

# Gradio App
with gr.Blocks() as demo:
    gr.Markdown("### 🧠 SSNAP Community Metric Extractor")

    file_input = gr.File(label="Upload Excel File (.xlsx)", file_types=[".xlsx"])
    sheet_state = gr.State()
    sheet_dropdown = gr.Dropdown(label="Select Sheet", choices=[], visible=True)
    metric_input = gr.Textbox(label="Enter Metric ID(s) (e.g., L27.3 or L27.3, L32.3)")
    quarter_input = gr.Textbox(label="Enter Quarter (e.g., 2025-Q1)")

    status_output = gr.Textbox(label="Status")
    download_output = gr.File(label="Download Cleaned CSV")
    submit_button = gr.Button("Extract and Export")

    def update_dropdown(file_obj):
        sheets, file_ref = get_sheet_names(file_obj)
        return gr.Dropdown(choices=sheets, value=sheets[0] if sheets else None), file_ref

    file_input.change(
        fn=update_dropdown,
        inputs=file_input,
        outputs=[sheet_dropdown, sheet_state]
    )

    submit_button.click(
        fn=extract_metrics,
        inputs=[sheet_state, sheet_dropdown, metric_input, quarter_input],
        outputs=[status_output, download_output]
    )

if __name__ == "__main__":
    demo.launch()
