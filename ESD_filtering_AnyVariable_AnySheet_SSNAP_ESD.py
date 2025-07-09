import pandas as pd
import gradio as gr
import os
import tempfile

def extract_esd_metrics(file_obj):
    if file_obj is None:
        return "❌ No file uploaded", None, None

    try:
        xls = pd.ExcelFile(file_obj.name)
        target_sheets = [s for s in xls.sheet_names if s.startswith("L")]
    except Exception as e:
        return f"❌ Failed to read Excel: {e}", None, None

    combined_records = []

    for sheet in target_sheets:
        try:
            df = xls.parse(sheet, header=None)

            # Extract team metadata from rows 0–3, columns F+ (index 5+)
            team_data = []
            for col in range(5, df.shape[1]):
                team_type = df.iloc[0, col]
                region = df.iloc[1, col]
                trust = df.iloc[2, col]
                esd_team = df.iloc[3, col]

                if pd.notna(esd_team):
                    team_data.append({
                        "col": col,
                        "Team Type": str(team_type).strip() if pd.notna(team_type) else None,
                        "Region": str(region).strip() if pd.notna(region) else None,
                        "Trust": str(trust).strip() if pd.notna(trust) else None,
                        "ESD Team": str(esd_team).strip(),
                    })

            # Loop through data rows where Data Type = % or median
            for i in range(4, df.shape[0]):
                data_type = str(df.iloc[i, 3]).lower()
                if "%" in data_type or "median" in data_type:
                    metric_label = df.iloc[i, 0]
                    metric_id = df.iloc[i, 1]

                    for team in team_data:
                        value = df.iloc[i, team["col"]]
                        if pd.notna(value) and str(value).strip() not in ["Too few to report", "Reported annually",".", "N/A", ""]:
                            combined_records.append({
                                **team,
                                "Sheet": sheet,
                                "Metric ID": metric_id,
                                "Metric Label": metric_label,
                                "Data Type": df.iloc[i, 3],
                                "Value": value
                            })
        except Exception:
            continue

    if not combined_records:
        return "❌ No % or median metrics found.", None, None

    df_combined = pd.DataFrame(combined_records)
    out_path = os.path.join(tempfile.gettempdir(), "SSNAP_ESD_Metrics_Output.xlsx")
    df_combined.to_excel(out_path, index=False)

    return "✅ ESD metrics with % or median extracted.", df_combined, out_path

# Gradio app interface
with gr.Blocks() as demo:
    gr.Markdown("### 🧠 SSNAP ESD Extractor – % and Median Metrics")

    file_input = gr.File(label="Upload SSNAP Excel File (.xlsx)", file_types=[".xlsx"])
    status_output = gr.Textbox(label="Status")
    table_output = gr.Dataframe(label="Extracted Metric Records")
    download_button = gr.File(label="Download Excel Output")

    file_input.change(
        fn=extract_esd_metrics,
        inputs=[file_input],
        outputs=[status_output, table_output, download_button]
    )

if __name__ == "__main__":
    demo.launch()
