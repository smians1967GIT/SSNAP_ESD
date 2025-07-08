import pandas as pd
import gradio as gr
import tempfile
import os

def extract_team_and_metrics(file_obj, region="London"):
    if file_obj is None:
        return "❌ No file uploaded.", None, None

    try:
        xls = pd.ExcelFile(file_obj.name)
        sheet_names = xls.sheet_names
    except Exception as e:
        return f"❌ Failed to read Excel file: {e}", None, None

    combined_records = []

    for sheet in sheet_names:
        try:
            df = xls.parse(sheet, header=None)

            # Get metadata rows
            trust_row = df.iloc[2, 1:]  # Trusts
            team_row = df.iloc[3, 1:]   # Team names

            # Build team lookup
            team_info = []
            for col in range(len(trust_row)):
                trust = trust_row.iloc[col]
                team = team_row.iloc[col]
                if pd.notna(trust) and pd.notna(team):
                    team_info.append({
                        "col": col + 1,  # +1 because first col is metric names
                        "Team Type": "ESD team",
                        "Region": region,
                        "Trust": str(trust).strip(),
                        "ESD Team": str(team).strip()
                    })

            # Loop through all metric rows where column D contains '%' or 'median'
            for i in range(len(df)):
                data_type = str(df.iloc[i, 3]).lower()
                if "%" in data_type or "median" in data_type:
                    metric_label = df.iloc[i, 0]
                    metric_id = df.iloc[i, 1]
                    for team in team_info:
                        value = df.iloc[i, team["col"]]
                        if pd.notna(value) and str(value).strip() not in ["Too few to report", ".", "N/A"]:
                            combined_records.append({
                                **team,
                                "Metric ID": metric_id,
                                "Metric Label": metric_label,
                                "Data Type": df.iloc[i, 3],
                                "Value": value
                            })

        except Exception:
            continue

    if not combined_records:
        return "❌ No matching metrics found with '%' or 'median' types", None, None

    df_combined = pd.DataFrame(combined_records)
    output_path = os.path.join(tempfile.gettempdir(), "SSNAP_ESD_Median_Percentage_Metrics.xlsx")
    df_combined.to_excel(output_path, index=False)
    return "✅ Extracted ESD teams with matching metrics", df_combined, output_path

# --- Gradio Interface ---
with gr.Blocks() as demo:
    gr.Markdown("### 🧠 SSNAP ESD Extractor: Teams + % and Median Metrics")

    file_input = gr.File(label="Upload SSNAP Excel File (.xlsx)", file_types=[".xlsx"])
    region_input = gr.Textbox(label="Enter Region (default: London)", value="London")
    status_output = gr.Textbox(label="Status")
    table_output = gr.Dataframe(label="Extracted Metrics")
    download_output = gr.File(label="Download Excel")

    file_input.change(
        fn=extract_team_and_metrics,
        inputs=[file_input, region_input],
        outputs=[status_output, table_output, download_output]
    )

if __name__ == "__main__":
    demo.launch()
