import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

# --- Core Reconciliation Logic (no changes here) ---
def compare_excel_files(input_file, output_file):
    results = {}
    try:
        input_wb = openpyxl.load_workbook(input_file, data_only=True)
        output_wb = openpyxl.load_workbook(output_file, data_only=True)
    except Exception as e:
        st.error(f"Fatal Error: Could not open or read the Excel files. Please check if they are valid. Details: {e}")
        return {}

    common_sheets = sorted(list(set(input_wb.sheetnames).intersection(set(output_wb.sheetnames))))

    for sheet_name in common_sheets:
        try:
            results[sheet_name] = {"correct_cells": [], "discrepancies": []}
            input_ws = input_wb[sheet_name]
            output_ws = output_wb[sheet_name]
            
            num_data_cols = input_ws.max_column
            num_data_rows = input_ws.max_row
            
            headers = {c: input_ws.cell(row=3, column=c).value for c in range(1, num_data_cols + 1)}

            # HEADER VALUE CHECK (Row 3)
            for col in range(1, num_data_cols + 1):
                template_cell = input_ws.cell(row=3, column=col)
                if template_cell.value is None or str(template_cell.value).strip() == '':
                    continue
                
                output_cell = output_ws.cell(row=3, column=col)
                column_name = headers.get(col)
                error_base = {"Cell": output_cell.coordinate, "Column": column_name}
                
                if template_cell.value != output_cell.value:
                    results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Value Mismatch", "Template_Value": template_cell.value, "Output_Value": output_cell.value})
                else:
                    results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": template_cell.value, "Output_Value": output_cell.value})

            # DATA VALUE CHECK (Row 4 onwards)
            for row_idx in range(4, num_data_rows + 1):
                for col_idx in range(1, num_data_cols + 1):
                    template_cell = input_ws.cell(row=row_idx, column=col_idx)
                    
                    if template_cell.value is None or str(template_cell.value).strip() == "":
                        continue

                    output_cell = output_ws.cell(row=row_idx, column=col_idx)
                    column_name = headers.get(col_idx, f"Column_{col_idx}")
                    error_base = {"Cell": output_cell.coordinate, "Column": column_name}
                    
                    if template_cell.value != output_cell.value:
                        results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Value Mismatch", "Template_Value": template_cell.value, "Output_Value": output_cell.value})
                    else:
                        results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": template_cell.value, "Output_Value": output_cell.value})
        
        except Exception as e:
            st.warning(f"Could not process sheet '{sheet_name}'. The following error occurred: {e}")
            if sheet_name in results:
                del results[sheet_name]
            continue
    
    return results

# --- Report Generation (no changes here) ---
def generate_excel_report(results):
    output_stream = io.BytesIO()
    writer = pd.ExcelWriter(output_stream, engine='xlsxwriter')
    workbook = writer.book

    all_results_list = []
    
    for sheet_name, sheet_results in results.items():
        if sheet_results:
            for error in sheet_results.get("discrepancies", []):
                all_results_list.append({
                    "SHEET": sheet_name, 
                    "CELL": error.get("Cell", "N/A"),
                    "FIELD": error.get("Column", "N/A"), 
                    "EXPECTED VALUE": str(error.get("Template_Value", "")),
                    "TEST VALUE": str(error.get("Output_Value", "")), 
                    "RIGHT/WRONG": "WRONG",
                    "Reason": error.get("Reason", "Unknown Error")
                })
            for correct in sheet_results.get("correct_cells", []):
                 all_results_list.append({
                    "SHEET": sheet_name, 
                    "CELL": correct.get("Cell", "N/A"),
                    "FIELD": correct.get("Column", "N/A"), 
                    "EXPECTED VALUE": str(correct.get("Template_Value", "")),
                    "TEST VALUE": str(correct.get("Output_Value", "")), 
                    "RIGHT/WRONG": "RIGHT",
                    "Reason": "OK"
                })

    columns = ["SHEET", "CELL", "FIELD", "EXPECTED VALUE", "TEST VALUE", "RIGHT/WRONG", "Reason"]
    if not all_results_list:
        detailed_df = pd.DataFrame(columns=columns)
    else:
        detailed_df = pd.DataFrame(all_results_list, columns=columns)

    correct_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "RIGHT"])
    wrong_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "WRONG"])
    total_count = correct_count + wrong_count
    accuracy_score = (correct_count / total_count * 100) if total_count > 0 else 100

    dashboard_sheet = workbook.add_worksheet("QA Dashboard")
    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter'})
    kpi_format = workbook.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter'})
    kpi_label_format = workbook.add_format({'font_size': 12, 'align': 'center', 'font_color': '#595959'})

    dashboard_sheet.merge_range('B2:F3', 'Data Reconciliation Dashboard', title_format)
    dashboard_sheet.merge_range('B5:C7', correct_count, kpi_format); dashboard_sheet.merge_range('B8:C8', 'Matching Cells', kpi_label_format)
    dashboard_sheet.merge_range('E5:F7', wrong_count, kpi_format); dashboard_sheet.merge_range('E8:F8', 'Mismatched Cells', kpi_label_format)
    dashboard_sheet.merge_range('B10:F12', f"{accuracy_score:.1f}%", kpi_format); dashboard_sheet.merge_range('B13:F13', 'Overall Accuracy Score', kpi_label_format)
    dashboard_sheet.set_column('B:F', 20)
    
    detailed_df.to_excel(writer, sheet_name="Detailed Test Results", index=False, startrow=1, header=False)
    worksheet = writer.sheets["Detailed Test Results"]

    header_format_yellow = workbook.add_format({'bold': True, 'font_color': '#000000', 'bg_color': '#FFFF00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    header_format_red = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#FF0000', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    header_format_green = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#70AD47', 'align': 'center', 'valign': 'vcenter', 'border': 1})

    for col_num, col_name in enumerate(detailed_df.columns):
        if col_num < 3:
            worksheet.write(0, col_num, col_name, header_format_yellow)
        elif col_num < 5:
            worksheet.write(0, col_num, col_name, header_format_red)
        else:
            worksheet.write(0, col_num, col_name, header_format_green)
    
    worksheet.set_column('A:G', 22)

    writer.close()
    output_stream.seek(0)
    return output_stream

# --- UI (no changes here) ---
st.set_page_config(page_title="Excel Reconciliation Tool", layout="wide")
st.title("Excel Data Reconciliation Tool")

if 'ran_comparison' not in st.session_state:
    st.session_state.ran_comparison = False
if 'results' not in st.session_state:
    st.session_state.results = {}

input_file = st.file_uploader("Upload Input (Source of Truth) Excel", type=['xlsx'])
output_file = st.file_uploader("Upload Output (File to Test) Excel", type=['xlsx'])

if input_file and output_file:
    if st.button("Run Reconciliation", type="primary"):
        with st.spinner("Performing cell-by-cell reconciliation..."):
            st.session_state.results = compare_excel_files(input_file, output_file)
        st.session_state.ran_comparison = True

if st.session_state.ran_comparison:
    results = st.session_state.get('results', {})
    if results:
        correct_count = sum(len(sheet.get('correct_cells', [])) for sheet in results.values() if sheet)
        wrong_count = sum(len(sheet.get('discrepancies', [])) for sheet in results.values() if sheet)
        total_count = correct_count + wrong_count
        
        accuracy_score = (correct_count / total_count * 100) if total_count > 0 else 100

        st.header("On-Screen Reconciliation Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric(label="📊 Accuracy Score", value=f"{accuracy_score:.2f}%")
        col2.metric(label="✅ Matching Cells", value=correct_count)
        col3.metric(label="❌ Mismatched Cells", value=wrong_count, delta_color="inverse")
                
        st.markdown("---")
        # --- THIS IS THE ONLY CHANGE ---
        st.download_button(
            label="📄 Download Full Test Report (Excel)", 
            data=generate_excel_report(st.session_state.results), 
            file_name=f"Test_Report_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
