# import streamlit as st
# import pandas as pd
# import openpyxl
# import io
# from datetime import datetime

# # --- Helper functions (no changes here) ---
# def get_number_format_name(format_code):
#     if format_code is None: return "General"
#     format_code = str(format_code).lower()
#     if 'yyyy' in format_code or 'd' in format_code or 'm' in format_code: return f"Date ({format_code.upper()})"
#     if '$' in format_code or 'accounting' in format_code: return "Currency"
#     if '%' in format_code: return "Percentage"
#     if format_code == 'general': return "General"
#     if '@' in format_code: return "Text"
#     return f"Custom ({format_code})"

# def get_color_name(rgb_code):
#     if rgb_code is None or rgb_code == '00000000': return "No Fill"
#     color_map = { "FFFFFFFF": "White", "FF000000": "Black", "FFFF0000": "Red", "FF00FF00": "Green", "FF0000FF": "Blue", "FFFFFF00": "Yellow", "FF00B0F0": "Light Blue", "FF92D050": "Light Green", "FFC0C0C0": "Light Gray" }
#     return color_map.get(str(rgb_code).upper(), f"Custom Color (#{rgb_code})")

# def check_specific_datatype(value):
#     if isinstance(value, (int, float)): return "Numeric"
#     if isinstance(value, str):
#         if str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']: return "Binary (Yes/No)"
#         return "Text"
#     return "Other"

# # --- Core comparison logic (no changes here) ---
# def compare_excel_files(input_file, output_file):
#     results = { "overall": {}, "sheets_summary": {} }
#     try:
#         input_wb, output_wb = openpyxl.load_workbook(input_file), openpyxl.load_workbook(output_file)
#         common_sheets = sorted(list(set(input_wb.sheetnames).intersection(set(output_wb.sheetnames))))

#         for sheet_name in common_sheets:
#             results[sheet_name] = {"successes": [], "discrepancies": [], "correct_cells": []}
#             input_ws, output_ws = input_wb[sheet_name], output_wb[sheet_name]
#             input_df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
#             output_df = pd.read_excel(output_file, sheet_name=sheet_name, engine='openpyxl')

#             if list(input_df.columns) != list(output_df.columns):
#                 results[sheet_name]["discrepancies"].append({"Category": "Formatting", "Cell": "Row 1", "Column": "N/A", "Reason": "Column Header Mismatch", "Template_Value": f"{list(input_df.columns)}", "Output_Value": f"{list(output_df.columns)}"})
#             else:
#                 results[sheet_name]["successes"].append("Column Headers are identical.")

#             wrong_cell_coords, max_rows, max_cols = set(), min(input_ws.max_row, output_ws.max_row), min(input_ws.max_column, output_ws.max_column)
#             total_cells_checked = max_rows * max_cols

#             for row in range(1, max_rows + 1):
#                 for col in range(1, max_cols + 1):
#                     input_cell, output_cell = input_ws.cell(row=row, column=col), output_ws.cell(row=row, column=col)
#                     error_base = {"Cell": input_cell.coordinate, "Column": input_df.columns[col-1] if col-1 < len(input_df.columns) else 'N/A'}
#                     cell_has_error = False

#                     def add_error(reason, cat, t_val, o_val):
#                         nonlocal cell_has_error; cell_has_error = True
#                         wrong_cell_coords.add(input_cell.coordinate)
#                         results[sheet_name]["discrepancies"].append({**error_base, "Category": cat, "Reason": reason, "Template_Value": t_val, "Output_Value": o_val})

#                     is_overflow = isinstance(output_cell.value, str) and len(output_cell.value) > (output_ws.column_dimensions.get(openpyxl.utils.get_column_letter(col), openpyxl.worksheet.dimensions.ColumnDimension(output_ws)).width or 8.43)
#                     if not is_overflow and input_cell.value != output_cell.value: add_error("Value Mismatch", "Accuracy", input_cell.value, output_cell.value)
#                     if check_specific_datatype(input_cell.value) != check_specific_datatype(output_cell.value): add_error("Type Mismatch", "Datatype", check_specific_datatype(input_cell.value), check_specific_datatype(output_cell.value))
#                     if (input_cell.font.name, input_cell.font.size, input_cell.font.bold) != (output_cell.font.name, output_cell.font.size, output_cell.font.bold): add_error("Font Mismatch", "Formatting", f"{input_cell.font.name}, {input_cell.font.size}pt, Bold:{input_cell.font.bold}", f"{output_cell.font.name}, {output_cell.font.size}pt, Bold:{output_cell.font.bold}")
#                     if input_cell.fill.fgColor.rgb != output_cell.fill.fgColor.rgb: add_error("Background Color", "Formatting", get_color_name(input_cell.fill.fgColor.rgb), get_color_name(output_cell.fill.fgColor.rgb))
#                     if input_cell.number_format != output_cell.number_format: add_error("Number Format", "Formatting", get_number_format_name(input_cell.number_format), get_number_format_name(output_cell.number_format))
#                     if input_cell.alignment.wrap_text != output_cell.alignment.wrap_text: add_error("Text Wrap Setting", "Formatting", f"Wrap: {input_cell.alignment.wrap_text}", f"Wrap: {output_cell.alignment.wrap_text}")
#                     if is_overflow and not output_cell.alignment.wrap_text: col_width = output_ws.column_dimensions.get(openpyxl.utils.get_column_letter(col), openpyxl.worksheet.dimensions.ColumnDimension(output_ws)).width or 8.43; add_error("Text Overflow", "Formatting", f"Fits width of ~{col_width:.0f}", f"Text length is {len(output_cell.value)}")

#                     if not cell_has_error:
#                         results[sheet_name]["correct_cells"].append({"Cell": input_cell.coordinate, "Column": error_base["Column"], "Template_Value": input_cell.value, "Output_Value": output_cell.value})

#             all_cols = set(input_ws.column_dimensions.keys()) | set(output_ws.column_dimensions.keys())
#             for col_letter in all_cols:
#                 input_width = (input_ws.column_dimensions.get(col_letter) or openpyxl.worksheet.dimensions.ColumnDimension(input_ws)).width or 8.43
#                 output_width = (output_ws.column_dimensions.get(col_letter) or openpyxl.worksheet.dimensions.ColumnDimension(output_ws)).width or 8.43
#                 if abs(input_width - output_width) > 0.1:
#                     results[sheet_name]["discrepancies"].append({"Category": "Formatting", "Cell": f"Column {col_letter}", "Column": col_letter, "Reason": "Column Width Mismatch", "Template_Value": f"{input_width:.2f}", "Output_Value": f"{output_width:.2f}"}); wrong_cell_coords.add(f"Col {col_letter}")
#             all_rows = set(input_ws.row_dimensions.keys()) | set(output_ws.row_dimensions.keys())
#             for row_idx in all_rows:
#                 input_height = (input_ws.row_dimensions.get(row_idx) or openpyxl.worksheet.dimensions.RowDimension(input_ws)).height or 15.0
#                 output_height = (output_ws.row_dimensions.get(row_idx) or openpyxl.worksheet.dimensions.RowDimension(output_ws)).height or 15.0
#                 if abs(input_height - output_height) > 0.1:
#                     results[sheet_name]["discrepancies"].append({"Category": "Formatting", "Cell": f"Row {row_idx}", "Column": "N/A", "Reason": "Row Height Mismatch", "Template_Value": f"{input_height:.2f}", "Output_Value": f"{output_height:.2f}"}); wrong_cell_coords.add(f"Row {row_idx}")

#             num_wrong_cells = len(wrong_cell_coords)
#             results[sheet_name]["successes"].append(f"Cell-Level Checks: {total_cells_checked - num_wrong_cells} out of {total_cells_checked} cells were correct.")
#             discrepancies = results[sheet_name]["discrepancies"]
#             results["sheets_summary"][sheet_name] = {"status": "Fail" if discrepancies else "Pass", "total_cells": total_cells_checked, "wrong_cells": num_wrong_cells, "accuracy_errors": len([d for d in discrepancies if d['Category'] == 'Accuracy']), "datatype_errors": len([d for d in discrepancies if d['Category'] == 'Datatype']), "formatting_errors": len([d for d in discrepancies if d['Category'] == 'Formatting'])}
#     except Exception as e: st.error(f"A critical error occurred during comparison: {e}")
#     return results

# # --- UPDATED: The Excel report function now includes Key Action Points ---
# def generate_excel_report(results):
#     output_stream = io.BytesIO()
#     writer = pd.ExcelWriter(output_stream, engine='xlsxwriter')
#     workbook = writer.book

#     # Calculate global metrics
#     sheets_summary = results.get("sheets_summary", {}); total_sheets = len(sheets_summary)
#     passed_sheets = len([s for s, data in sheets_summary.items() if data.get("status") == "Pass"]); failed_sheets = total_sheets - passed_sheets
#     total_cells = sum(d.get('total_cells', 0) for d in sheets_summary.values()); total_wrong = sum(d.get('wrong_cells', 0) for d in sheets_summary.values() if isinstance(d.get('wrong_cells'), int)); total_correct = total_cells - total_wrong
#     cell_pass_percentage = (total_correct / total_cells * 100) if total_cells > 0 else 100
#     total_accuracy_errors = sum(d.get('accuracy_errors', 0) for d in sheets_summary.values()); total_datatype_errors = sum(d.get('datatype_errors', 0) for d in sheets_summary.values()); total_formatting_errors = sum(d.get('formatting_errors', 0) for d in sheets_summary.values())

#     # Part 1: High-Impact Strategy Dashboard
#     strategy_sheet = workbook.add_worksheet("QA Dashboard")
#     title_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter', 'font_color': '#2F5496'})
#     kpi_format = workbook.add_format({'bold': True, 'font_size': 24, 'align': 'center', 'valign': 'vcenter'})
#     kpi_label_format = workbook.add_format({'font_size': 11, 'align': 'center', 'font_color': '#595959'})
#     header_format = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#DDEBF7', 'border': 1})
#     section_header_format = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#2F5496'})

#     strategy_sheet.merge_range('B2:I3', 'Excel Quality Assurance Dashboard', title_format)
#     strategy_sheet.write('B5', 'Cell-Level Metrics', section_header_format)
#     strategy_sheet.merge_range('B6:C8', f"{cell_pass_percentage:.1f}%", kpi_format); strategy_sheet.merge_range('B9:C9', 'Overall Cell Match %', kpi_label_format)
#     strategy_sheet.merge_range('E6:F8', total_correct, kpi_format); strategy_sheet.merge_range('E9:F9', 'Correct Cells', kpi_label_format)
#     strategy_sheet.merge_range('H6:I8', total_wrong, kpi_format); strategy_sheet.merge_range('H9:I9', 'Wrong Cells', kpi_label_format)
    
#     strategy_sheet.write('B11', 'Sheet-Level Metrics', section_header_format)
#     strategy_sheet.merge_range('B12:C14', total_sheets, kpi_format); strategy_sheet.merge_range('B15:C15', 'Total Sheets', kpi_label_format)
#     strategy_sheet.merge_range('E12:F14', passed_sheets, kpi_format); strategy_sheet.merge_range('E15:F15', 'Sheets Passed', kpi_label_format)
#     strategy_sheet.merge_range('H12:I14', failed_sheets, kpi_format); strategy_sheet.merge_range('H15:I15', 'Sheets Failed', kpi_label_format)

#     strategy_sheet.merge_range('B18:D18', 'Key Insights', header_format)
#     sheet_with_most_errors, max_errors = "", 0
#     for sheet, data in sheets_summary.items():
#         if data.get('wrong_cells', 0) > max_errors: max_errors, sheet_with_most_errors = data.get('wrong_cells', 0), sheet
#     most_common_error = max({"Accuracy": total_accuracy_errors, "Datatype": total_datatype_errors, "Formatting": total_formatting_errors}, key=lambda k: {"Accuracy": total_accuracy_errors, "Datatype": total_datatype_errors, "Formatting": total_formatting_errors}[k]) if total_wrong > 0 else "None"
    
#     strategy_sheet.write('B19', f"1. Highest Risk Area: The sheet '{sheet_with_most_errors}' has the most errors ({max_errors}). Review first.")
#     strategy_sheet.write('B20', f"2. Systemic Problem: The most common issue category is '{most_common_error}'.")
    
#     error_df = pd.DataFrame([{"Category": "Accuracy Errors", "Count": total_accuracy_errors}, {"Category": "Datatype Errors", "Count": total_datatype_errors}, {"Category": "Formatting Errors", "Count": total_formatting_errors}])
#     error_df.to_excel(writer, sheet_name="QA Dashboard", startrow=21, startcol=1, index=False)
    
#     chart = workbook.add_chart({'type': 'bar'}); chart_data = [total_accuracy_errors, total_datatype_errors, total_formatting_errors]
#     strategy_sheet.write_column('M1', ["Accuracy", "Datatype", "Formatting"]); strategy_sheet.write_column('N1', chart_data)
#     chart.add_series({'name': 'Error Distribution', 'categories': '=\'QA Dashboard\'!$M$1:$M$3', 'values': '=\'QA Dashboard\'!$N$1:$N$3'})
#     chart.set_title({'name': 'Error Breakdown by Category'}); chart.set_legend({'none': True})
#     strategy_sheet.insert_chart('F18', chart, {'x_offset': 25, 'y_offset': 10})
    
#     # --- THIS IS THE NEWLY ADDED SECTION ---
#     strategy_sheet.merge_range('B26:D26', 'Key Action Points', header_format)
#     strategy_sheet.write('B27', f"1. Prioritize all 'Accuracy' errors to ensure data correctness.")
#     strategy_sheet.write('B28', f"2. Focus on sheet '{sheet_with_most_errors}' to resolve the highest number of issues quickly.")
#     strategy_sheet.write('B29', f"3. Investigate the root cause of '{most_common_error}' errors, as this is the most systemic problem.")
#     # --- END OF NEW SECTION ---

#     strategy_sheet.set_column('B:I', 20)

#     # Part 2: Summary Sheet
#     pd.DataFrame({"Metric": ["Overall Cell Match Percentage", "Total Cells Checked", "Total Correct Cells", "Total Wrong Cells", "Total Sheets Checked", "Sheets Passed", "Sheets Failed"],"Value": [f"{cell_pass_percentage:.2f}%", total_cells, total_correct, total_wrong, total_sheets, passed_sheets, failed_sheets]}).to_excel(writer, sheet_name='Summary', startrow=1, index=False, header=False)
#     summary_sheet = writer.sheets['Summary']; summary_sheet.write('A1', 'Overall Test Summary'); summary_sheet.set_column('A:A', 30); summary_sheet.set_column('B:B', 20)
#     categorized_data = [{"Sheet Name": n, "Status": d.get("status"), "Accuracy Errors": d.get("accuracy_errors", 0), "Datatype Errors": d.get("datatype_errors", 0), "Formatting Errors": d.get("formatting_errors", 0)} for n, d in sheets_summary.items()]
#     summary_df2 = pd.DataFrame(categorized_data); summary_df2.to_excel(writer, sheet_name='Summary', startrow=9, index=False)
#     summary_sheet.write('A9', 'Per-Sheet Error Breakdown'); summary_sheet.set_column('C:F', 20)

#     # Part 3: Detailed Sheets with Separated Tables
#     for sheet_name, sheet_results in results.items():
#         if sheet_name in ["overall", "sheets_summary"]: continue
#         report_sheet_name = f"{sheet_name}_Report"; worksheet = workbook.add_worksheet(report_sheet_name); current_row = 0
#         incorrect_cells = pd.DataFrame(sheet_results.get("discrepancies", []))
#         if not incorrect_cells.empty:
#             worksheet.write(current_row, 0, "Incorrect Cells", header_format); current_row += 1
#             incorrect_cells[["Cell", "Column", "Reason", "Template_Value", "Output_Value"]].to_excel(writer, sheet_name=report_sheet_name, startrow=current_row, index=False); current_row += len(incorrect_cells) + 2
#         correct_cells = pd.DataFrame(sheet_results.get("correct_cells", []))
#         if not correct_cells.empty:
#             worksheet.write(current_row, 0, "Correct Cells", header_format); current_row += 1
#             correct_cells[["Cell", "Column", "Template_Value", "Output_Value"]].to_excel(writer, sheet_name=report_sheet_name, startrow=current_row, index=False)
#         worksheet.set_column('A:F', 25)
        
#     writer.close(); output_stream.seek(0)
#     return output_stream

# # --- UI (no changes needed) ---
# st.set_page_config(page_title="Advanced Excel Comparator", layout="wide"); st.title("Advanced Excel Quality Assurance Tool")
# input_file = st.file_uploader("Upload Input (Template) Excel", type=['xlsx']); output_file = st.file_uploader("Upload Output (To Test) Excel", type=['xlsx'])
# if input_file and output_file:
#     if st.button("Run Advanced Comparison", type="primary"):
#         with st.spinner("Performing deep analysis..."): st.session_state['results'] = compare_excel_files(input_file, output_file)
#         st.session_state['ran_comparison'] = True
# if 'ran_comparison' in st.session_state:
#     results = st.session_state['results']; sheets_summary = results.get("sheets_summary", {});
#     total_sheets = len(sheets_summary); passed_sheets = len([s for s, data in sheets_summary.items() if data.get("status") == "Pass"]); failed_sheets = total_sheets - passed_sheets
#     total_cells = sum(d.get('total_cells', 0) for d in sheets_summary.values()); total_wrong = sum(d.get('wrong_cells', 0) for d in sheets_summary.values() if isinstance(d.get('wrong_cells'), int)); total_correct = total_cells - total_wrong
#     cell_pass_percentage = (total_correct / total_cells * 100) if total_cells > 0 else 100
#     st.header("On-Screen Comparison Summary")
#     st.subheader("Cell-Level Summary")
#     col1, col2, col3 = st.columns(3)
#     col1.metric(label="📊 Overall Cell Match Percentage", value=f"{cell_pass_percentage:.2f}%"); col2.metric(label="✅ Correct Cells", value=total_correct); col3.metric(label="❌ Wrong Cells", value=total_wrong, delta_color="inverse")
#     st.subheader("Sheet-Level Summary")
#     col4, col5, col6 = st.columns(3)
#     col4.metric(label="📋 Total Sheets Checked", value=total_sheets); col5.metric(label="✅ Sheets Passed", value=passed_sheets); col6.metric(label="❌ Sheets Failed", value=failed_sheets, delta_color="inverse")
#     for sheet_name, sheet_results in results.items():
#         if sheet_name in ["overall", "sheets_summary"]: continue
#         with st.expander(f"Results for Sheet: '{sheet_name}'", expanded=bool(sheet_results.get('discrepancies'))):
#             st.subheader("✅ Passed Checks")
#             if sheet_results.get('successes'):
#                 for success in sheet_results['successes']: st.markdown(f"- {success}")
#             else: st.write("No specific checks passed for this sheet.")
#             st.subheader("❌ Discrepancies Found")
#             discrepancy_list = sheet_results.get('discrepancies', [])
#             if discrepancy_list:
#                 categories = {"Accuracy": [], "Datatype": [], "Formatting": []}
#                 for error in discrepancy_list: categories.get(error.get("Category", "Other"), []).append(error)
#                 for category, errors in categories.items():
#                     if errors:
#                         st.markdown(f"**{category} Errors ({len(errors)} found):**")
#                         for error in errors: st.markdown(f"  - **Cell:** `{error['Cell']}` | **Reason:** *{error['Reason']}* | **Template:** `{error['Template_Value']}` | **Output:** `{error['Output_Value']}`")
#             else: st.write("No discrepancies found in this sheet.")
#     st.markdown("---"); st.download_button(label="📄 Download Full Categorized Report (Excel)", data=generate_excel_report(results), file_name=f"Advanced_Test_Report_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# import streamlit as st
# import pandas as pd
# import openpyxl
# import io
# from datetime import datetime

# # --- Helper functions (no changes here) ---
# def get_number_format_name(format_code):
#     if format_code is None: return "General"
#     format_code = str(format_code).lower()
#     if 'yyyy' in format_code or 'd' in format_code or 'm' in format_code: return f"Date ({format_code.upper()})"
#     if '$' in format_code or 'accounting' in format_code: return "Currency"
#     if '%' in format_code: return "Percentage"
#     if format_code == 'general': return "General"
#     if '@' in format_code: return "Text"
#     return f"Custom ({format_code})"

# def get_color_name(rgb_code):
#     if rgb_code is None or rgb_code == '00000000': return "No Fill"
#     color_map = { "FFFFFFFF": "White", "FF000000": "Black", "FFFF0000": "Red", "FF00FF00": "Green", "FF0000FF": "Blue", "FFFFFF00": "Yellow", "FF00B0F0": "Light Blue", "FF92D050": "Light Green", "FFC0C0C0": "Light Gray" }
#     return color_map.get(str(rgb_code).upper(), f"Custom Color (#{rgb_code})")

# def check_specific_datatype(value):
#     if isinstance(value, (int, float)): return "Numeric"
#     if isinstance(value, str):
#         if str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']: return "Binary (Yes/No)"
#         return "Text"
#     return "Other"

# # --- Core comparison logic (no changes here) ---
# def compare_excel_files(input_file, output_file):
#     results = { "overall": {}, "sheets_summary": {} }
#     try:
#         input_wb, output_wb = openpyxl.load_workbook(input_file), openpyxl.load_workbook(output_file)
#         common_sheets = sorted(list(set(input_wb.sheetnames).intersection(set(output_wb.sheetnames))))

#         for sheet_name in common_sheets:
#             results[sheet_name] = {"successes": [], "discrepancies": [], "correct_cells": []}
#             input_ws, output_ws = input_wb[sheet_name], output_wb[sheet_name]
#             input_df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
#             output_df = pd.read_excel(output_file, sheet_name=sheet_name, engine='openpyxl')

#             if list(input_df.columns) != list(output_df.columns):
#                 results[sheet_name]["discrepancies"].append({"Category": "Formatting", "Cell": "Row 1", "Column": "N/A", "Reason": "Column Header Mismatch", "Template_Value": f"{list(input_df.columns)}", "Output_Value": f"{list(output_df.columns)}"})
#             else:
#                 results[sheet_name]["successes"].append("Column Headers are identical.")

#             num_data_rows = min(input_df.shape[0], output_df.shape[0])
#             num_data_cols = min(input_df.shape[1], output_df.shape[1])
#             wrong_cell_coords = set()
#             total_cells_checked = (num_data_rows + 1) * num_data_cols
            
#             for row in range(1, num_data_rows + 2):
#                 for col in range(1, num_data_cols + 1):
#                     input_cell, output_cell = input_ws.cell(row=row, column=col), output_ws.cell(row=row, column=col)
#                     column_name = input_df.columns[col-1] if col-1 < len(input_df.columns) else 'N/A'
#                     error_base = {"Cell": input_cell.coordinate, "Column": column_name}
#                     cell_has_error = False

#                     def add_error(reason, cat, t_val, o_val):
#                         nonlocal cell_has_error; cell_has_error = True
#                         wrong_cell_coords.add(input_cell.coordinate)
#                         results[sheet_name]["discrepancies"].append({**error_base, "Category": cat, "Reason": reason, "Template_Value": t_val, "Output_Value": o_val})

#                     is_overflow = isinstance(output_cell.value, str) and len(output_cell.value) > (output_ws.column_dimensions.get(openpyxl.utils.get_column_letter(col), openpyxl.worksheet.dimensions.ColumnDimension(output_ws)).width or 8.43)
#                     if not is_overflow and input_cell.value != output_cell.value: add_error("Value Mismatch", "Accuracy", input_cell.value, output_cell.value)
#                     if check_specific_datatype(input_cell.value) != check_specific_datatype(output_cell.value): add_error("Type Mismatch", "Datatype", check_specific_datatype(input_cell.value), check_specific_datatype(output_cell.value))
#                     if (input_cell.font.name, input_cell.font.size, input_cell.font.bold) != (output_cell.font.name, output_cell.font.size, output_cell.font.bold): add_error("Font Mismatch", "Formatting", f"{input_cell.font.name}, {input_cell.font.size}pt, Bold:{input_cell.font.bold}", f"{output_cell.font.name}, {output_cell.font.size}pt, Bold:{output_cell.font.bold}")
#                     if input_cell.fill.fgColor.rgb != output_cell.fill.fgColor.rgb: add_error("Background Color", "Formatting", get_color_name(input_cell.fill.fgColor.rgb), get_color_name(output_cell.fill.fgColor.rgb))
#                     if input_cell.number_format != output_cell.number_format: add_error("Number Format", "Formatting", get_number_format_name(input_cell.number_format), get_number_format_name(output_cell.number_format))
#                     if input_cell.alignment.wrap_text != output_cell.alignment.wrap_text: add_error("Text Wrap Setting", "Formatting", f"Wrap: {input_cell.alignment.wrap_text}", f"Wrap: {output_cell.alignment.wrap_text}")
#                     if is_overflow and not output_cell.alignment.wrap_text: col_width = output_ws.column_dimensions.get(openpyxl.utils.get_column_letter(col), openpyxl.worksheet.dimensions.ColumnDimension(output_ws)).width or 8.43; add_error("Text Overflow", "Formatting", f"Fits width of ~{col_width:.0f}", f"Text length is {len(output_cell.value)}")

#                     if not cell_has_error:
#                         results[sheet_name]["correct_cells"].append({"Cell": input_cell.coordinate, "Column": error_base["Column"], "Template_Value": input_cell.value, "Output_Value": output_cell.value})

#             all_cols = set(input_ws.column_dimensions.keys()) | set(output_ws.column_dimensions.keys())
#             for col_letter in all_cols:
#                 input_width = (input_ws.column_dimensions.get(col_letter) or openpyxl.worksheet.dimensions.ColumnDimension(input_ws)).width or 8.43
#                 output_width = (output_ws.column_dimensions.get(col_letter) or openpyxl.worksheet.dimensions.ColumnDimension(output_ws)).width or 8.43
#                 if abs(input_width - output_width) > 0.1:
#                     results[sheet_name]["discrepancies"].append({"Category": "Formatting", "Cell": f"Column {col_letter}", "Column": col_letter, "Reason": "Column Width Mismatch", "Template_Value": f"{input_width:.2f}", "Output_Value": f"{output_width:.2f}"}); wrong_cell_coords.add(f"Col {col_letter}")
#             all_rows = set(input_ws.row_dimensions.keys()) | set(output_ws.row_dimensions.keys())
#             for row_idx in all_rows:
#                 input_height = (input_ws.row_dimensions.get(row_idx) or openpyxl.worksheet.dimensions.RowDimension(input_ws)).height or 15.0
#                 output_height = (output_ws.row_dimensions.get(row_idx) or openpyxl.worksheet.dimensions.RowDimension(output_ws)).height or 15.0
#                 if abs(input_height - output_height) > 0.1:
#                     results[sheet_name]["discrepancies"].append({"Category": "Formatting", "Cell": f"Row {row_idx}", "Column": "N/A", "Reason": "Row Height Mismatch", "Template_Value": f"{input_height:.2f}", "Output_Value": f"{output_height:.2f}"}); wrong_cell_coords.add(f"Row {row_idx}")

#             num_wrong_cells = len(wrong_cell_coords)
#             results[sheet_name]["successes"].append(f"Cell-Level Checks: {total_cells_checked - num_wrong_cells} out of {total_cells_checked} cells were correct.")
#             discrepancies = results[sheet_name]["discrepancies"]
#             results["sheets_summary"][sheet_name] = {"status": "Fail" if discrepancies else "Pass", "total_cells": total_cells_checked, "wrong_cells": num_wrong_cells, "accuracy_errors": len([d for d in discrepancies if d['Category'] == 'Accuracy']), "datatype_errors": len([d for d in discrepancies if d['Category'] == 'Datatype']), "formatting_errors": len([d for d in discrepancies if d['Category'] == 'Formatting'])}
#     except Exception as e: st.error(f"A critical error occurred during comparison: {e}")
#     return results

# # --- UPDATED: The Excel report function now includes the Data Accuracy % KPI ---
# def generate_excel_report(results):
#     output_stream = io.BytesIO()
#     writer = pd.ExcelWriter(output_stream, engine='xlsxwriter')
#     workbook = writer.book

#     # Calculate global metrics
#     sheets_summary = results.get("sheets_summary", {}); total_sheets = len(sheets_summary)
#     passed_sheets = len([s for s, data in sheets_summary.items() if data.get("status") == "Pass"]); failed_sheets = total_sheets - passed_sheets
#     total_cells = sum(d.get('total_cells', 0) for d in sheets_summary.values()); total_wrong = sum(d.get('wrong_cells', 0) for d in sheets_summary.values() if isinstance(d.get('wrong_cells'), int)); total_correct = total_cells - total_wrong
#     cell_pass_percentage = (total_correct / total_cells * 100) if total_cells > 0 else 100
#     total_accuracy_errors = sum(d.get('accuracy_errors', 0) for d in sheets_summary.values()); total_datatype_errors = sum(d.get('datatype_errors', 0) for d in sheets_summary.values()); total_formatting_errors = sum(d.get('formatting_errors', 0) for d in sheets_summary.values())
    
#     # --- NEW: Calculate the Data Accuracy Percentage ---
#     data_accuracy_percentage = ((total_cells - total_accuracy_errors) / total_cells * 100) if total_cells > 0 else 100

#     # Part 1: High-Impact Strategy Dashboard
#     strategy_sheet = workbook.add_worksheet("QA Dashboard")
#     title_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter', 'font_color': '#2F5496'})
#     kpi_format = workbook.add_format({'bold': True, 'font_size': 24, 'align': 'center', 'valign': 'vcenter'})
#     kpi_label_format = workbook.add_format({'font_size': 11, 'align': 'center', 'font_color': '#595959'})
#     header_format = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#DDEBF7', 'border': 1})
#     section_header_format = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#2F5496'})

#     strategy_sheet.merge_range('B2:J3', 'Excel Quality Assurance Dashboard', title_format)
    
#     # --- UPDATED: Dashboard layout to include the new KPI ---
#     strategy_sheet.write('B5', 'Cell-Level Metrics', section_header_format)
#     strategy_sheet.merge_range('B6:D8', f"{data_accuracy_percentage:.1f}%", kpi_format); strategy_sheet.merge_range('B9:D9', 'Data Accuracy %', kpi_label_format)
#     strategy_sheet.merge_range('F6:G8', f"{cell_pass_percentage:.1f}%", kpi_format); strategy_sheet.merge_range('F9:G9', 'Overall Match %', kpi_label_format)
#     strategy_sheet.merge_range('I6:J8', total_wrong, kpi_format); strategy_sheet.merge_range('I9:J9', 'Total Wrong Cells', kpi_label_format)
    
#     strategy_sheet.write('B11', 'Sheet-Level Metrics', section_header_format)
#     strategy_sheet.merge_range('B12:D14', total_sheets, kpi_format); strategy_sheet.merge_range('B15:D15', 'Total Sheets', kpi_label_format)
#     strategy_sheet.merge_range('F12:G14', passed_sheets, kpi_format); strategy_sheet.merge_range('F15:G15', 'Sheets Passed', kpi_label_format)
#     strategy_sheet.merge_range('I12:J14', failed_sheets, kpi_format); strategy_sheet.merge_range('I15:J15', 'Sheets Failed', kpi_label_format)

#     strategy_sheet.merge_range('B18:D18', 'Key Insights', header_format)
#     sheet_with_most_errors, max_errors = "", 0
#     for sheet, data in sheets_summary.items():
#         if data.get('wrong_cells', 0) > max_errors: max_errors, sheet_with_most_errors = data.get('wrong_cells', 0), sheet
#     most_common_error = max({"Accuracy": total_accuracy_errors, "Datatype": total_datatype_errors, "Formatting": total_formatting_errors}, key=lambda k: {"Accuracy": total_accuracy_errors, "Datatype": total_datatype_errors, "Formatting": total_formatting_errors}[k]) if total_wrong > 0 else "None"
    
#     strategy_sheet.write('B19', f"1. Highest Risk Area: The sheet '{sheet_with_most_errors}' has the most errors ({max_errors}). Review first.")
#     strategy_sheet.write('B20', f"2. Systemic Problem: The most common issue category is '{most_common_error}'.")
    
#     error_df = pd.DataFrame([{"Category": "Accuracy Errors", "Count": total_accuracy_errors}, {"Category": "Datatype Errors", "Count": total_datatype_errors}, {"Category": "Formatting Errors", "Count": total_formatting_errors}])
#     error_df.to_excel(writer, sheet_name="QA Dashboard", startrow=21, startcol=1, index=False)
    
#     chart = workbook.add_chart({'type': 'bar'}); chart_data = [total_accuracy_errors, total_datatype_errors, total_formatting_errors]
#     strategy_sheet.write_column('M1', ["Accuracy", "Datatype", "Formatting"]); strategy_sheet.write_column('N1', chart_data)
#     chart.add_series({'name': 'Error Distribution', 'categories': '=\'QA Dashboard\'!$M$1:$M$3', 'values': '=\'QA Dashboard\'!$N$1:$N$3'})
#     chart.set_title({'name': 'Error Breakdown by Category'}); chart.set_legend({'none': True})
#     strategy_sheet.insert_chart('F18', chart, {'x_offset': 25, 'y_offset': 10})
    
#     strategy_sheet.merge_range('B26:D26', 'Key Action Points', header_format)
#     strategy_sheet.write('B27', f"1. Prioritize all 'Accuracy' errors to ensure data correctness.")
#     strategy_sheet.write('B28', f"2. Focus on sheet '{sheet_with_most_errors}' to resolve the highest number of issues quickly.")
#     strategy_sheet.write('B29', f"3. Investigate the root cause of '{most_common_error}' errors, as this is the most systemic problem.")
#     strategy_sheet.set_column('B:J', 15)

#     # Part 2: Summary Sheet
#     pd.DataFrame({"Metric": ["Overall Cell Match Percentage", "Total Cells Checked", "Total Correct Cells", "Total Wrong Cells", "Total Sheets Checked", "Sheets Passed", "Sheets Failed"],"Value": [f"{cell_pass_percentage:.2f}%", total_cells, total_correct, total_wrong, total_sheets, passed_sheets, failed_sheets]}).to_excel(writer, sheet_name='Summary', startrow=1, index=False, header=False)
#     summary_sheet = writer.sheets['Summary']; summary_sheet.write('A1', 'Overall Test Summary'); summary_sheet.set_column('A:A', 30); summary_sheet.set_column('B:B', 20)
#     categorized_data = [{"Sheet Name": n, "Status": d.get("status"), "Accuracy Errors": d.get("accuracy_errors", 0), "Datatype Errors": d.get("datatype_errors", 0), "Formatting Errors": d.get("formatting_errors", 0)} for n, d in sheets_summary.items()]
#     summary_df2 = pd.DataFrame(categorized_data); summary_df2.to_excel(writer, sheet_name='Summary', startrow=9, index=False)
#     summary_sheet.write('A9', 'Per-Sheet Error Breakdown'); summary_sheet.set_column('C:F', 20)

#     # Part 3: Detailed Sheets with Separated Tables
#     for sheet_name, sheet_results in results.items():
#         if sheet_name in ["overall", "sheets_summary"]: continue
#         report_sheet_name = f"{sheet_name}_Report"; worksheet = workbook.add_worksheet(report_sheet_name); current_row = 0
#         incorrect_cells = pd.DataFrame(sheet_results.get("discrepancies", []))
#         if not incorrect_cells.empty:
#             worksheet.write(current_row, 0, "Incorrect Cells", header_format); current_row += 1
#             incorrect_cells[["Cell", "Column", "Reason", "Template_Value", "Output_Value"]].to_excel(writer, sheet_name=report_sheet_name, startrow=current_row, index=False); current_row += len(incorrect_cells) + 2
#         correct_cells = pd.DataFrame(sheet_results.get("correct_cells", []))
#         if not correct_cells.empty:
#             worksheet.write(current_row, 0, "Correct Cells", header_format); current_row += 1
#             correct_cells[["Cell", "Column", "Template_Value", "Output_Value"]].to_excel(writer, sheet_name=report_sheet_name, startrow=current_row, index=False)
#         worksheet.set_column('A:F', 25)
        
#     writer.close(); output_stream.seek(0)
#     return output_stream

# # --- UI (no changes needed) ---
# st.set_page_config(page_title="Advanced Excel Comparator", layout="wide"); st.title("Advanced Excel Quality Assurance Tool")
# input_file = st.file_uploader("Upload Input (Template) Excel", type=['xlsx']); output_file = st.file_uploader("Upload Output (To Test) Excel", type=['xlsx'])
# if input_file and output_file:
#     if st.button("Run Advanced Comparison", type="primary"):
#         with st.spinner("Performing deep analysis..."): st.session_state['results'] = compare_excel_files(input_file, output_file)
#         st.session_state['ran_comparison'] = True
# if 'ran_comparison' in st.session_state:
#     results = st.session_state['results']; sheets_summary = results.get("sheets_summary", {});
#     total_sheets = len(sheets_summary); passed_sheets = len([s for s, data in sheets_summary.items() if data.get("status") == "Pass"]); failed_sheets = total_sheets - passed_sheets
#     total_cells = sum(d.get('total_cells', 0) for d in sheets_summary.values()); total_wrong = sum(d.get('wrong_cells', 0) for d in sheets_summary.values() if isinstance(d.get('wrong_cells'), int)); total_correct = total_cells - total_wrong
#     cell_pass_percentage = (total_correct / total_cells * 100) if total_cells > 0 else 100
#     st.header("On-Screen Comparison Summary")
#     st.subheader("Cell-Level Summary")
#     col1, col2, col3 = st.columns(3)
#     col1.metric(label="📊 Overall Cell Match Percentage", value=f"{cell_pass_percentage:.2f}%"); col2.metric(label="✅ Correct Cells", value=total_correct); col3.metric(label="❌ Wrong Cells", value=total_wrong, delta_color="inverse")
#     st.subheader("Sheet-Level Summary")
#     col4, col5, col6 = st.columns(3)
#     col4.metric(label="📋 Total Sheets Checked", value=total_sheets); col5.metric(label="✅ Sheets Passed", value=passed_sheets); col6.metric(label="❌ Sheets Failed", value=failed_sheets, delta_color="inverse")
#     for sheet_name, sheet_results in results.items():
#         if sheet_name in ["overall", "sheets_summary"]: continue
#         with st.expander(f"Results for Sheet: '{sheet_name}'", expanded=bool(sheet_results.get('discrepancies'))):
#             st.subheader("✅ Passed Checks")
#             if sheet_results.get('successes'):
#                 for success in sheet_results['successes']: st.markdown(f"- {success}")
#             else: st.write("No specific checks passed for this sheet.")
#             st.subheader("❌ Discrepancies Found")
#             discrepancy_list = sheet_results.get('discrepancies', [])
#             if discrepancy_list:
#                 categories = {"Accuracy": [], "Datatype": [], "Formatting": []}
#                 for error in discrepancy_list: categories.get(error.get("Category", "Other"), []).append(error)
#                 for category, errors in categories.items():
#                     if errors:
#                         st.markdown(f"**{category} Errors ({len(errors)} found):**")
#                         for error in errors: st.markdown(f"  - **Cell:** `{error['Cell']}` | **Reason:** *{error['Reason']}* | **Template:** `{error['Template_Value']}` | **Output:** `{error['Output_Value']}`")
#             else: st.write("No discrepancies found in this sheet.")
#     st.markdown("---"); st.download_button(label="📄 Download Full Categorized Report (Excel)", data=generate_excel_report(results), file_name=f"Advanced_Test_Report_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# import streamlit as st
# import pandas as pd
# import openpyxl
# import io
# from datetime import datetime

# # --- Helper functions (no changes here) ---
# def get_number_format_name(format_code):
#     if format_code is None: return "General"
#     format_code = str(format_code).lower()
#     if 'yyyy' in format_code or 'd' in format_code or 'm' in format_code: return f"Date ({format_code.upper()})"
#     if '$' in format_code or 'accounting' in format_code: return "Currency"
#     if '%' in format_code: return "Percentage"
#     if format_code == 'general': return "General"
#     if '@' in format_code: return "Text"
#     return f"Custom ({format_code})"

# def get_color_name(rgb_code):
#     if rgb_code is None or rgb_code == '00000000': return "No Fill"
#     color_map = { "FFFFFFFF": "White", "FF000000": "Black", "FFFF0000": "Red", "FF00FF00": "Green", "FF0000FF": "Blue", "FFFFFF00": "Yellow", "FF00B0F0": "Light Blue", "FF92D050": "Light Green", "FFC0C0C0": "Light Gray" }
#     return color_map.get(str(rgb_code).upper(), f"Custom Color (#{rgb_code})")

# def check_specific_datatype(value):
#     if isinstance(value, (int, float)): return "Numeric"
#     if isinstance(value, str):
#         if str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']: return "Binary (Yes/No)"
#         return "Text"
#     return "Other"

# # --- Core comparison logic (no changes here) ---
# def compare_excel_files(input_file, output_file):
#     results = { "overall": {}, "sheets_summary": {} }
#     try:
#         input_wb, output_wb = openpyxl.load_workbook(input_file), openpyxl.load_workbook(output_file)
#         common_sheets = sorted(list(set(input_wb.sheetnames).intersection(set(output_wb.sheetnames))))

#         for sheet_name in common_sheets:
#             results[sheet_name] = {"successes": [], "discrepancies": [], "correct_cells": []}
#             input_ws, output_ws = input_wb[sheet_name], output_wb[sheet_name]
#             input_df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
#             output_df = pd.read_excel(output_file, sheet_name=sheet_name, engine='openpyxl')

#             if list(input_df.columns) != list(output_df.columns):
#                 results[sheet_name]["discrepancies"].append({"Category": "Formatting", "Cell": "Row 1", "Column": "N/A", "Reason": "Column Header Mismatch", "Template_Value": f"{list(input_df.columns)}", "Output_Value": f"{list(output_df.columns)}"})
#             else:
#                 results[sheet_name]["successes"].append("Column Headers are identical.")

#             num_data_rows = min(input_df.shape[0], output_df.shape[0])
#             num_data_cols = min(input_df.shape[1], output_df.shape[1])
#             wrong_cell_coords = set()
#             total_cells_checked = (num_data_rows + 1) * num_data_cols
            
#             for row in range(1, num_data_rows + 2):
#                 for col in range(1, num_data_cols + 1):
#                     input_cell, output_cell = input_ws.cell(row=row, column=col), output_ws.cell(row=row, column=col)
#                     column_name = input_df.columns[col-1] if col-1 < len(input_df.columns) else 'N/A'
#                     error_base = {"Cell": input_cell.coordinate, "Column": column_name}
#                     cell_has_error = False

#                     def add_error(reason, cat, t_val, o_val):
#                         nonlocal cell_has_error; cell_has_error = True
#                         wrong_cell_coords.add(input_cell.coordinate)
#                         results[sheet_name]["discrepancies"].append({**error_base, "Category": cat, "Reason": reason, "Template_Value": t_val, "Output_Value": o_val})

#                     is_overflow = isinstance(output_cell.value, str) and len(output_cell.value) > (output_ws.column_dimensions.get(openpyxl.utils.get_column_letter(col), openpyxl.worksheet.dimensions.ColumnDimension(output_ws)).width or 8.43)
#                     if not is_overflow and input_cell.value != output_cell.value: add_error("Value Mismatch", "Accuracy", input_cell.value, output_cell.value)
#                     if check_specific_datatype(input_cell.value) != check_specific_datatype(output_cell.value): add_error("Type Mismatch", "Datatype", check_specific_datatype(input_cell.value), check_specific_datatype(output_cell.value))
#                     if (input_cell.font.name, input_cell.font.size, input_cell.font.bold) != (output_cell.font.name, output_cell.font.size, output_cell.font.bold): add_error("Font Mismatch", "Formatting", f"{input_cell.font.name}, {input_cell.font.size}pt, Bold:{input_cell.font.bold}", f"{output_cell.font.name}, {output_cell.font.size}pt, Bold:{output_cell.font.bold}")
#                     if input_cell.fill.fgColor.rgb != output_cell.fill.fgColor.rgb: add_error("Background Color", "Formatting", get_color_name(input_cell.fill.fgColor.rgb), get_color_name(output_cell.fill.fgColor.rgb))
#                     if input_cell.number_format != output_cell.number_format: add_error("Number Format", "Formatting", get_number_format_name(input_cell.number_format), get_number_format_name(output_cell.number_format))
#                     if input_cell.alignment.wrap_text != output_cell.alignment.wrap_text: add_error("Text Wrap Setting", "Formatting", f"Wrap: {input_cell.alignment.wrap_text}", f"Wrap: {output_cell.alignment.wrap_text}")
#                     if is_overflow and not output_cell.alignment.wrap_text: col_width = output_ws.column_dimensions.get(openpyxl.utils.get_column_letter(col), openpyxl.worksheet.dimensions.ColumnDimension(output_ws)).width or 8.43; add_error("Text Overflow", "Formatting", f"Fits width of ~{col_width:.0f}", f"Text length is {len(output_cell.value)}")

#                     if not cell_has_error:
#                         results[sheet_name]["correct_cells"].append({"Cell": input_cell.coordinate, "Column": error_base["Column"], "Template_Value": input_cell.value, "Output_Value": output_cell.value})

#             all_cols = set(input_ws.column_dimensions.keys()) | set(output_ws.column_dimensions.keys())
#             for col_letter in all_cols:
#                 input_width = (input_ws.column_dimensions.get(col_letter) or openpyxl.worksheet.dimensions.ColumnDimension(input_ws)).width or 8.43
#                 output_width = (output_ws.column_dimensions.get(col_letter) or openpyxl.worksheet.dimensions.ColumnDimension(output_ws)).width or 8.43
#                 if abs(input_width - output_width) > 0.1:
#                     results[sheet_name]["discrepancies"].append({"Category": "Formatting", "Cell": f"Column {col_letter}", "Column": col_letter, "Reason": "Column Width Mismatch", "Template_Value": f"{input_width:.2f}", "Output_Value": f"{output_width:.2f}"}); wrong_cell_coords.add(f"Col {col_letter}")
#             all_rows = set(input_ws.row_dimensions.keys()) | set(output_ws.row_dimensions.keys())
#             for row_idx in all_rows:
#                 input_height = (input_ws.row_dimensions.get(row_idx) or openpyxl.worksheet.dimensions.RowDimension(input_ws)).height or 15.0
#                 output_height = (output_ws.row_dimensions.get(row_idx) or openpyxl.worksheet.dimensions.RowDimension(output_ws)).height or 15.0
#                 if abs(input_height - output_height) > 0.1:
#                     results[sheet_name]["discrepancies"].append({"Category": "Formatting", "Cell": f"Row {row_idx}", "Column": "N/A", "Reason": "Row Height Mismatch", "Template_Value": f"{input_height:.2f}", "Output_Value": f"{output_height:.2f}"}); wrong_cell_coords.add(f"Row {row_idx}")

#             num_wrong_cells = len(wrong_cell_coords)
#             results[sheet_name]["successes"].append(f"Cell-Level Checks: {total_cells_checked - num_wrong_cells} out of {total_cells_checked} cells were correct.")
#             discrepancies = results[sheet_name]["discrepancies"]
#             results["sheets_summary"][sheet_name] = {"status": "Fail" if discrepancies else "Pass", "total_cells": total_cells_checked, "wrong_cells": num_wrong_cells, "accuracy_errors": len([d for d in discrepancies if d['Category'] == 'Accuracy']), "datatype_errors": len([d for d in discrepancies if d['Category'] == 'Datatype']), "formatting_errors": len([d for d in discrepancies if d['Category'] == 'Formatting'])}
#     except Exception as e: st.error(f"A critical error occurred during comparison: {e}")
#     return results

# # --- UPDATED: The Excel report function restores all original KPIs and adds the new one ---
# def generate_excel_report(results):
#     output_stream = io.BytesIO()
#     writer = pd.ExcelWriter(output_stream, engine='xlsxwriter')
#     workbook = writer.book

#     # Calculate global metrics
#     sheets_summary = results.get("sheets_summary", {}); total_sheets = len(sheets_summary)
#     passed_sheets = len([s for s, data in sheets_summary.items() if data.get("status") == "Pass"]); failed_sheets = total_sheets - passed_sheets
#     total_cells = sum(d.get('total_cells', 0) for d in sheets_summary.values()); 
#     total_accuracy_errors = sum(d.get('accuracy_errors', 0) for d in sheets_summary.values()); 
#     total_datatype_errors = sum(d.get('datatype_errors', 0) for d in sheets_summary.values()); 
#     total_formatting_errors = sum(d.get('formatting_errors', 0) for d in sheets_summary.values())
#     total_wrong = total_accuracy_errors + total_datatype_errors + total_formatting_errors
#     total_correct = total_cells - total_wrong
    
#     cell_pass_percentage = (total_correct / total_cells * 100) if total_cells > 0 else 100
#     data_accuracy_percentage = ((total_cells - total_accuracy_errors) / total_cells * 100) if total_cells > 0 else 100
#     formatting_accuracy_percentage = ((total_cells - total_formatting_errors) / total_cells * 100) if total_cells > 0 else 100

#     # Part 1: High-Impact Strategy Dashboard
#     strategy_sheet = workbook.add_worksheet("QA Dashboard")
#     title_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter', 'font_color': '#2F5496'})
#     kpi_format = workbook.add_format({'bold': True, 'font_size': 24, 'align': 'center', 'valign': 'vcenter'})
#     kpi_label_format = workbook.add_format({'font_size': 11, 'align': 'center', 'font_color': '#595959'})
#     header_format = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#DDEBF7', 'border': 1})
#     section_header_format = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#2F5496'})

#     strategy_sheet.merge_range('B2:K3', 'Excel Quality Assurance Dashboard', title_format)
    
#     # --- UPDATED: Dashboard layout to include all KPIs without removing any ---
#     strategy_sheet.write('B5', 'Cell-Level Metrics', section_header_format)
#     strategy_sheet.merge_range('B6:C8', f"{cell_pass_percentage:.1f}%", kpi_format); strategy_sheet.merge_range('B9:C9', 'Overall Cell Match %', kpi_label_format)
#     strategy_sheet.merge_range('D6:E8', f"{data_accuracy_percentage:.1f}%", kpi_format); strategy_sheet.merge_range('D9:E9', 'Data Accuracy %', kpi_label_format)
#     strategy_sheet.merge_range('F6:G8', f"{formatting_accuracy_percentage:.1f}%", kpi_format); strategy_sheet.merge_range('F9:G9', 'Formatting Accuracy %', kpi_label_format)
#     strategy_sheet.merge_range('H6:I8', total_correct, kpi_format); strategy_sheet.merge_range('H9:I9', 'Correct Cells', kpi_label_format)
#     strategy_sheet.merge_range('J6:K8', total_wrong, kpi_format); strategy_sheet.merge_range('J9:K9', 'Wrong Cells', kpi_label_format)
    
#     # Adjusting subsequent sections to fit the new layout
#     strategy_sheet.write('B11', 'Sheet-Level Metrics', section_header_format)
#     strategy_sheet.merge_range('B12:C14', total_sheets, kpi_format); strategy_sheet.merge_range('B15:C15', 'Total Sheets', kpi_label_format)
#     strategy_sheet.merge_range('E12:F14', passed_sheets, kpi_format); strategy_sheet.merge_range('E15:F15', 'Sheets Passed', kpi_label_format)
#     strategy_sheet.merge_range('H12:I14', failed_sheets, kpi_format); strategy_sheet.merge_range('H15:I15', 'Sheets Failed', kpi_label_format)

#     strategy_sheet.merge_range('B18:D18', 'Key Insights', header_format)
#     sheet_with_most_errors, max_errors = "", 0
#     for sheet, data in sheets_summary.items():
#         if data.get('wrong_cells', 0) > max_errors: max_errors, sheet_with_most_errors = data.get('wrong_cells', 0), sheet
#     most_common_error = max({"Accuracy": total_accuracy_errors, "Datatype": total_datatype_errors, "Formatting": total_formatting_errors}, key=lambda k: {"Accuracy": total_accuracy_errors, "Datatype": total_datatype_errors, "Formatting": total_formatting_errors}[k]) if total_wrong > 0 else "None"
    
#     strategy_sheet.write('B19', f"1. Highest Risk Area: The sheet '{sheet_with_most_errors}' has the most errors ({max_errors}). Review first.")
#     strategy_sheet.write('B20', f"2. Systemic Problem: The most common issue category is '{most_common_error}'.")
    
#     error_df = pd.DataFrame([{"Category": "Accuracy Errors", "Count": total_accuracy_errors}, {"Category": "Datatype Errors", "Count": total_datatype_errors}, {"Category": "Formatting Errors", "Count": total_formatting_errors}])
#     error_df.to_excel(writer, sheet_name="QA Dashboard", startrow=21, startcol=1, index=False)
    
#     chart = workbook.add_chart({'type': 'bar'}); chart_data = [total_accuracy_errors, total_datatype_errors, total_formatting_errors]
#     strategy_sheet.write_column('M1', ["Accuracy", "Datatype", "Formatting"]); strategy_sheet.write_column('N1', chart_data)
#     chart.add_series({'name': 'Error Distribution', 'categories': '=\'QA Dashboard\'!$M$1:$M$3', 'values': '=\'QA Dashboard\'!$N$1:$N$3'})
#     chart.set_title({'name': 'Error Breakdown by Category'}); chart.set_legend({'none': True})
#     strategy_sheet.insert_chart('F18', chart, {'x_offset': 25, 'y_offset': 10})
    
#     strategy_sheet.merge_range('B26:D26', 'Key Action Points', header_format)
#     strategy_sheet.write('B27', f"1. Prioritize all 'Accuracy' errors to ensure data correctness.")
#     strategy_sheet.write('B28', f"2. Focus on sheet '{sheet_with_most_errors}' to resolve the highest number of issues quickly.")
#     strategy_sheet.write('B29', f"3. Investigate the root cause of '{most_common_error}' errors, as this is the most systemic problem.")
#     strategy_sheet.set_column('B:L', 15)

#     # Part 2: Summary Sheet
#     pd.DataFrame({"Metric": ["Overall Cell Match Percentage", "Total Cells Checked", "Total Correct Cells", "Total Wrong Cells", "Total Sheets Checked", "Sheets Passed", "Sheets Failed"],"Value": [f"{cell_pass_percentage:.2f}%", total_cells, total_correct, total_wrong, total_sheets, passed_sheets, failed_sheets]}).to_excel(writer, sheet_name='Summary', startrow=1, index=False, header=False)
#     summary_sheet = writer.sheets['Summary']; summary_sheet.write('A1', 'Overall Test Summary'); summary_sheet.set_column('A:A', 30); summary_sheet.set_column('B:B', 20)
#     categorized_data = [{"Sheet Name": n, "Status": d.get("status"), "Accuracy Errors": d.get("accuracy_errors", 0), "Datatype Errors": d.get("datatype_errors", 0), "Formatting Errors": d.get("formatting_errors", 0)} for n, d in sheets_summary.items()]
#     summary_df2 = pd.DataFrame(categorized_data); summary_df2.to_excel(writer, sheet_name='Summary', startrow=9, index=False)
#     summary_sheet.write('A9', 'Per-Sheet Error Breakdown'); summary_sheet.set_column('C:F', 20)

#     # Part 3: Detailed Sheets with Separated Tables
#     for sheet_name, sheet_results in results.items():
#         if sheet_name in ["overall", "sheets_summary"]: continue
#         report_sheet_name = f"{sheet_name}_Report"; worksheet = workbook.add_worksheet(report_sheet_name); current_row = 0
#         incorrect_cells = pd.DataFrame(sheet_results.get("discrepancies", []))
#         if not incorrect_cells.empty:
#             worksheet.write(current_row, 0, "Incorrect Cells", header_format); current_row += 1
#             incorrect_cells[["Cell", "Column", "Reason", "Template_Value", "Output_Value"]].to_excel(writer, sheet_name=report_sheet_name, startrow=current_row, index=False); current_row += len(incorrect_cells) + 2
#         correct_cells = pd.DataFrame(sheet_results.get("correct_cells", []))
#         if not correct_cells.empty:
#             worksheet.write(current_row, 0, "Correct Cells", header_format); current_row += 1
#             correct_cells[["Cell", "Column", "Template_Value", "Output_Value"]].to_excel(writer, sheet_name=report_sheet_name, startrow=current_row, index=False)
#         worksheet.set_column('A:F', 25)
        
#     writer.close(); output_stream.seek(0)
#     return output_stream

# # --- UI (no changes needed) ---
# st.set_page_config(page_title="Advanced Excel Comparator", layout="wide"); st.title("Advanced Excel Quality Assurance Tool")
# input_file = st.file_uploader("Upload Input (Template) Excel", type=['xlsx']); output_file = st.file_uploader("Upload Output (To Test) Excel", type=['xlsx'])
# if input_file and output_file:
#     if st.button("Run Advanced Comparison", type="primary"):
#         with st.spinner("Performing deep analysis..."): st.session_state['results'] = compare_excel_files(input_file, output_file)
#         st.session_state['ran_comparison'] = True
# if 'ran_comparison' in st.session_state:
#     results = st.session_state['results']; sheets_summary = results.get("sheets_summary", {});
#     total_sheets = len(sheets_summary); passed_sheets = len([s for s, data in sheets_summary.items() if data.get("status") == "Pass"]); failed_sheets = total_sheets - passed_sheets
#     total_cells = sum(d.get('total_cells', 0) for d in sheets_summary.values()); total_wrong = sum(d.get('wrong_cells', 0) for d in sheets_summary.values() if isinstance(d.get('wrong_cells'), int)); total_correct = total_cells - total_wrong
#     cell_pass_percentage = (total_correct / total_cells * 100) if total_cells > 0 else 100
#     st.header("On-Screen Comparison Summary")
#     st.subheader("Cell-Level Summary")
#     col1, col2, col3 = st.columns(3)
#     col1.metric(label="📊 Overall Cell Match Percentage", value=f"{cell_pass_percentage:.2f}%"); col2.metric(label="✅ Correct Cells", value=total_correct); col3.metric(label="❌ Wrong Cells", value=total_wrong, delta_color="inverse")
#     st.subheader("Sheet-Level Summary")
#     col4, col5, col6 = st.columns(3)
#     col4.metric(label="📋 Total Sheets Checked", value=total_sheets); col5.metric(label="✅ Sheets Passed", value=passed_sheets); col6.metric(label="❌ Sheets Failed", value=failed_sheets, delta_color="inverse")
#     for sheet_name, sheet_results in results.items():
#         if sheet_name in ["overall", "sheets_summary"]: continue
#         with st.expander(f"Results for Sheet: '{sheet_name}'", expanded=bool(sheet_results.get('discrepancies'))):
#             st.subheader("✅ Passed Checks")
#             if sheet_results.get('successes'):
#                 for success in sheet_results['successes']: st.markdown(f"- {success}")
#             else: st.write("No specific checks passed for this sheet.")
#             st.subheader("❌ Discrepancies Found")
#             discrepancy_list = sheet_results.get('discrepancies', [])
#             if discrepancy_list:
#                 categories = {"Accuracy": [], "Datatype": [], "Formatting": []}
#                 for error in discrepancy_list: categories.get(error.get("Category", "Other"), []).append(error)
#                 for category, errors in categories.items():
#                     if errors:
#                         st.markdown(f"**{category} Errors ({len(errors)} found):**")
#                         for error in errors: st.markdown(f"  - **Cell:** `{error['Cell']}` | **Reason:** *{error['Reason']}* | **Template:** `{error['Template_Value']}` | **Output:** `{error['Output_Value']}`")
#             else: st.write("No discrepancies found in this sheet.")
#     st.markdown("---"); st.download_button(label="📄 Download Full Categorized Report (Excel)", data=generate_excel_report(results), file_name=f"Advanced_Test_Report_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# import streamlit as st
# import pandas as pd
# import openpyxl
# import io
# from datetime import datetime

# # --- Helper functions ---

# def detect_data_type_row(worksheet):
#     """Scans the first 15 rows to find the one defining data types."""
#     type_keywords = {'numeric', 'text', 'binary', 'decimal', 'date'}
#     for i, row in enumerate(worksheet.iter_rows(min_row=1, max_row=15, values_only=True)):
#         matches = sum(1 for cell_value in row if isinstance(cell_value, str) and any(keyword in cell_value.lower() for keyword in type_keywords))
#         # If at least 3 cells in the row match our keywords, we've found the rule row
#         if matches > 2:
#             return i + 1  # Return 1-based row index
#     return None

# def validate_value_against_type(rule, value):
#     """Checks if a value from the output file conforms to the data type rule from the template."""
#     # An empty cell is considered valid as it has no data to be "wrong"
#     if value is None or str(value).strip() == '': 
#         return True
    
#     if not isinstance(rule, str):
#         return False
        
#     rule = rule.lower().strip()
    
#     # This correctly handles numbers that are stored as text (e.g., "30")
#     if "numeric" in rule or "decimal" in rule:
#         if isinstance(value, (int, float)):
#             return True
#         try:
#             float(str(value))
#             return True
#         except (ValueError, TypeError):
#             return False
            
#     if "text" in rule:
#         # A pure text rule should not accept numbers unless they are explicitly strings.
#         # This prevents a numeric `30` from passing a `Text` rule.
#         return isinstance(value, str)
        
#     if "binary" in rule:
#         return str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']
        
#     if "date" in rule:
#         return isinstance(value, datetime)
        
#     return False

# # --- Core logic focused ONLY on Data Type Accuracy ---
# def compare_excel_files(input_file, output_file):
#     results = { "overall": {}, "sheets_summary": {} }
#     try:
#         # We use data_only=True to get the actual values from cells, not formulas
#         input_wb = openpyxl.load_workbook(input_file, data_only=True)
#         output_wb = openpyxl.load_workbook(output_file, data_only=True)
#         common_sheets = sorted(list(set(input_wb.sheetnames).intersection(set(output_wb.sheetnames))))

#         for sheet_name in common_sheets:
#             results[sheet_name] = {"successes": [], "discrepancies": [], "correct_cells": []}
#             input_ws = input_wb[sheet_name]
#             output_ws = output_wb[sheet_name]
            
#             type_rule_row_index = detect_data_type_row(input_ws)
#             if type_rule_row_index is None:
#                 st.warning(f"Could not detect a data type definition row in template sheet '{sheet_name}'. Skipping.")
#                 continue
                
#             rules = {col: input_ws.cell(row=type_rule_row_index, column=col).value for col in range(1, input_ws.max_column + 1)}
            
#             # Use pandas to get a reliable count of data rows, assuming headers are just above the rule row
#             output_df = pd.read_excel(output_file, sheet_name=sheet_name, header=type_rule_row_index - 1)
#             num_data_rows = output_df.shape[0]
#             num_data_cols = output_df.shape[1]
#             total_cells_checked = num_data_rows * num_data_cols
            
#             for r_idx in range(num_data_rows):
#                 for c_idx in range(num_data_cols):
#                     # Calculate the exact cell coordinates for openpyxl
#                     output_row = type_rule_row_index + 1 + r_idx
#                     output_col = c_idx + 1
                    
#                     output_cell = output_ws.cell(row=output_row, column=output_col)
#                     rule_string = rules.get(output_col)
                    
#                     column_name = output_df.columns[c_idx] if c_idx < len(output_df.columns) else 'N/A'
#                     error_base = {"Cell": output_cell.coordinate, "Column": column_name}
                    
#                     # Perform the single, focused check for data type accuracy
#                     if not validate_value_against_type(rule_string, output_cell.value):
#                         results[sheet_name]["discrepancies"].append({
#                             **error_base, 
#                             "Category": "Datatype",
#                             "Reason": "Datatype Mismatch", 
#                             "Template_Value": f"Rule: '{rule_string}'", 
#                             "Output_Value": f"Value: '{output_cell.value}' (Type: {type(output_cell.value).__name__})"
#                         })
#                     else:
#                         results[sheet_name]["correct_cells"].append({
#                             **error_base, 
#                             "Template_Value": f"Rule: '{rule_string}'", 
#                             "Output_Value": output_cell.value
#                         })

#             discrepancies = results[sheet_name]["discrepancies"]
#             results["sheets_summary"][sheet_name] = {
#                 "status": "Fail" if discrepancies else "Pass", 
#                 "total_cells": total_cells_checked, 
#                 "datatype_errors": len(discrepancies)
#             }
#     except Exception as e:
#         st.error(f"A critical error occurred: {e}")
#     return results

# # --- Simplified report generation logic for this focused test ---
# def generate_excel_report(results):
#     output_stream = io.BytesIO()
#     writer = pd.ExcelWriter(output_stream, engine='xlsxwriter')
#     workbook = writer.book

#     sheets_summary = results.get("sheets_summary", {});
#     total_cells = sum(d.get('total_cells', 0) for d in sheets_summary.values() if d); 
#     total_datatype_errors = sum(d.get('datatype_errors', 0) for d in sheets_summary.values() if d); 
#     total_correct = total_cells - total_datatype_errors
#     datatype_accuracy_percentage = (total_correct / total_cells * 100) if total_cells > 0 else 100

#     strategy_sheet = workbook.add_worksheet("QA Dashboard")
#     title_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter'})
#     kpi_format = workbook.add_format({'bold': True, 'font_size': 24, 'align': 'center', 'valign': 'vcenter'})
#     kpi_label_format = workbook.add_format({'font_size': 11, 'align': 'center', 'font_color': '#595959'})
#     header_format = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#DDEBF7', 'border': 1})
    
#     strategy_sheet.merge_range('B2:H3', 'Datatype Accuracy Dashboard', title_format)
#     strategy_sheet.merge_range('B5:C7', f"{datatype_accuracy_percentage:.1f}%", kpi_format); strategy_sheet.merge_range('B8:C8', 'Datatype Accuracy %', kpi_label_format)
#     strategy_sheet.merge_range('E5:F7', total_correct, kpi_format); strategy_sheet.merge_range('E8:F8', 'Correct Cells', kpi_label_format)
#     strategy_sheet.merge_range('H5:I7', total_datatype_errors, kpi_format); strategy_sheet.merge_range('H8:I8', 'Datatype Errors', kpi_label_format)
#     strategy_sheet.set_column('B:I', 20)
    
#     for sheet_name, sheet_results in results.items():
#         if sheet_name in ["overall", "sheets_summary"]: continue
#         report_sheet_name = f"{sheet_name}_Report"; worksheet = workbook.add_worksheet(report_sheet_name); current_row = 0
        
#         incorrect_cells = pd.DataFrame(sheet_results.get("discrepancies", []))
#         if not incorrect_cells.empty:
#             worksheet.write(current_row, 0, "Datatype Errors", header_format); current_row += 1
#             incorrect_cells[["Cell", "Column", "Reason", "Template_Value", "Output_Value"]].to_excel(writer, sheet_name=report_sheet_name, startrow=current_row, index=False); current_row += len(incorrect_cells) + 2
            
#         correct_cells = pd.DataFrame(sheet_results.get("correct_cells", []))
#         if not correct_cells.empty:
#             worksheet.write(current_row, 0, "Correct Cells", header_format); current_row += 1
#             correct_cells[["Cell", "Column", "Template_Value", "Output_Value"]].to_excel(writer, sheet_name=report_sheet_name, startrow=current_row, index=False)
            
#         worksheet.set_column('A:F', 25)
        
#     writer.close(); output_stream.seek(0)
#     return output_stream

# # --- Simplified UI for this focused test ---
# st.set_page_config(page_title="Excel Datatype Validator", layout="wide")
# st.title("Excel Datatype Validator Tool")

# if 'ran_comparison' not in st.session_state:
#     st.session_state.ran_comparison = False

# input_file = st.file_uploader("Upload Input (Master Template) Excel", type=['xlsx'])
# output_file = st.file_uploader("Upload Output (To Test) Excel", type=['xlsx'])

# if input_file and output_file:
#     if st.button("Run Datatype Validation", type="primary"):
#         with st.spinner("Validating data types..."):
#             st.session_state['results'] = compare_excel_files(input_file, output_file)
#         st.session_state['ran_comparison'] = True

# if st.session_state.ran_comparison:
#     results = st.session_state.get('results', {})
#     sheets_summary = results.get("sheets_summary", {});
#     total_cells = sum(d.get('total_cells', 0) for d in sheets_summary.values() if d); 
#     total_wrong = sum(d.get('datatype_errors', 0) for d in sheets_summary.values() if d)
#     total_correct = total_cells - total_wrong
#     datatype_accuracy_percentage = (total_correct / total_cells * 100) if total_cells > 0 else 100
    
#     st.header("On-Screen Validation Summary")
#     col1, col2, col3 = st.columns(3)
#     col1.metric(label="📊 Datatype Accuracy Percentage", value=f"{datatype_accuracy_percentage:.2f}%")
#     col2.metric(label="✅ Correct Cells", value=total_correct)
#     col3.metric(label="❌ Datatype Errors", value=total_wrong, delta_color="inverse")
    
#     for sheet_name, sheet_results in results.items():
#         if sheet_name in ["overall", "sheets_summary"]: continue
#         with st.expander(f"Results for Sheet: '{sheet_name}'", expanded=bool(sheet_results.get('discrepancies'))):
#             st.subheader("❌ Datatype Errors Found")
#             discrepancy_list = sheet_results.get('discrepancies', [])
#             if discrepancy_list:
#                 for error in discrepancy_list:
#                     st.markdown(f"  - **Cell:** `{error['Cell']}` | **Rule:** `{error['Template_Value']}` | **Actual Value:** `{error['Output_Value']}`")
#             else:
#                 st.success("No datatype errors found in this sheet.")
            
#     st.markdown("---")
#     st.download_button(
#         label="📄 Download Datatype Report (Excel)",
#         data=generate_excel_report(results),
#         file_name=f"Datatype_Test_Report_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx",
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#     )

# import streamlit as st
# import pandas as pd
# import openpyxl
# import io
# from datetime import datetime

# # --- Helper functions ---

# def check_specific_datatype(value):
#     """Determines the specific data type of a given value."""
#     if isinstance(value, (int, float)): return "Numeric"
#     if isinstance(value, str):
#         if str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']: return "Binary(Yes/No)"
#         return "Text"
#     if isinstance(value, datetime): return "Date"
#     return "Other"

# def detect_data_type_row(worksheet):
#     """Scans the first 15 rows to find the one defining data types."""
#     type_keywords = {'numeric', 'text', 'binary', 'decimal', 'date'}
#     for i, row in enumerate(worksheet.iter_rows(min_row=1, max_row=15, values_only=True)):
#         matches = sum(1 for cell_value in row if isinstance(cell_value, str) and any(keyword in cell_value.lower() for keyword in type_keywords))
#         if matches > 2:
#             return i + 1
#     return None

# def validate_value_against_type(rule, value):
#     """Checks if a value conforms to a data type rule string."""
#     if value is None or str(value).strip() == '': return True
#     if not isinstance(rule, str): return False
#     rule = rule.lower().strip()
#     if "numeric" in rule or "decimal" in rule:
#         if isinstance(value, (int, float)): return True
#         try:
#             float(str(value)); return True
#         except (ValueError, TypeError):
#             return False
#     if "text" in rule:
#         return isinstance(value, str)
#     if "binary" in rule:
#         return str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']
#     if "date" in rule:
#         return isinstance(value, datetime)
#     return False

# # --- Core Comparison Logic with All Checks ---
# def compare_excel_files(input_file, output_file):
#     results = {}
#     try:
#         # Load two versions of the template: one for formats, one for raw data values
#         input_wb_format = openpyxl.load_workbook(input_file)
#         input_wb_data = openpyxl.load_workbook(input_file, data_only=True)
#         # Load the output file to be checked
#         output_wb = openpyxl.load_workbook(output_file)
        
#         common_sheets = sorted(list(set(input_wb_format.sheetnames).intersection(set(output_wb.sheetnames))))

#         for sheet_name in common_sheets:
#             results[sheet_name] = {"correct_cells": [], "discrepancies": []}
#             input_ws = input_wb_format[sheet_name]
#             input_ws_data = input_wb_data[sheet_name]
#             output_ws = output_wb[sheet_name]
            
#             type_rule_row_index = detect_data_type_row(input_ws_data)
#             if type_rule_row_index is None:
#                 st.warning(f"Could not detect a data type definition row in template sheet '{sheet_name}'. Skipping.")
#                 continue

#             output_df = pd.read_excel(output_file, sheet_name=sheet_name, header=type_rule_row_index - 1)
#             num_data_cols = output_df.shape[1]

#             # --- 1. ACCURACY & FORMATTING TEST (Headers) ---
#             for row in range(1, type_rule_row_index + 1):
#                 for col in range(1, num_data_cols + 1):
#                     template_cell_fmt = input_ws.cell(row=row, column=col)
#                     template_cell_val = input_ws_data.cell(row=row, column=col)
#                     output_cell = output_ws.cell(row=row, column=col)
#                     error_base = {"Cell": output_cell.coordinate, "Column": f"Header Col_{col}"}
                    
#                     template_type = check_specific_datatype(template_cell_val.value)
#                     output_type = check_specific_datatype(output_cell.value)
                    
#                     is_correct = True
#                     # Accuracy Checks
#                     if template_type != output_type:
#                         results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Header Datatype Mismatch", "Category": "Accuracy", "Template_Value": f"{template_cell_val.value} (Type: {template_type})", "Output_Value": f"{output_cell.value} (Type: {output_type})"})
#                         is_correct = False
#                     elif template_cell_val.value != output_cell.value:
#                         results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Header Value Mismatch", "Category": "Accuracy", "Template_Value": template_cell_val.value, "Output_Value": output_cell.value})
#                         is_correct = False
                    
#                     # Formatting Checks
#                     if (template_cell_fmt.font.name != output_cell.font.name or int(template_cell_fmt.font.size) != int(output_cell.font.size) or template_cell_fmt.font.bold != output_cell.font.bold):
#                          results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Header Font Mismatch", "Category": "Formatting", "Template_Value": f"{template_cell_fmt.font.name}, {int(template_cell_fmt.font.size)}pt, Bold:{template_cell_fmt.font.bold}", "Output_Value": f"{output_cell.font.name}, {int(output_cell.font.size)}pt, Bold:{output_cell.font.bold}"})
#                          is_correct = False

#                     if is_correct:
#                         results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": template_cell_val.value, "Output_Value": output_cell.value})

#             # --- 2. DATATYPE & FORMATTING TEST (Data) ---
#             rules = {col: input_ws_data.cell(row=type_rule_row_index, column=col).value for col in range(1, num_data_cols + 1)}
#             format_template_cells = {col: input_ws.cell(row=type_rule_row_index, column=col) for col in range(1, num_data_cols + 1)}

#             for r_idx, data_row in enumerate(output_df.itertuples(index=False)):
#                 for c_idx, value in enumerate(data_row):
#                     col = c_idx + 1
#                     row = type_rule_row_index + 1 + r_idx
#                     rule_string = rules.get(col)
#                     format_template_cell = format_template_cells.get(col)
#                     cell_coord = f"{openpyxl.utils.get_column_letter(col)}{row}"
#                     column_name = output_df.columns[c_idx] if c_idx < len(output_df.columns) else 'N/A'
#                     output_cell = output_ws.cell(row=row, column=col)
#                     error_base = {"Cell": cell_coord, "Column": column_name}
                    
#                     is_correct = True
#                     # Datatype Check
#                     if not validate_value_against_type(rule_string, value):
#                         results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Datatype Mismatch", "Category": "Datatype", "Template_Value": f"Rule: '{rule_string}'", "Output_Value": f"Value: '{value}'"})
#                         is_correct = False
                    
#                     # Formatting Checks
#                     if format_template_cell:
#                         if (format_template_cell.font.name != output_cell.font.name or int(format_template_cell.font.size) != int(output_cell.font.size) or format_template_cell.font.bold != output_cell.font.bold):
#                             results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Font Mismatch", "Category": "Formatting", "Template_Value": f"{format_template_cell.font.name}, {int(format_template_cell.font.size)}pt, Bold:{format_template_cell.font.bold}", "Output_Value": f"{output_cell.font.name}, {int(output_cell.font.size)}pt, Bold:{output_cell.font.bold}"})
#                             is_correct = False

#                     if is_correct:
#                         results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": f"Rule: '{rule_string}'", "Output_Value": value})

#     except Exception as e:
#         st.error(f"A critical error occurred: {e}")
#     return results

# # --- New Report Generation to Match Your Format ---
# def generate_excel_report(results, epic_number):
#     output_stream = io.BytesIO()
#     writer = pd.ExcelWriter(output_stream, engine='xlsxwriter')
#     workbook = writer.book

#     all_results_list = []
    
#     # Process all discrepancies (Wrong Fields)
#     for sheet_name, sheet_results in results.items():
#         for error in sheet_results.get("discrepancies", []):
#             all_results_list.append({
#                 "EPIC #": epic_number,
#                 "SHEET": sheet_name,
#                 "CELL": error.get("Cell", "N/A"),
#                 "FIELD": error.get("Column", "N/A"),
#                 "EXPECTED VALUE": str(error.get("Template_Value", "")),
#                 "TEST VALUE": str(error.get("Output_Value", "")),
#                 "RIGHT/WRONG": "WRONG",
#                 "Correct fields": "",
#                 "Wrong fields": error.get("Reason", "Unknown Error")
#             })

#     # Process all correct cells (Correct Fields)
#     for sheet_name, sheet_results in results.items():
#         for correct in sheet_results.get("correct_cells", []):
#              all_results_list.append({
#                 "EPIC #": epic_number,
#                 "SHEET": sheet_name,
#                 "CELL": correct.get("Cell", "N/A"),
#                 "FIELD": correct.get("Column", "N/A"),
#                 "EXPECTED VALUE": str(correct.get("Template_Value", "")),
#                 "TEST VALUE": str(correct.get("Output_Value", "")),
#                 "RIGHT/WRONG": "RIGHT",
#                 "Correct fields": "OK",
#                 "Wrong fields": ""
#             })

#     if not all_results_list:
#         detailed_df = pd.DataFrame(columns=["EPIC #", "SHEET", "CELL", "FIELD", "EXPECTED VALUE", "TEST VALUE", "RIGHT/WRONG", "Correct fields", "Wrong fields"])
#     else:
#         detailed_df = pd.DataFrame(all_results_list)

#     # --- Sheet 2: Detailed Test Results ---
#     detailed_df.to_excel(writer, sheet_name="Detailed Test Results", index=False)
#     worksheet = writer.sheets["Detailed Test Results"]
#     worksheet.set_column('A:I', 22)

#     # --- Sheet 1: QA Dashboard ---
#     correct_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "RIGHT"])
#     wrong_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "WRONG"])
#     total_count = correct_count + wrong_count
#     performance_score = (correct_count / total_count * 100) if total_count > 0 else 100

#     dashboard_sheet = workbook.add_worksheet("QA Dashboard")
#     title_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter'})
#     kpi_format = workbook.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter'})
#     kpi_label_format = workbook.add_format({'font_size': 12, 'align': 'center', 'font_color': '#595959'})

#     dashboard_sheet.merge_range('B2:F3', 'Quality Assurance Dashboard', title_format)
#     dashboard_sheet.merge_range('B5:C7', correct_count, kpi_format); dashboard_sheet.merge_range('B8:C8', 'Correct Count', kpi_label_format)
#     dashboard_sheet.merge_range('E5:F7', wrong_count, kpi_format); dashboard_sheet.merge_range('E8:F8', 'Wrong Count', kpi_label_format)
#     dashboard_sheet.merge_range('B10:F12', f"{performance_score:.1f}%", kpi_format); dashboard_sheet.merge_range('B13:F13', 'Overall Performance Score', kpi_label_format)
#     dashboard_sheet.set_column('B:F', 20)

#     writer.close()
#     output_stream.seek(0)
#     return output_stream

# # --- Final UI with EPIC # Input ---
# st.set_page_config(page_title="Excel Validator", layout="wide")
# st.title("Excel Validator Tool")

# # Initialize session state variables
# if 'ran_comparison' not in st.session_state:
#     st.session_state.ran_comparison = False
# if 'results' not in st.session_state:
#     st.session_state.results = {}
# if 'epic_number' not in st.session_state:
#     st.session_state.epic_number = ""

# epic_number = st.text_input("Enter EPIC #")
# input_file = st.file_uploader("Upload Input (Master Template) Excel", type=['xlsx'])
# output_file = st.file_uploader("Upload Output (To Test) Excel", type=['xlsx'])

# if epic_number and input_file and output_file:
#     if st.button("Run Full Validation", type="primary"):
#         st.session_state.epic_number = epic_number
#         with st.spinner("Performing full validation..."):
#             st.session_state.results = compare_excel_files(input_file, output_file)
#         st.session_state.ran_comparison = True

# if st.session_state.ran_comparison:
#     results = st.session_state.get('results', {})
#     correct_count = sum(len(sheet.get('correct_cells', [])) for sheet in results.values())
#     wrong_count = sum(len(sheet.get('discrepancies', [])) for sheet in results.values())
#     total_count = correct_count + wrong_count
#     performance_score = (correct_count / total_count * 100) if total_count > 0 else 100

#     st.header("On-Screen Validation Summary")
#     col1, col2, col3 = st.columns(3)
#     col1.metric(label="📊 Performance Score", value=f"{performance_score:.2f}%")
#     col2.metric(label="✅ Correct Fields", value=correct_count)
#     col3.metric(label="❌ Wrong Fields", value=wrong_count, delta_color="inverse")
            
#     st.markdown("---")
#     st.download_button(
#         label="📄 Download Full Business Report (Excel)", 
#         data=generate_excel_report(results, st.session_state.epic_number), 
#         file_name=f"EPIC_{st.session_state.epic_number}_Validation_Report_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx", 
#         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#     )
# import streamlit as st
# import pandas as pd
# import openpyxl
# import io
# from datetime import datetime

# # --- Helper functions ---

# def check_specific_datatype(value):
#     """Determines the specific data type of a given value for reporting."""
#     if value is None or str(value).strip() == '':
#         return "Empty"
#     if isinstance(value, (int, float)): 
#         return "Numeric"
#     if isinstance(value, str):
#         # Check for binary first, as it's a subset of string
#         if str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']: 
#             return "Binary(Yes/No)"
#         return "Text"
#     if isinstance(value, datetime): 
#         return "Date"
#     # For any other type, try to convert to string for reporting
#     try:
#         float(value)
#         return "Numeric"
#     except (ValueError, TypeError):
#         return "Other"
    
# def validate_value_against_type(rule, value):
#     """Checks if a value conforms to a data type rule string. Empty is always valid."""
#     if value is None or str(value).strip() == '': 
#         return True
#     if not isinstance(rule, str): 
#         return False # A missing rule in the template means validation fails for non-empty cells.
    
#     rule = rule.lower().strip()
    
#     if "numeric" in rule or "decimal" in rule:
#         if isinstance(value, (int, float)): return True
#         try:
#             float(str(value)); return True
#         except (ValueError, TypeError):
#             return False
#     if "text" in rule:
#         return isinstance(value, str)
#     if "binary" in rule:
#         return str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']
#     if "date" in rule:
#         return isinstance(value, datetime)
#     return False

# # --- UPDATED: Core Logic with ALL FORMAT CHECKS REMOVED ---
# def compare_excel_files(input_file, output_file):
#     results = {}
#     try:
#         # We only need the data, so load both workbooks with data_only=True
#         input_wb = openpyxl.load_workbook(input_file, data_only=True)
#         output_wb = openpyxl.load_workbook(output_file, data_only=True)
#     except Exception as e:
#         st.error(f"Fatal Error: Could not open or read the Excel files. Please check if they are valid. Details: {e}")
#         return {}

#     common_sheets = sorted(list(set(input_wb.sheetnames).intersection(set(output_wb.sheetnames))))

#     for sheet_name in common_sheets:
#         try:
#             results[sheet_name] = {"correct_cells": [], "discrepancies": []}
#             input_ws = input_wb[sheet_name]
#             output_ws = output_wb[sheet_name]
            
#             num_data_cols = input_ws.max_column
#             headers = {c: input_ws.cell(row=3, column=c).value for c in range(1, num_data_cols + 1)}
#             rules = {col: input_ws.cell(row=4, column=col).value for col in range(1, num_data_cols + 1)}

#             # --- 1. HEADER VALUE CHECK (Row 3) - No format check ---
#             for col in range(1, num_data_cols + 1):
#                 template_cell = input_ws.cell(row=3, column=col)
#                 output_cell = output_ws.cell(row=3, column=col)
#                 column_name = headers.get(col)
#                 error_base = {"Cell": output_cell.coordinate, "Column": column_name}
                
#                 if template_cell.value != output_cell.value:
#                     results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Header Value Mismatch", "Template_Value": template_cell.value, "Output_Value": output_cell.value})
#                 else:
#                     results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": template_cell.value, "Output_Value": output_cell.value})

#             # --- 2. DATA TYPE CHECK (Row 4 onwards) - No format check ---
#             for row_idx in range(4, output_ws.max_row + 1):
#                 row_cells = [output_ws.cell(row=row_idx, column=c).value for c in range(1, num_data_cols + 1)]
#                 if all(cell is None or str(cell).strip() == "" for cell in row_cells):
#                     continue

#                 for col_idx in range(1, num_data_cols + 1):
#                     rule_string = rules.get(col_idx)
#                     output_cell = output_ws.cell(row=row_idx, column=col_idx)
#                     column_name = headers.get(col_idx, f"Column_{col_idx}")
#                     error_base = {"Cell": output_cell.coordinate, "Column": column_name}
                    
#                     if not validate_value_against_type(rule_string, output_cell.value):
#                         detected_type = check_specific_datatype(output_cell.value)
#                         results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Datatype Mismatch", "Template_Value": rule_string, "Output_Value": detected_type})
#                     else:
#                         results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": rule_string, "Output_Value": output_cell.value})
        
#         except Exception as e:
#             st.warning(f"Could not process sheet '{sheet_name}'. The following error occurred: {e}")
#             if sheet_name in results:
#                 del results[sheet_name]
#             continue
    
#     return results

# # --- Report Generation (no changes here) ---
# def generate_excel_report(results, epic_number):
#     output_stream = io.BytesIO()
#     writer = pd.ExcelWriter(output_stream, engine='xlsxwriter')
#     workbook = writer.book

#     all_results_list = []
    
#     for sheet_name, sheet_results in results.items():
#         if sheet_results:
#             for error in sheet_results.get("discrepancies", []):
#                 all_results_list.append({
#                     "EPIC #": epic_number, "SHEET": sheet_name, "CELL": error.get("Cell", "N/A"),
#                     "FIELD": error.get("Column", "N/A"), "EXPECTED VALUE": str(error.get("Template_Value", "")),
#                     "TEST VALUE": str(error.get("Output_Value", "")), "RIGHT/WRONG": "WRONG",
#                     "Correct fields": "", "Wrong fields": error.get("Reason", "Unknown Error")
#                 })
#             for correct in sheet_results.get("correct_cells", []):
#                  all_results_list.append({
#                     "EPIC #": epic_number, "SHEET": sheet_name, "CELL": correct.get("Cell", "N/A"),
#                     "FIELD": correct.get("Column", "N/A"), "EXPECTED VALUE": str(correct.get("Template_Value", "")),
#                     "TEST VALUE": str(correct.get("Output_Value", "")), "RIGHT/WRONG": "RIGHT",
#                     "Correct fields": "OK", "Wrong fields": ""
#                 })

#     columns = ["EPIC #", "SHEET", "CELL", "FIELD", "EXPECTED VALUE", "TEST VALUE", "RIGHT/WRONG", "Correct fields", "Wrong fields"]
#     if not all_results_list:
#         detailed_df = pd.DataFrame(columns=columns)
#     else:
#         detailed_df = pd.DataFrame(all_results_list, columns=columns)

#     correct_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "RIGHT"])
#     wrong_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "WRONG"])
#     total_count = correct_count + wrong_count
#     performance_score = (correct_count / total_count * 100) if total_count > 0 else 100

#     dashboard_sheet = workbook.add_worksheet("QA Dashboard")
#     title_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter'})
#     kpi_format = workbook.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter'})
#     kpi_label_format = workbook.add_format({'font_size': 12, 'align': 'center', 'font_color': '#595959'})
#     dashboard_sheet.merge_range('B2:F3', 'Quality Assurance Dashboard', title_format)
#     dashboard_sheet.merge_range('B5:C7', correct_count, kpi_format); dashboard_sheet.merge_range('B8:C8', 'Correct Count', kpi_label_format)
#     dashboard_sheet.merge_range('E5:F7', wrong_count, kpi_format); dashboard_sheet.merge_range('E8:F8', 'Wrong Count', kpi_label_format)
#     dashboard_sheet.merge_range('B10:F12', f"{performance_score:.1f}%", kpi_format); dashboard_sheet.merge_range('B13:F13', 'Overall Performance Score', kpi_label_format)
#     dashboard_sheet.set_column('B:F', 20)

#     detailed_df.to_excel(writer, sheet_name="Detailed Test Results", index=False, startrow=1, header=False)
#     worksheet = writer.sheets["Detailed Test Results"]

#     header_format_yellow = workbook.add_format({'bold': True, 'font_color': '#000000', 'bg_color': '#FFFF00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
#     header_format_red = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#FF0000', 'align': 'center', 'valign': 'vcenter', 'border': 1})
#     header_format_green = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#70AD47', 'align': 'center', 'valign': 'vcenter', 'border': 1})

#     for col_num, col_name in enumerate(detailed_df.columns):
#         if col_num < 3:
#             worksheet.write(0, col_num, col_name, header_format_yellow)
#         elif col_num < 6:
#             worksheet.write(0, col_num, col_name, header_format_red)
#         else:
#             worksheet.write(0, col_num, col_name, header_format_green)
    
#     worksheet.set_column('A:I', 22)

#     writer.close()
#     output_stream.seek(0)
#     return output_stream

# # --- UI (no changes here) ---
# st.set_page_config(page_title="Excel Validator Tool", layout="wide")
# st.title("Excel Validator Tool")

# if 'ran_comparison' not in st.session_state:
#     st.session_state.ran_comparison = False
# if 'results' not in st.session_state:
#     st.session_state.results = {}
# if 'epic_number' not in st.session_state:
#     st.session_state.epic_number = ""

# epic_number = st.text_input("Enter EPIC #")
# input_file = st.file_uploader("Upload Input (Template) Excel", type=['xlsx'])
# output_file = st.file_uploader("Upload Output (File to Test) Excel", type=['xlsx'])

# if epic_number and input_file and output_file:
#     if st.button("Run Full Validation", type="primary"):
#         st.session_state.epic_number = epic_number
#         with st.spinner("Performing validation..."):
#             st.session_state.results = compare_excel_files(input_file, output_file)
#         st.session_state.ran_comparison = True

# if st.session_state.ran_comparison:
#     results = st.session_state.get('results', {})
#     if results:
#         correct_count = sum(len(sheet.get('correct_cells', [])) for sheet in results.values() if sheet)
#         wrong_count = sum(len(sheet.get('discrepancies', [])) for sheet in results.values() if sheet)
#         total_count = correct_count + wrong_count
        
#         performance_score = (correct_count / total_count * 100) if total_count > 0 else 100

#         st.header("On-Screen Validation Summary")
#         col1, col2, col3 = st.columns(3)
#         col1.metric(label="📊 Performance Score", value=f"{performance_score:.2f}%")
#         col2.metric(label="✅ Correct Fields", value=correct_count)
#         col3.metric(label="❌ Wrong Fields", value=wrong_count, delta_color="inverse")
                
#         st.markdown("---")
#         st.download_button(
#             label="📄 Download Full Business Report (Excel)", 
#             data=generate_excel_report(results, st.session_state.epic_number), 
#             file_name=f"EPIC_{st.session_state.epic_number}_Validation_Report_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx", 
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )

# import streamlit as st
# import pandas as pd
# import openpyxl
# import io
# from datetime import datetime

# # --- Helper functions (no changes here) ---

# def check_specific_datatype(value):
#     """Determines the specific data type of a given value for reporting."""
#     if value is None or str(value).strip() == '':
#         return "Empty"
#     if isinstance(value, (int, float)): 
#         return "Numeric"
#     if isinstance(value, str):
#         if str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']: 
#             return "Binary(Yes/No)"
#         return "Text"
#     if isinstance(value, datetime): 
#         return "Date"
#     try:
#         float(value)
#         return "Numeric"
#     except (ValueError, TypeError):
#         return "Other"
    
# def validate_value_against_type(rule, value):
#     """Checks if a value conforms to a data type rule string. Empty is always valid."""
#     if value is None or str(value).strip() == '': 
#         return True
#     if not isinstance(rule, str): 
#         return False
    
#     rule = rule.lower().strip()
    
#     if "numeric" in rule or "decimal" in rule:
#         if isinstance(value, (int, float)): return True
#         try:
#             float(str(value)); return True
#         except (ValueError, TypeError):
#             return False
#     if "text" in rule:
#         return isinstance(value, str)
#     if "binary" in rule:
#         return str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']
#     if "date" in rule:
#         return isinstance(value, datetime)
#     return False

# # --- Core Comparison Logic (no changes here) ---
# def compare_excel_files(input_file, output_file):
#     results = {}
#     try:
#         input_wb = openpyxl.load_workbook(input_file, data_only=True)
#         output_wb = openpyxl.load_workbook(output_file, data_only=True)
#     except Exception as e:
#         st.error(f"Fatal Error: Could not open or read the Excel files. Please check if they are valid. Details: {e}")
#         return {}

#     common_sheets = sorted(list(set(input_wb.sheetnames).intersection(set(output_wb.sheetnames))))

#     for sheet_name in common_sheets:
#         try:
#             results[sheet_name] = {"correct_cells": [], "discrepancies": []}
#             input_ws = input_wb[sheet_name]
#             output_ws = output_wb[sheet_name]
            
#             num_data_cols = input_ws.max_column
#             headers = {c: input_ws.cell(row=3, column=c).value for c in range(1, num_data_cols + 1)}
#             rules = {col: input_ws.cell(row=4, column=col).value for col in range(1, num_data_cols + 1)}

#             # 1. HEADER VALUE CHECK (Row 3)
#             for col in range(1, num_data_cols + 1):
#                 template_cell = input_ws.cell(row=3, column=col)
#                 output_cell = output_ws.cell(row=3, column=col)
#                 column_name = headers.get(col)
#                 error_base = {"Cell": output_cell.coordinate, "Column": column_name}
                
#                 if template_cell.value != output_cell.value:
#                     results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Header Value Mismatch", "Template_Value": template_cell.value, "Output_Value": output_cell.value})
#                 else:
#                     results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": template_cell.value, "Output_Value": output_cell.value})

#             # 2. DATA TYPE CHECK (Row 4 onwards)
#             for row_idx in range(4, output_ws.max_row + 1):
#                 row_cells = [output_ws.cell(row=row_idx, column=c).value for c in range(1, num_data_cols + 1)]
#                 if all(cell is None or str(cell).strip() == "" for cell in row_cells):
#                     continue

#                 for col_idx in range(1, num_data_cols + 1):
#                     rule_string = rules.get(col_idx)
#                     output_cell = output_ws.cell(row=row_idx, column=col_idx)
#                     column_name = headers.get(col_idx, f"Column_{col_idx}")
#                     error_base = {"Cell": output_cell.coordinate, "Column": column_name}
                    
#                     if not validate_value_against_type(rule_string, output_cell.value):
#                         detected_type = check_specific_datatype(output_cell.value)
#                         results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Datatype Mismatch", "Template_Value": rule_string, "Output_Value": detected_type})
#                     else:
#                         results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": rule_string, "Output_Value": output_cell.value})
        
#         except Exception as e:
#             st.warning(f"Could not process sheet '{sheet_name}'. The following error occurred: {e}")
#             if sheet_name in results:
#                 del results[sheet_name]
#             continue
    
#     return results

# # --- UPDATED: Report Generation with the new "Accuracy Score" ---
# def generate_excel_report(results, epic_number):
#     output_stream = io.BytesIO()
#     writer = pd.ExcelWriter(output_stream, engine='xlsxwriter')
#     workbook = writer.book

#     all_results_list = []
    
#     for sheet_name, sheet_results in results.items():
#         if sheet_results:
#             for error in sheet_results.get("discrepancies", []):
#                 all_results_list.append({
#                     "EPIC #": epic_number, "SHEET": sheet_name, "CELL": error.get("Cell", "N/A"),
#                     "FIELD": error.get("Column", "N/A"), "EXPECTED VALUE": str(error.get("Template_Value", "")),
#                     "TEST VALUE": str(error.get("Output_Value", "")), "RIGHT/WRONG": "WRONG",
#                     "Correct fields": "", "Wrong fields": error.get("Reason", "Unknown Error")
#                 })
#             for correct in sheet_results.get("correct_cells", []):
#                  all_results_list.append({
#                     "EPIC #": epic_number, "SHEET": sheet_name, "CELL": correct.get("Cell", "N/A"),
#                     "FIELD": correct.get("Column", "N/A"), "EXPECTED VALUE": str(correct.get("Template_Value", "")),
#                     "TEST VALUE": str(correct.get("Output_Value", "")), "RIGHT/WRONG": "RIGHT",
#                     "Correct fields": "OK", "Wrong fields": ""
#                 })

#     columns = ["EPIC #", "SHEET", "CELL", "FIELD", "EXPECTED VALUE", "TEST VALUE", "RIGHT/WRONG", "Correct fields", "Wrong fields"]
#     if not all_results_list:
#         detailed_df = pd.DataFrame(columns=columns)
#     else:
#         detailed_df = pd.DataFrame(all_results_list, columns=columns)

#     # --- THIS IS THE UPDATED SECTION ---
#     # Overall KPI calculations (includes headers)
#     correct_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "RIGHT"])
#     wrong_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "WRONG"])
#     total_count = correct_count + wrong_count
#     performance_score = (correct_count / total_count * 100) if total_count > 0 else 100

#     # New KPI calculations (for data cells ONLY)
#     # We extract the row number from the 'CELL' coordinate string (e.g., 'A4' -> 4)
#     # and filter for rows >= 4.
#     data_rows_df = detailed_df[pd.to_numeric(detailed_df['CELL'].str.extract(r'(\d+)')[0], errors='coerce') >= 4]
    
#     correct_data_count = len(data_rows_df[data_rows_df["RIGHT/WRONG"] == "RIGHT"])
#     total_data_count = len(data_rows_df)
#     accuracy_score = (correct_data_count / total_data_count * 100) if total_data_count > 0 else 100

#     # Create Sheet 1: QA Dashboard with the new layout
#     dashboard_sheet = workbook.add_worksheet("QA Dashboard")
#     title_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter'})
#     kpi_format = workbook.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter'})
#     kpi_label_format = workbook.add_format({'font_size': 12, 'align': 'center', 'font_color': '#595959'})

#     dashboard_sheet.merge_range('B2:G3', 'Quality Assurance Dashboard', title_format)

#     # Row 1 of KPIs
#     dashboard_sheet.merge_range('B5:D7', f"{performance_score:.1f}%", kpi_format)
#     dashboard_sheet.merge_range('B8:D8', 'Overall Performance Score (All Checks)', kpi_label_format)
#     dashboard_sheet.merge_range('E5:G7', f"{accuracy_score:.1f}%", kpi_format)
#     dashboard_sheet.merge_range('E8:G8', 'Overall Accuracy Score (Data Only)', kpi_label_format)

#     # Row 2 of KPIs
#     dashboard_sheet.merge_range('B10:D12', correct_count, kpi_format)
#     dashboard_sheet.merge_range('B13:D13', 'Total Correct Fields', kpi_label_format)
#     dashboard_sheet.merge_range('E10:G12', wrong_count, kpi_format)
#     dashboard_sheet.merge_range('E13:G13', 'Total Wrong Fields', kpi_label_format)
    
#     dashboard_sheet.set_column('B:G', 22)
#     # --- END OF UPDATED SECTION ---
    
#     # Create Sheet 2: Detailed Test Results
#     detailed_df.to_excel(writer, sheet_name="Detailed Test Results", index=False, startrow=1, header=False)
#     worksheet = writer.sheets["Detailed Test Results"]

#     header_format_yellow = workbook.add_format({'bold': True, 'font_color': '#000000', 'bg_color': '#FFFF00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
#     header_format_red = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#FF0000', 'align': 'center', 'valign': 'vcenter', 'border': 1})
#     header_format_green = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#70AD47', 'align': 'center', 'valign': 'vcenter', 'border': 1})

#     for col_num, col_name in enumerate(detailed_df.columns):
#         if col_num < 3:
#             worksheet.write(0, col_num, col_name, header_format_yellow)
#         elif col_num < 6:
#             worksheet.write(0, col_num, col_name, header_format_red)
#         else:
#             worksheet.write(0, col_num, col_name, header_format_green)
    
#     worksheet.set_column('A:I', 22)

#     writer.close()
#     output_stream.seek(0)
#     return output_stream

# # --- UI (no changes here) ---
# st.set_page_config(page_title="Excel Validator Tool", layout="wide")
# st.title("Excel Validator Tool")

# if 'ran_comparison' not in st.session_state:
#     st.session_state.ran_comparison = False
# if 'results' not in st.session_state:
#     st.session_state.results = {}
# if 'epic_number' not in st.session_state:
#     st.session_state.epic_number = ""

# epic_number = st.text_input("Enter EPIC #")
# input_file = st.file_uploader("Upload Input (Template) Excel", type=['xlsx'])
# output_file = st.file_uploader("Upload Output (File to Test) Excel", type=['xlsx'])

# if epic_number and input_file and output_file:
#     if st.button("Run Full Validation", type="primary"):
#         st.session_state.epic_number = epic_number
#         with st.spinner("Performing validation..."):
#             st.session_state.results = compare_excel_files(input_file, output_file)
#         st.session_state.ran_comparison = True

# if st.session_state.ran_comparison:
#     results = st.session_state.get('results', {})
#     if results:
#         correct_count = sum(len(sheet.get('correct_cells', [])) for sheet in results.values() if sheet)
#         wrong_count = sum(len(sheet.get('discrepancies', [])) for sheet in results.values() if sheet)
#         total_count = correct_count + wrong_count
        
#         performance_score = (correct_count / total_count * 100) if total_count > 0 else 100

#         st.header("On-Screen Validation Summary")
#         col1, col2, col3 = st.columns(3)
#         col1.metric(label="📊 Performance Score", value=f"{performance_score:.2f}%")
#         col2.metric(label="✅ Correct Fields", value=correct_count)
#         col3.metric(label="❌ Wrong Fields", value=wrong_count, delta_color="inverse")
                
#         st.markdown("---")
#         st.download_button(
#             label="📄 Download Full Business Report (Excel)", 
#             data=generate_excel_report(results, st.session_state.epic_number), 
#             file_name=f"EPIC_{st.session_state.epic_number}_Validation_Report_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx", 
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )

import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

# --- Helper functions (no changes here) ---

def check_specific_datatype(value):
    """Determines the specific data type of a given value for reporting."""
    if value is None or str(value).strip() == '':
        return "Empty"
    if isinstance(value, (int, float)): 
        return "Numeric"
    if isinstance(value, str):
        if str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']: 
            return "Binary(Yes/No)"
        return "Text"
    if isinstance(value, datetime): 
        return "Date"
    try:
        float(value)
        return "Numeric"
    except (ValueError, TypeError):
        return "Other"
    
def validate_value_against_type(rule, value):
    """Checks if a value conforms to a data type rule string."""
    if value is None or str(value).strip() == '': 
        return True
    if not isinstance(rule, str): 
        return False
    
    rule = rule.lower().strip()
    
    if "numeric" in rule or "decimal" in rule:
        if isinstance(value, (int, float)): return True
        try:
            float(str(value)); return True
        except (ValueError, TypeError):
            return False
    if "text" in rule:
        return isinstance(value, str)
    if "binary" in rule:
        return str(value).lower() in ['yes', 'no', 'y', 'n', 'true', 'false']
    if "date" in rule:
        return isinstance(value, datetime)
    return False

# --- UPDATED: Core Logic to Skip Empty Template/Output Cells ---
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
            headers = {c: input_ws.cell(row=3, column=c).value for c in range(1, num_data_cols + 1)}
            rules = {col: input_ws.cell(row=4, column=col).value for col in range(1, num_data_cols + 1)}

            # 1. HEADER VALUE CHECK (Row 3)
            for col in range(1, num_data_cols + 1):
                template_cell = input_ws.cell(row=3, column=col)
                
                # --- NEW: Skip if the template header cell is empty ---
                if template_cell.value is None or str(template_cell.value).strip() == '':
                    continue
                
                output_cell = output_ws.cell(row=3, column=col)
                column_name = headers.get(col)
                error_base = {"Cell": output_cell.coordinate, "Column": column_name}
                
                if template_cell.value != output_cell.value:
                    results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Header Value Mismatch", "Template_Value": template_cell.value, "Output_Value": output_cell.value})
                else:
                    results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": template_cell.value, "Output_Value": output_cell.value})

            # 2. DATA TYPE CHECK (Row 4 onwards)
            for row_idx in range(4, output_ws.max_row + 1):
                row_cells = [output_ws.cell(row=row_idx, column=c).value for c in range(1, num_data_cols + 1)]
                if all(cell is None or str(cell).strip() == "" for cell in row_cells):
                    continue

                for col_idx in range(1, num_data_cols + 1):
                    rule_string = rules.get(col_idx)
                    output_cell = output_ws.cell(row=row_idx, column=col_idx)
                    
                    # --- NEW: Skip if the template rule is empty OR the output data is empty ---
                    if (rule_string is None or str(rule_string).strip() == '') or \
                       (output_cell.value is None or str(output_cell.value).strip() == ''):
                        continue

                    column_name = headers.get(col_idx, f"Column_{col_idx}")
                    error_base = {"Cell": output_cell.coordinate, "Column": column_name}
                    
                    if not validate_value_against_type(rule_string, output_cell.value):
                        detected_type = check_specific_datatype(output_cell.value)
                        results[sheet_name]["discrepancies"].append({**error_base, "Reason": "Datatype Mismatch", "Template_Value": rule_string, "Output_Value": detected_type})
                    else:
                        results[sheet_name]["correct_cells"].append({**error_base, "Template_Value": rule_string, "Output_Value": output_cell.value})
        
        except Exception as e:
            st.warning(f"Could not process sheet '{sheet_name}'. The following error occurred: {e}")
            if sheet_name in results:
                del results[sheet_name]
            continue
    
    return results

# --- Report Generation (no changes here) ---
def generate_excel_report(results, epic_number):
    output_stream = io.BytesIO()
    writer = pd.ExcelWriter(output_stream, engine='xlsxwriter')
    workbook = writer.book

    all_results_list = []
    
    for sheet_name, sheet_results in results.items():
        if sheet_results:
            for error in sheet_results.get("discrepancies", []):
                all_results_list.append({
                    "EPIC #": epic_number, "SHEET": sheet_name, "CELL": error.get("Cell", "N/A"),
                    "FIELD": error.get("Column", "N/A"), "EXPECTED VALUE": str(error.get("Template_Value", "")),
                    "TEST VALUE": str(error.get("Output_Value", "")), "RIGHT/WRONG": "WRONG",
                    "Correct fields": "", "Wrong fields": error.get("Reason", "Unknown Error")
                })
            for correct in sheet_results.get("correct_cells", []):
                 all_results_list.append({
                    "EPIC #": epic_number, "SHEET": sheet_name, "CELL": correct.get("Cell", "N/A"),
                    "FIELD": correct.get("Column", "N/A"), "EXPECTED VALUE": str(correct.get("Template_Value", "")),
                    "TEST VALUE": str(correct.get("Output_Value", "")), "RIGHT/WRONG": "RIGHT",
                    "Correct fields": "OK", "Wrong fields": ""
                })

    columns = ["EPIC #", "SHEET", "CELL", "FIELD", "EXPECTED VALUE", "TEST VALUE", "RIGHT/WRONG", "Correct fields", "Wrong fields"]
    if not all_results_list:
        detailed_df = pd.DataFrame(columns=columns)
    else:
        detailed_df = pd.DataFrame(all_results_list, columns=columns)

    # Overall KPI calculations (includes headers)
    correct_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "RIGHT"])
    wrong_count = len(detailed_df[detailed_df["RIGHT/WRONG"] == "WRONG"])
    total_count = correct_count + wrong_count
    performance_score = (correct_count / total_count * 100) if total_count > 0 else 100

    # Accuracy KPI for data cells only
    if not detailed_df.empty:
        data_rows_df = detailed_df[pd.to_numeric(detailed_df['CELL'].str.extract(r'(\d+)')[0], errors='coerce') >= 4]
        correct_data_count = len(data_rows_df[data_rows_df["RIGHT/WRONG"] == "RIGHT"])
        total_data_count = len(data_rows_df)
        accuracy_score = (correct_data_count / total_data_count * 100) if total_data_count > 0 else 100
    else:
        accuracy_score = 100.0

    # Create Sheet 1: QA Dashboard
    dashboard_sheet = workbook.add_worksheet("QA Dashboard")
    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter'})
    kpi_format = workbook.add_format({'bold': True, 'font_size': 28, 'align': 'center', 'valign': 'vcenter'})
    kpi_label_format = workbook.add_format({'font_size': 12, 'align': 'center', 'font_color': '#595959'})

    dashboard_sheet.merge_range('B2:G3', 'Quality Assurance Dashboard', title_format)
    dashboard_sheet.merge_range('B5:D7', f"{performance_score:.1f}%", kpi_format)
    dashboard_sheet.merge_range('B8:D8', 'Overall Performance Score (All Checks)', kpi_label_format)
    dashboard_sheet.merge_range('E5:G7', f"{accuracy_score:.1f}%", kpi_format)
    dashboard_sheet.merge_range('E8:G8', 'Overall Accuracy Score (Data Only)', kpi_label_format)
    dashboard_sheet.merge_range('B10:D12', correct_count, kpi_format)
    dashboard_sheet.merge_range('B13:D13', 'Total Correct Fields', kpi_label_format)
    dashboard_sheet.merge_range('E10:G12', wrong_count, kpi_format)
    dashboard_sheet.merge_range('E13:G13', 'Total Wrong Fields', kpi_label_format)
    
    dashboard_sheet.set_column('B:G', 22)
    
    # Create Sheet 2: Detailed Test Results
    detailed_df.to_excel(writer, sheet_name="Detailed Test Results", index=False, startrow=1, header=False)
    worksheet = writer.sheets["Detailed Test Results"]

    header_format_yellow = workbook.add_format({'bold': True, 'font_color': '#000000', 'bg_color': '#FFFF00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    header_format_red = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#FF0000', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    header_format_green = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#70AD47', 'align': 'center', 'valign': 'vcenter', 'border': 1})

    for col_num, col_name in enumerate(detailed_df.columns):
        if col_num < 3:
            worksheet.write(0, col_num, col_name, header_format_yellow)
        elif col_num < 6:
            worksheet.write(0, col_num, col_name, header_format_red)
        else:
            worksheet.write(0, col_num, col_name, header_format_green)
    
    worksheet.set_column('A:I', 22)

    writer.close()
    output_stream.seek(0)
    return output_stream

# --- UI (no changes here) ---
st.set_page_config(page_title="Excel Validator Tool", layout="wide")
st.title("Excel Validator Tool")

if 'ran_comparison' not in st.session_state:
    st.session_state.ran_comparison = False
if 'results' not in st.session_state:
    st.session_state.results = {}
if 'epic_number' not in st.session_state:
    st.session_state.epic_number = ""

epic_number = st.text_input("Enter EPIC #")
input_file = st.file_uploader("Upload Input (Template) Excel", type=['xlsx'])
output_file = st.file_uploader("Upload Output (File to Test) Excel", type=['xlsx'])

if epic_number and input_file and output_file:
    if st.button("Run Full Validation", type="primary"):
        st.session_state.epic_number = epic_number
        with st.spinner("Performing validation..."):
            st.session_state.results = compare_excel_files(input_file, output_file)
        st.session_state.ran_comparison = True

if st.session_state.ran_comparison:
    results = st.session_state.get('results', {})
    if results:
        correct_count = sum(len(sheet.get('correct_cells', [])) for sheet in results.values() if sheet)
        wrong_count = sum(len(sheet.get('discrepancies', [])) for sheet in results.values() if sheet)
        total_count = correct_count + wrong_count
        
        performance_score = (correct_count / total_count * 100) if total_count > 0 else 100

        st.header("On-Screen Validation Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric(label="📊 Performance Score", value=f"{performance_score:.2f}%")
        col2.metric(label="✅ Correct Fields", value=correct_count)
        col3.metric(label="❌ Wrong Fields", value=wrong_count, delta_color="inverse")
                
        st.markdown("---")
        st.download_button(
            label="📄 Download Full Business Report (Excel)", 
            data=generate_excel_report(results, st.session_state.epic_number), 
            file_name=f"EPIC_{st.session_state.epic_number}_Validation_Report_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

