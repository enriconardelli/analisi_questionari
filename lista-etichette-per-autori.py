# This script reads specific Excel files and prints the second row of a specified sheet
# to a text file named "lista-etichette-per-autori.txt".

import pandas as pd
import openpyxl as px
import sys
import os
from collections import Counter
from collections import defaultdict
import re
from docx import Document
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def stampa_su_file_ON(nome_file):
    output_file = open(nome_file, "w")
    # Redirect standard output to the file
    sys.stdout = output_file
    return output_file

def stampa_su_file_OFF(output_file):
    output_file.close()
    # Reset standard output back to console
    sys.stdout = sys.__stdout__

autori_file = {
    "ADZ": "questionario-insegnanti-campione-60-ADZ.xlsx",
    "CM": "questionario-insegnanti-campione-60-CM.xlsx",
    "EN": "questionario-insegnanti-campione-60-EN.xlsx",
    "GA": "questionario-insegnanti-campione-60-GA.xlsx",
    "LF": "questionario-insegnanti-campione-60-LF.xlsx",
    "ML": "questionario-insegnanti-campione-60-ML.xlsx"
}
NUMAUTORI = len(autori_file)

autori_filtered_file = {
    "ADZ": "ADZ_filtered.xlsx",
    "CM": "CM_filtered.xlsx",
    "EN": "EN_filtered.xlsx",
    "GA": "GA_filtered.xlsx",
    "LF": "LF_filtered.xlsx",
    "ML": "ML_filtered.xlsx"
}

autori_labeled_file = {
    "ADZ": "ADZ_labeled.xlsx",
    "CM": "CM_labeled.xlsx",
    "EN": "EN_labeled.xlsx",
    "GA": "GA_labeled.xlsx",
    "LF": "LF_labeled.xlsx",
    "ML": "ML_labeled.xlsx"
}

autori_combined_file = {
    "ADZ": "ADZ_combined.xlsx",
    "CM": "CM_combined.xlsx",
    "EN": "EN_combined.xlsx",
    "GA": "GA_combined.xlsx",
    "LF": "LF_combined.xlsx",
    "ML": "ML_combined.xlsx"
}

autori_combined_cleaned_file = {
    "ADZ": "ADZ_combined-cleaned.xlsx",
    "CM": "CM_combined-cleaned.xlsx",
    "EN": "EN_combined-cleaned.xlsx",
    "GA": "GA_combined-cleaned.xlsx",
    "LF": "LF_combined-cleaned.xlsx",
    "ML": "ML_combined-cleaned.xlsx"
}

autori_macroed_file = {
    "ADZ": "ADZ_macroed.xlsx",
    "CM": "CM_macroed.xlsx",
    "EN": "EN_macroed.xlsx",
    "GA": "GA_macroed.xlsx",
    "LF": "LF_macroed.xlsx",
    "ML": "ML_macroed.xlsx"
}

autori_macroed_merged_file = {
    "ADZ": "ADZ_macroed-merged.xlsx",
    "CM": "CM_macroed-merged.xlsx",
    "EN": "EN_macroed-merged.xlsx",
    "GA": "GA_macroed-merged.xlsx",
    "LF": "LF_macroed-merged.xlsx",
    "ML": "ML_macroed-merged.xlsx"
}

autori_formatted_file = {
    "ADZ": "ADZ_formatted.xlsx",
    "CM": "CM_formatted.xlsx",
    "EN": "EN_formatted.xlsx",
    "GA": "GA_formatted.xlsx",
    "LF": "LF_formatted.xlsx",
    "ML": "ML_formatted.xlsx"
}

# output_file = stampa_su_file_ON("lista-etichette-per-autori.txt")
# for name, file_path in autori_combined_file.items():
#     wb = px.load_workbook(file_path, data_only=True)
#     sheet = wb["Combined"]
#     etichette = list(sheet.iter_rows(min_row=2, max_row=2, min_col=1, values_only=True))[0] # tutti i nomi delle etichette
#     quante_etichette = len(etichette)
#     print(f"{name} has {quante_etichette} etichette.")  
#     # Find and print duplicated etichette with their counts
#     duplicates = [item for item, count in Counter(etichette).items() if count > 1]
#     if duplicates:
#         print(f"Duplicated etichette for {name}:")
#         for item in duplicates:
#             print(f"  {item}: {etichette.count(item)} times")

# CODICE PER FILTRARE LE COLONNE
# This code filters out columns starting with "ETICHETTE" or "ETICHETE" from the first three sheets of each workbook.
# for name, file_path in autori_file.items():
#     wb = px.load_workbook(file_path)
#     new_wb = px.Workbook()
#     # Remove the default sheet created by openpyxl
#     new_wb.remove(new_wb.active)
#     for sheet_name in wb.sheetnames[:3]:
#         ws = wb[sheet_name]
#         # Read all rows as values
#         rows = list(ws.values)
#         if len(rows) < 2:
#             continue  # Skip if not enough rows
#         # The second row contains the column names
#         col_names = list(rows[1])
#         # Find columns to keep (not starting with ETICHETTE or ETICHETE)
#         keep_indices = [i for i, name in enumerate(col_names) if not (isinstance(name, str) and (name.startswith("ETICHETTE") or name.startswith("ETICHETE")))]
#         # Prepare filtered rows (keep all rows, but only selected columns)
#         filtered_rows = []
#         for row in rows:
#             filtered_rows.append([row[i] for i in keep_indices])
#         # Create new sheet and write filtered rows
#         new_ws = new_wb.create_sheet(title=sheet_name)
#         for r in filtered_rows:
#             new_ws.append(r)
#     # Save the new workbook
#     new_file = f"{name}_filtered.xlsx"
#     new_wb.save(new_file)

# CODICE PER ETICHETTARE LE COLONNE CON LE DOMANDE
# For each Excel file in autori_filtered_file, create a new Excel file with concatenated column names
# for name, file_path in autori_filtered_file.items():
#     wb = px.load_workbook(file_path, data_only=True)
#     new_wb = px.Workbook()
#     new_wb.remove(new_wb.active)  # Remove default sheet
#     for idx, sheet_name in enumerate(wb.sheetnames[:3]):
#         ws = wb[sheet_name]
#         rows = list(ws.values)
#         if len(rows) < 2:
#             continue  # Skip if not enough rows
#         # The second row contains the column names
#         col_names = list(rows[1])
#         label = f"D{idx+1}"
#         new_col_names = [f"{label}_{col}" if col is not None else None for col in col_names]
#         # Prepare new rows: replace second row with new_col_names
#         new_rows = []
#         for i, row in enumerate(rows):
#             if i == 1:
#                 new_rows.append(new_col_names)
#             else:
#                 new_rows.append(list(row))
#         # Write to new sheet
#         new_ws = new_wb.create_sheet(title=sheet_name)
#         for row in new_rows:
#             new_ws.append(row)
#     new_file = f"{name}_labeled.xlsx"
#     new_wb.save(new_file)

# CODICE PER COMBINARE I 3 FOGLI IN UNO SOLO
# This code combines the first three sheets of each workbook into a new workbook,
# taking the first 130 columns from the first sheet, 94 from the second, and 102 from the third.
# for name, file_path in autori_labeled_file.items():
#     wb = px.load_workbook(file_path, data_only=True)
#     sheet_names = wb.sheetnames[:3]
#     ws1 = wb[sheet_names[0]]
#     ws2 = wb[sheet_names[1]]
#     ws3 = wb[sheet_names[2]]
#     # Get all rows as lists of values
#     rows1 = list(ws1.values)
#     rows2 = list(ws2.values)
#     rows3 = list(ws3.values)
#     # Find the minimum number of rows to avoid index errors
#     min_rows = min(len(rows1), len(rows2), len(rows3))
#     combined_rows = []
#     for i in range(min_rows):
#         part1 = list(rows1[i])[:130]
#         part2 = list(rows2[i])[:94]
#         part3 = list(rows3[i])[:102]
#         combined_row = part1 + part2 + part3
#         combined_rows.append(combined_row)
#     # Write to a new Excel file
#     new_wb = px.Workbook()
#     new_ws = new_wb.active
#     new_ws.title = "Combined"
#     for row in combined_rows:
#         new_ws.append(row)
#     new_wb.save(f"{name}_combined.xlsx")

# # CODICE PER CREARE DIZIONARIO CON CATEGORIE E MACRO ETICHETTE ED ETICHETTE
# def parse_macro_etichette(docx_path):

#     doc = Document(docx_path)
#     lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

#     result = defaultdict(dict)
#     current_category = None

#     for line in lines:
#         if line.startswith("CAT_"):
#             current_category = line.replace("CAT_", "", 1).strip()
#         elif ':' in line and current_category:
#             # Find macro-label (before colon)
#             macro_label_match = re.match(r'^\s*([^:]+):\s*(.*)', line)
#             if macro_label_match:
#                 macro_label = macro_label_match.group(1).strip()
#                 labels_str = macro_label_match.group(2)
#                 # Remove trailing full stop and split by comma
#                 labels = [lbl.strip().rstrip('.') for lbl in labels_str.split(',') if lbl.strip()]
#                 # Remove empty string possibly caused by trailing period
#                 labels = [lbl for lbl in labels if lbl]
#                 result[current_category][macro_label] = labels
#     return dict(result)
# macro_dict = parse_macro_etichette("macro-etichette-per-riclassificazione.docx")

# # CODICE PER STAMPARE IL DIZIONARIO COME ALBERO
# def print_macro_dict_tree(d, indent=0):
#     for cat, macros in d.items():
#         print(" " * indent + f"{cat}:")
#         for macro, labels in macros.items():
#             print(" " * (indent + 2) + f"{macro}:")
#             for label in labels:
#                 print(" " * (indent + 4) + f"{label}")

# # Print the dictionary with categories, macro labels and labels as a tree structure on a file:
# output_file = stampa_su_file_ON("dizionario-macro-etichette.txt")
# print_macro_dict_tree(macro_dict)
# # # print(macro_dict)
# stampa_su_file_OFF(output_file)

# # Build an inverse dictionary mapping each label to its macro-label and category
# label_to_macro_cat = {}
# for category, macros in macro_dict.items():
#     for macro, labels in macros.items():
#         for label in labels:
#             label_to_macro_cat[label] = {"macro_label": macro, "category": category}

# # Print the inverse dictionary on a file:
# output_file = stampa_su_file_ON("dizionario-inverso-macro-etichette.txt")
# for i, (label, info) in enumerate(label_to_macro_cat.items()):
#     # print(f"{label}: macro_label={info['macro_label']}, category={info['category']}")
#     print(f"{label}: {info['macro_label']}, {info['category']}")
# stampa_su_file_OFF(output_file)

# # Check for missing labels in label_to_macro_cat for each combined_cleaned file
# output_file = stampa_su_file_ON("etichette-non-in-dizionario.txt")
# for name, file_path in autori_combined_cleaned_file.items():
#     wb = px.load_workbook(file_path, data_only=True)
#     sheet = wb["Combined"]
#     headers = list(sheet.iter_rows(min_row=2, max_row=2, min_col=1, values_only=True))[0]
#     missing = [h for h in headers if h and h not in label_to_macro_cat]
#     if missing:
#         print(f"{name} - missing labels ({len(missing)}):")
#         for m in missing:
#             print(f"  {m}")
#     else:
#         print(f"{name} - all labels found in dictionary.")
# stampa_su_file_OFF(output_file)

# # Check labels in label_to_macro_cat that are missing from the header row of each combined_cleaned file
# output_file = stampa_su_file_ON("etichette-non-in-excel-files.txt")
# for name, file_path in autori_combined_cleaned_file.items():
#     wb = px.load_workbook(file_path, data_only=True)
#     sheet = wb["Combined"]
#     # Read the header row (row 2, since openpyxl is 1-based and header=1 means second row)
#     headers = list(sheet.iter_rows(min_row=2, max_row=2, min_col=1, values_only=True))[0]
#     not_in_headers = [label for label in label_to_macro_cat if label not in headers]
#     if not_in_headers:
#         print(f"{name} - labels in dictionary but not in header row 2 ({len(not_in_headers)}):")
#         for label in not_in_headers:
#             print(f"  {label}")
#     else:
#         print(f"{name} - all dictionary labels found in header row 2.")
# stampa_su_file_OFF(output_file)

etichette_non_in_dizionario = set()
# these are the number of the columns that are not in the dictionary
etichette_non_in_dizionario.update([1, 2, 3, 4, 5, 6, 111, 112, 204, 205])

# # Process each combined_cleaned file, replacing labels with macro labels
# # and replacing non-empty cells with "X" unless the column is in etichette_non_in_dizionario.
# for name, file_path in autori_combined_cleaned_file.items():
#     wb = px.load_workbook(file_path, data_only=True)
#     sheet = wb["Combined"]
#     # Read all rows as lists of values
#     rows = list(sheet.values)
#     num_rows = len(rows)
#     num_cols = sheet.max_column
#     # Prepare new rows for the new sheet
#     new_rows = []
#     # Copy the first row, but only cells with column index 2, 112, 205 which have the text of the questions
#     if num_rows > 0:
#         selected_indices = [2, 112, 205]
#         new_first_row = []
#         for col_idx in range(1, num_cols + 1):
#             if col_idx in selected_indices:
#                 new_first_row.append(rows[0][col_idx - 1])
#             else:
#                 new_first_row.append("")
#         new_rows.append(new_first_row)
#     # Prepare the second row (with possible replacements of labels with their macro labels
#     if num_rows > 1:
#         new_second_row = []
#         for col_idx in range(num_cols):
#             value = rows[1][col_idx]
#             if value in label_to_macro_cat:
#                 macro_part = label_to_macro_cat[value]["macro_label"]
#                 category_part = label_to_macro_cat[value]["category"]
#                 new_label = f"{category_part} -- {macro_part}"
#                 new_second_row.append(new_label.lower())  # Convert to lowercase
#             else:
#                 new_second_row.append(value)
#         new_rows.append(new_second_row)
#         # Copy the rest of the rows, replacing non-empty cells with "X" unless the column is in etichette_non_in_dizionario
#         for i in range(2, num_rows):
#             new_row = []
#             for col_idx, cell in enumerate(rows[i], start=1):  # openpyxl columns are 1-based
#                 if col_idx in etichette_non_in_dizionario:
#                     new_row.append(cell)
#                 else:
#                     if cell is not None and str(cell).strip() != "":
#                         new_row.append("X")
#                     else:
#                        new_row.append("")
#             new_rows.append(new_row)
#     # Write to a new Excel file
#     new_wb = px.Workbook()
#     new_ws = new_wb.active
#     new_ws.title = "Processed"
#     for row in new_rows:
#         new_ws.append(row)
#     new_file = f"{name}_macroed.xlsx"
#     new_wb.save(new_file)

# # Process each macroed file to merge columns with the same header
# # and write the results to a new Excel file.
# for name, file_path in autori_macroed_file.items():
#     wb = px.load_workbook(file_path, data_only=True)
#     sheet = wb["Processed"]
#     rows = list(sheet.values)
#     # Headers are in row 2 (index 1)
#     headers = list(rows[1])
#     col_indices_by_header = defaultdict(list)
#     for idx, header in enumerate(headers):
#         if header is not None and str(header).strip() != "":
#             col_indices_by_header[header].append(idx)
#     # Prepare merged columns
#     merged_headers = list(col_indices_by_header.keys())
#     merged_rows = []
#     # First row: copy as is (usually question text or similar)
#     merged_rows.append([rows[0][col_indices_by_header[h][0]] for h in merged_headers])
#     # Second row: merged headers
#     merged_rows.append(merged_headers)
#     # For each data row, merge columns with the same header
#     for row in rows[2:]:
#         merged_row = []
#         for h in merged_headers:
#             indices = col_indices_by_header[h]
#             # If only one index and it is in etichette_non_in_dizionario, copy the cell value as is
#             if len(indices) == 1 and (indices[0] + 1) in etichette_non_in_dizionario:
#                 cell_value = row[indices[0]] if indices[0] < len(row) else None
#                 merged_row.append(cell_value)
#             else:
#                 # If any of the merged cells is "X", set "X", else empty
#                 cell_values = [row[i] if i < len(row) else None for i in indices]
#                 if any(str(cell).strip().upper() == "X" for cell in cell_values if cell is not None):
#                     merged_row.append("X")
#                 else:
#                     merged_row.append("")
#         merged_rows.append(merged_row)
#     # Write to new Excel file
#     new_wb = px.Workbook()
#     new_ws = new_wb.active
#     new_ws.title = "Merged"
#     for row in merged_rows:
#         new_ws.append(row)
#     new_file = f"{name}_macroed-merged.xlsx"
#     new_wb.save(new_file)

# Format specified columns in each macroed_merged file and save as new Excel files
columns_to_enlarge = [2, 51, 76]  # Change as needed
for name, file_path in autori_macroed_merged_file.items():
    wb = px.load_workbook(file_path)
    ws = wb["Merged"]
    for col_idx in columns_to_enlarge:
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 70  # ~500 pixels (1 Excel width ≈ 7 pixels)
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='center')
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        if col_idx not in columns_to_enlarge:
            ws.column_dimensions[col_letter].width = 15    
    new_file = f"{name}_formatted.xlsx"
    wb.save(new_file)

# Format specified columns in each formatted file to center align vertically
# and horizontally, except for specified columns that should not be center aligned vertically.
columns_not_to_center_align_vertically = [2, 51, 76]  # Change as needed
for name, file_path in autori_formatted_file.items():
    wb = px.load_workbook(file_path)
    ws = wb["Merged"]
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        if col_idx not in columns_not_to_center_align_vertically:
            for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    wb.save(file_path)

exit()

