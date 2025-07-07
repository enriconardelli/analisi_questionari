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
from openpyxl.utils import column_index_from_string
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

autori_substituted_file = {
    "ADZ": "ADZ_substituted.xlsx",
    "CM": "CM_substituted.xlsx",
    "EN": "EN_substituted.xlsx",
    "GA": "GA_substituted.xlsx",
    "LF": "LF_substituted.xlsx",
    "ML": "ML_substituted.xlsx"
}

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

autori_ordered_file = {
    "ADZ": "ADZ_ordered.xlsx",
    "CM": "CM_ordered.xlsx",
    "EN": "EN_ordered.xlsx",
    "GA": "GA_ordered.xlsx",
    "LF": "LF_ordered.xlsx",
    "ML": "ML_ordered.xlsx"
}

autori_formatted_file = {
    "ADZ": "ADZ_formatted.xlsx",
    "CM": "CM_formatted.xlsx",
    "EN": "EN_formatted.xlsx",
    "GA": "GA_formatted.xlsx",
    "LF": "LF_formatted.xlsx",
    "ML": "ML_formatted.xlsx"
}

# CODICE PER STAMPARE UN DIZIONARIO COME ALBERO
def print_dict_as_tree(d, indent=0):
    for cat, macros in d.items():
        print(" " * indent + f"{cat}:")
        for macro, labels in macros.items():
            print(" " * (indent + 2) + f"{macro}:")
            for label in labels:
                print(" " * (indent + 4) + f"{label}")

# # CODICE PER VERIFICARE DUPLICAZIONI ETICHETTE DEI FILE CON I FOGLI COMBINATI
# # NORMALMENTE QUESTO CODICE NON VA ESEGUITO, MA SOLO IN CASO DI PROBLEMI CON LE ETICHETTE
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

# # CODICE PER SOSTITUIRE I NOMI DELLE COLONNE
# # This code replaces specific headers in the second row of each sheet in the Excel files
# # with their corresponding substitutions defined in the header_substitutions_by_sheet dictionary
# # which derives from the analysis done during the first reclassification attempt
# header_substitutions_by_sheet = {
#     "Domanda1": {
#         "costruzione mappe": "attività costruzione di mappe",
#         "sviluppo contenuti2": "attività sviluppo contenuti",
#         "ludico-motoria2": "attività ludico-motoria",
#         "uso della tecnologia": "attività uso della tecnologia",
#         "uso della tecnologia2": "uso della tecnologia",
# },
#     "Domanda 2": {
#         "comprensione dei comandi di progrmmazione": "comprensione dei comandi di programmazione",
#     },
#     "Domanda 3": {
#         "osservazione, memoria, analisi": "osservazione/ memoria/ analisi"
#     }
# }
# for name, file_path in autori_file.items():
#     wb = px.load_workbook(file_path)
#     for sheet_name in wb.sheetnames[:3]:
#         ws = wb[sheet_name]
#         rows = list(ws.values)
#         if len(rows) < 2:
#             continue
#         headers = list(rows[1])
#         # Get substitutions for this sheet, or empty dict if none
#         substitutions = header_substitutions_by_sheet.get(sheet_name, {})
#         new_headers = [
#             substitutions.get(str(h).strip(), h) if h is not None else None
#             for h in headers
#         ]
#         for col_idx, new_header in enumerate(new_headers, start=1):
#             ws.cell(row=2, column=col_idx, value=new_header)
#     new_file = f"{name}_substituted.xlsx"
#     wb.save(new_file)

# # CODICE PER FILTRARE LE COLONNE
# # This code filters out columns starting with "ETICHETTE" or "ETICHETE" from the first three sheets of each workbook.
# for name, file_path in autori_substituted_file.items():
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

# # CODICE PER ETICHETTARE LE COLONNE CON LE DOMANDE
# # For each Excel file in autori_filtered_file, create a new Excel file with with column names which are the concatenation of Dn + original column name
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

# # CODICE PER COMBINARE I 3 FOGLI IN UNO SOLO
# # This code combines the first three sheets of each workbook into a new workbook,
# # taking the first 130 columns from the first sheet, 94 from the second, and 102 from the third.
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

# # CODICE PER RIMUOVERE DAI FILE COMBINATI LE COLONNE CHE NON INTERESSANO
# # BS: D1_rianalizzare
# # [DF-DX] tutte le colonne con i dettagli della piattaforma/ambiente
# # HP: D2_ambiguità
# # Remove columns BS, DF, DG, DH, HP (A1 reference) from each combined file and save as cleaned file
# columns_to_remove = ["BS", "HP"]
# # Add columns from DF to DX (inclusive) to columns_to_remove
# df_index = column_index_from_string("DF")
# dx_index = column_index_from_string("DX")
# columns_to_remove += [get_column_letter(i) for i in range(df_index, dx_index + 1)]
# for name, file_path in autori_combined_file.items():
#     wb = px.load_workbook(file_path, data_only=True)
#     ws = wb["Combined"]
#     # Get column indices (1-based) for columns to remove
#     col_indices_to_remove = [column_index_from_string(col) for col in columns_to_remove]
#     # Read all rows as lists of values
#     rows = list(ws.values)
#     # Remove the specified columns from each row
#     cleaned_rows = []
#     for row in rows:
#         cleaned_row = [cell for idx, cell in enumerate(row, start=1) if idx not in col_indices_to_remove]
#         cleaned_rows.append(cleaned_row)
#     # Write to a new Excel file
#     new_wb = px.Workbook()
#     new_ws = new_wb.active
#     new_ws.title = "Combined"
#     for row in cleaned_rows:
#         new_ws.append(row)
#     new_file = f"{name}_combined-cleaned.xlsx"
#     new_wb.save(new_file)

# # # ----------------------------------------------------------------------
# # # ----------------------------------------------------------------------

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
# macro_dict = parse_macro_etichette("macro-etichette-per-riclassificazione-FINALE.docx")

# # Build an inverse dictionary mapping each label to its macro-label and category and checking unicity of labels
# label_to_macro_cat = {}
# for category, macros in macro_dict.items():
#     for macro, labels in macros.items():
#         for label in labels:
#             if label in label_to_macro_cat:
#                 print(f"WARNING: label '{label}' already exists in label_to_macro_cat (category: {label_to_macro_cat[label]['category']}, macro_label: {label_to_macro_cat[label]['macro_label']})")
#             label_to_macro_cat[label] = {"macro_label": macro, "category": category}

# # Print the dictionary with categories, macro labels and labels as a tree structure on a file:
# output_file = stampa_su_file_ON("dizionario-macro-etichette-FINALE.txt")
# print_dict_as_tree(macro_dict)
# stampa_su_file_OFF(output_file)

# # Print the inverse dictionary on a file:
# output_file = stampa_su_file_ON("dizionario-inverso-macro-etichette-FINALE.txt")
# for i, (label, info) in enumerate(label_to_macro_cat.items()):
#     # print(f"{label}: macro_label={info['macro_label']}, category={info['category']}")
#     print(f"{label}: {info['macro_label']}, {info['category']}")
# stampa_su_file_OFF(output_file)

# # Check for missing labels in label_to_macro_cat for each combined_cleaned file
# output_file = stampa_su_file_ON("etichette-non-in-dizionario-FINALE.txt")
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

# # Check for duplicated labels in the header row (row 2) of each combined_cleaned file
# output_file = stampa_su_file_ON("etichette-duplicate-in-excel-files-FINALE.txt")
# for name, file_path in autori_combined_cleaned_file.items():
#     wb = px.load_workbook(file_path, data_only=True)
#     sheet = wb["Combined"]
#     headers = list(sheet.iter_rows(min_row=2, max_row=2, min_col=1, values_only=True))[0]
#     header_counts = Counter(headers)
#     duplicates = [label for label, count in header_counts.items() if label and count > 1]
#     if duplicates:
#         print(f"{name} - duplicated labels in header row 2 ({len(duplicates)}):")
#         for label in duplicates:
#             print(f"  {label}: {header_counts[label]} times")
#     else:
#         print(f"{name} - no duplicated labels in header row 2.")
# stampa_su_file_OFF(output_file)

# # Check labels in label_to_macro_cat that are missing from the header row of each combined_cleaned file
# output_file = stampa_su_file_ON("etichette-non-in-excel-files-FINALE.txt")
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

# -------------------------------------------------------------------
# ----------------------------------------------------------------------

# # CODICE PER TROVARE L'INDICE DI COLONNA DELLE ETICHETTE DEI FILE EXCEL CHE NON SONO PRESENTI IN DIZIONARIO
# nomi_etichette_non_in_dizionario = {"D1_ID", "D1_Domanda 1", "D1_ore sett.", "D1_pacchetto ore", "D1_altro tempo", "D1_classe", "D2_ID", "D2_Domanda 2", "D3_ID", "D3_Domanda 3"}
# etichette_non_in_dizionario = set()
# excel_wb = px.load_workbook("ADZ_combined-cleaned.xlsx", data_only=True)
# excel_ws = excel_wb.active  # or specify the sheet name if needed
# headers = list(excel_ws.iter_rows(min_row=2, max_row=2, min_col=1, values_only=True))[0]
# for idx, header in enumerate(headers, start=1):
#     if header in nomi_etichette_non_in_dizionario:
#         etichette_non_in_dizionario.add(idx)
# # print(f"etichette_non_in_dizionario: {etichette_non_in_dizionario}")

# # CODICE PER SOSTITUIRE LE ETICHETTE CON LE RISPETTIVE MACRO ETICHETTE IN PREVISIONE DELLA LORO FUSIONE
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
#                 new_label = f"{category_part}\n\n{macro_part}"
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

# # CODICE PER UNIRE LE COLONNE CHE APPARTENGONO ALLA STESSA MACRO-ETICHETTA
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

# # CODICE PER ORDINARE LE COLONNE SECONDO I REQUISITI DESIDERATI
# initial_labels = ["D1_ID", "D1_Domanda 1", "D1_ore sett.", "D1_pacchetto ore", "D1_altro tempo", "D1_classe"]
# order_of_categories_for_D1 = ["materiali/strumenti fisici", "attività", "piattaforma/ ambiente", "metodologia"] 
# order_of_categories_for_D2_D3 = ["obiettivi disciplinari", "obiettivi pertinenti", "abilità strumentali digitali", "interdisciplinarità", "soft skills/ trasversali/ metacognizione", "risposta generica d2", "focus d2", "risposta generica d3"]
# # Helper to extract category and label from header
# def split_header(header):
#     if header is None:
#         return None, None
#     parts = str(header).split('\n\n')
#     if len(parts) >= 2:
#         return parts[0].strip(), parts[1].strip()
#     else:
#         return str(header).strip(), ""
# # Process each macroed_merged file to order columns according to the specified rules
# for name, file_path in autori_macroed_merged_file.items():
#     wb = px.load_workbook(file_path)
#     ws = wb["Merged"]
#     rows = list(ws.values)
#     if len(rows) < 2:
#         continue
#     question_row = list(rows[0])
#     headers = list(rows[1])
#     data_rows = rows[2:]
#     # Build a mapping from header to column index
#     header_to_index = {str(h): idx for idx, h in enumerate(headers) if h is not None}
#     # print("Header to index mapping:", header_to_index)
#     # 1. Initial labels (in exact order)
#     ordered_indices = [header_to_index[h] for h in initial_labels if h in header_to_index]
#     # print("Ordered indices for initial labels:", ordered_indices)
#     # 2. D1 categories: columns whose header starts with a category in order_of_categories_for_D1
#     d1_indices = []
#     for cat in order_of_categories_for_D1:
#         # Find all headers that start with this category
#         matching = []
#         for idx, h in enumerate(headers):
#             cat_part, label_part = split_header(h)
#             if cat_part == cat:
#                 matching.append((label_part, idx))
#         # Sort by label_part
#         matching.sort()
#         d1_indices.extend(idx for label, idx in matching)
#     # print("D1 category indices:", d1_indices)
#     # 3. D2_Domanda 2 and D3_Domanda 3
#     for_d2d3 = ["D2_Domanda 2", "D2_ID", "D3_Domanda 3"]
#     d2d3_indices = [header_to_index[h] for h in for_d2d3 if h in header_to_index]
#     # print("D2 and D3 indices:", d2d3_indices)
#     # 4. D2/D3 categories: columns whose header starts with a category in order_of_categories_for_D2_D3
#     d2d3_cat_indices = []
#     for cat in order_of_categories_for_D2_D3:
#         matching = []
#         for idx, h in enumerate(headers):
#             cat_part, label_part = split_header(h)
#             if cat_part == cat:
#                 matching.append((label_part, idx))
#         matching.sort()
#         d2d3_cat_indices.extend(idx for label, idx in matching)
#     # print("D2/D3 category indices:", d2d3_cat_indices)
#     # Concatenate all indices in the required order (remove duplicates, keep first occurrence)
#     all_indices = []
#     for group in [ordered_indices, d1_indices, d2d3_indices, d2d3_cat_indices]:
#         for idx in group:
#             if idx not in all_indices:
#                 all_indices.append(idx)
#     # print("All ordered indices:", all_indices)
#     # Prepare new rows
#     new_rows = []
#     new_rows.append([question_row[idx] if idx < len(question_row) else "" for idx in all_indices])
#     new_rows.append([headers[idx] if idx < len(headers) else "" for idx in all_indices])
#     for row in data_rows:
#         new_rows.append([row[idx] if idx < len(row) else "" for idx in all_indices])
#     # Write to new Excel file
#     new_wb = px.Workbook()
#     new_ws = new_wb.active
#     new_ws.title = "Ordered"
#     for row in new_rows:
#         new_ws.append(row)
#     new_file = f"{name}_ordered.xlsx"
#     new_wb.save(new_file)

# # CODICE PER FORMATTARE I FILES
# headers_of_columns_with_answer_text = ["D1_Domanda 1", "D2_Domanda 2", "D3_Domanda 3"]
# for name, file_path in autori_ordered_file.items():
#     wb = px.load_workbook(file_path)
#     ws = wb["Ordered"]
#     headers = list(ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=ws.max_column, values_only=True))[0]
#     header_to_index = {str(h): idx + 1 for idx, h in enumerate(headers) if h is not None}

#     # Set wrap_text=True, vertical='center' for all cells
#     for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
#         for cell in row:
#             cell.alignment = Alignment(wrap_text=True, vertical='center')

#     # Adjust column widths and horizontal alignment
#     for idx, header in enumerate(headers, start=1):
#         col_letter = get_column_letter(idx)
#         if header in headers_of_columns_with_answer_text:
#             ws.column_dimensions[col_letter].width = 70
#         else:
#             ws.column_dimensions[col_letter].width = 15
#             # Set horizontal='center' for all cells in this column
#             for row in ws.iter_rows(min_col=idx, max_col=idx, min_row=1, max_row=ws.max_row):
#                 for cell in row:
#                     cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

#     new_file = f"{name}_formatted.xlsx"
#     wb.save(new_file)

categories_to_consider = ["materiali/strumenti fisici", "attività", "piattaforma/ ambiente", "metodologia", "obiettivi disciplinari", "obiettivi pertinenti", "abilità strumentali digitali", "interdisciplinarità", "soft skills/ trasversali/ metacognizione", "risposta generica d2", "focus d2", "risposta generica d3"]

C_i_base_values = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
# for author, file_path in list(autori_formatted_file.items())[:1]:
for author, file_path in autori_formatted_file.items():
    wb = px.load_workbook(file_path, data_only=True)
    ws = wb["Ordered"]
    # Get headers from row 2
    headers = list(ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=ws.max_column, values_only=True))[0]
    header_to_col = {str(h): idx + 1 for idx, h in enumerate(headers) if h is not None}
    # Find index of D1_ID column
    d1_id_col = header_to_col.get("D1_ID")
    if not d1_id_col:
        continue
    # For each macro-label, get its column index
    macro_label_to_col = {}
    for macro_label in header_to_col:
        if macro_label not in ["D1_ID", "D1_Domanda 1", "D1_ore sett.", "D1_pacchetto ore", "D1_altro tempo", "D1_classe", "D2_ID", "D2_Domanda 2", "D3_Domanda 3"]:
            macro_label_to_col[macro_label] = header_to_col[macro_label]
    # print("macro_label_to_col:", macro_label_to_col)
    # Iterate over data rows (from row 3 onwards)
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=ws.max_column, values_only=True):
        ID = row[d1_id_col - 1]
        if not ID:
            continue
        for macro_label, col_idx in macro_label_to_col.items():
            cell_value = row[col_idx - 1]
            if cell_value == "X":
                C_i_base_values[macro_label][ID][author] += 1

# Print the C_i_base_values dictionary as a tree structure
output_file = stampa_su_file_ON("C_i_base_values_by_macro_etichette-.txt")
for macro_label in C_i_base_values:
    for ID in C_i_base_values[macro_label]:
        for author in C_i_base_values[macro_label][ID]:
            value = C_i_base_values[macro_label][ID][author]
            macro_label_print = macro_label.replace("\n\n", " -- ")
            print(f"Macro-etichetta: {macro_label_print}, ID: {ID}, Author: {author}, Value: {value}")
stampa_su_file_OFF(output_file)

# Sum C_i_base_values across all authors to produce C_i_values
C_i_values = defaultdict(lambda: defaultdict(int))
for macro_label in C_i_base_values:
    for ID in C_i_base_values[macro_label]:
        for author in C_i_base_values[macro_label][ID]:
            C_i_values[macro_label][ID] += C_i_base_values[macro_label][ID][author]

# Print the C_i_values dictionary as a tree structure
output_file = stampa_su_file_ON("C_i_values_by_macro_etichette.txt")
for macro_label in C_i_values:
    macro_label_print = macro_label.replace("\n\n", " -- ")
    print(f"Macro-etichetta: {macro_label_print}")
    for ID in C_i_values[macro_label]:
        print(f"  ID: {ID}, Value: {C_i_values[macro_label][ID]}")
stampa_su_file_OFF(output_file)

PR_L = defaultdict(lambda: [0, 0])
for macro_label in C_i_values:
    how_many_answers = 0
    for ID in C_i_values[macro_label]:
        PR_L[macro_label][0] += 6 - C_i_values[macro_label][ID]
        how_many_answers += 1
    PR_L[macro_label][0] = round(PR_L[macro_label][0] / how_many_answers, 3)
    PR_L[macro_label][1] = how_many_answers

# Print the PR_L dictionary as a tree structure according to the given requirements
output_file = stampa_su_file_ON("macro_etichette_e_problematicita_ordine_categorie.txt")
print("macro-etichetta __ problematicità PR_L __ numero di risposte R in cui è stata usata \n   (ordinato per categoria e macro-etichetta)")
print("-----------------------------------------------------------------------------------")
# Group macro_labels by category and sort macro-labels within each category
macro_labels_by_category = defaultdict(list)
for macro_label in PR_L:
    if "\n\n" in macro_label:
        category, macro = macro_label.split("\n\n", 1)
    else:
        category, macro = macro_label, ""
    macro_labels_by_category[category].append((macro, macro_label))
for category in categories_to_consider:
    if category in macro_labels_by_category:
        # Sort macro-labels alphabetically by the second part
        for macro, macro_label in sorted(macro_labels_by_category[category], key=lambda x: x[0]):
            macro_label_print = macro_label.replace("\n\n", " -- ")
            print(f"{macro_label_print} __ PR_L: {PR_L[macro_label][0]} __ R: {PR_L[macro_label][1]}")
stampa_su_file_OFF(output_file)

# Print the PR_L dictionary as a tree structure, ordered by decreasing PR_L value
output_file = stampa_su_file_ON("macro_etichette_e_problematicita_ordine_problema.txt")
print("macro-etichetta __ problematicità PR_L __ numero di risposte R in cui è stata usata \n   (ordinato per PR_L decrescente)")
print("-----------------------------------------------------------------------------------")
for macro_label, values in sorted(PR_L.items(), key=lambda x: x[1][0], reverse=True):
    macro_label_print = macro_label.replace("\n\n", " -- ")
    print(f"{macro_label_print} \n   PR_L: {values[0]} __ R: {values[1]}")
stampa_su_file_OFF(output_file)

exit()

