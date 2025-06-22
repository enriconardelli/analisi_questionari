# This script reads specific Excel files and prints the second row of a specified sheet
# to a text file named "lista-etichette-per-autori.txt".

import pandas as pd
import openpyxl as px
import sys
import os
from collections import Counter

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
for name, file_path in autori_labeled_file.items():
    wb = px.load_workbook(file_path, data_only=True)
    sheet_names = wb.sheetnames[:3]
    ws1 = wb[sheet_names[0]]
    ws2 = wb[sheet_names[1]]
    ws3 = wb[sheet_names[2]]

    # Get all rows as lists of values
    rows1 = list(ws1.values)
    rows2 = list(ws2.values)
    rows3 = list(ws3.values)

    # Find the minimum number of rows to avoid index errors
    min_rows = min(len(rows1), len(rows2), len(rows3))

    combined_rows = []
    for i in range(min_rows):
        part1 = list(rows1[i])[:130]
        part2 = list(rows2[i])[:94]
        part3 = list(rows3[i])[:102]
        combined_row = part1 + part2 + part3
        combined_rows.append(combined_row)

    # Write to a new Excel file
    new_wb = px.Workbook()
    new_ws = new_wb.active
    new_ws.title = "Combined"
    for row in combined_rows:
        new_ws.append(row)
    new_wb.save(f"{name}_combined.xlsx")

exit()


# Prepare a dictionary to collect first rows
data = {}
for name, file_path in autori_file.items():
    df = pd.read_excel(file_path, header=0, sheet_name="Domanda 3")
    # header=0 means the first row is used as column names
    # then the first row of data will be the second row in the Excel sheet
    # iloc[0] is the FIRST row with data (0-based index)
    etichette = df.iloc[0].tolist() # Convert to list
    if len(etichette) != len(set(etichette)):
        print(f"Warning: {name} has duplicate values in etichette.")
    data[name] = etichette


# Create a DataFrame where each key becomes a column
result_df = pd.DataFrame(data)

# Save to a new Excel file
result_df.to_excel('first_rows_combined.xlsx', index=False)