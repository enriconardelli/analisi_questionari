import pandas as pd
import openpyxl
import pprint
from collections import Counter

# def confronta(a1, a2, a3, a4, a5, a6):
#     disagreements = {}
#     all_keys = set(a1.keys()).union(a2.keys(), a3.keys(), a4.keys(), a5.keys(), a6.keys()) # probabilmente non serve
#     #print("all_keys", all_keys)
#     for key in all_keys:
#         values = [a1.get(key), a2.get(key), a3.get(key), a4.get(key), a5.get(key), a6.get(key)]
#         unique_values = set(frozenset(v) if isinstance(v, set) else v for v in values)
#         if len(unique_values) > 1:  # If there is disagreement
#             disagreements[key] = len(values) - max(values.count(v) for v in unique_values)
#     return disagreements



autori_file = {
    "ADZ": "questionario-insegnanti-campione-60-ADZ.xlsx",
    "CM": "questionario-insegnanti-campione-60-CM.xlsx",
    "EN": "questionario-insegnanti-campione-60-EN.xlsx",
    "GA": "questionario-insegnanti-campione-60-GA.xlsx",
    "LF": "questionario-insegnanti-campione-60-LF.xlsx",
    "ML": "questionario-insegnanti-campione-60-ML.xlsx"
}

NUMAUTORI = len(autori_file)

domande = ["Domanda1", "Domanda 2", "Domanda 3"]

dataframes = {}

for macroetichetta, file_name in autori_file.items():
    file_path = f"{file_name}"
    dataframes[macroetichetta] = {}
    for i, domanda in enumerate(domande, start=0):
        df = pd.read_excel(file_path, sheet_name=i, skiprows=1, dtype=str)
        df.fillna('', inplace=True)  # Replace NaN with empty strings
        df.set_index(df.columns[0], inplace=True)  # Set the first column (ID) as the primary key
        dataframes[macroetichetta][domanda] = df

#print(dataframes["ML"])


etichette_semplici = ["ore sett.", "pacchetto ore", "altro tempo", "classe"]

file_path_estraggo = "estraggo.xlsx"
wb = openpyxl.load_workbook(file_path_estraggo, data_only=True)

etichette_complesse = {}

for domanda in domande:
    sheet = wb[domanda]
    etichette_complesse[domanda] = {}
    
    keys = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    possibili_etichette = list(sheet.iter_rows(min_row=2, max_row=2, values_only=True))[0]

    current_key = None
    for col, macroetichetta in enumerate(keys):
        if macroetichetta:  # Non-empty cell in the first row
            current_key = macroetichetta
            etichette_complesse[domanda][current_key] = []
        if current_key and possibili_etichette[col]:  # Add values under the current key
            etichette_complesse[domanda][current_key].append(possibili_etichette[col])

def print_etichette_complesse_as_tree(etichette_complesse):
    for domanda, keys in etichette_complesse.items():
        print(domanda)
        for key, values in keys.items():
            print(f"  {key}")
            for value in values:
                print(f"    {value}")

#print_etichette_complesse_as_tree(etichette_complesse)


d1 = {}

# Get the IDs from the first dataframe (all dataframes have the same IDs)
ids = dataframes["ADZ"]["Domanda1"].index

for id_ in ids:
    d1[id_] = {}
    for autore, df in dataframes.items():
        d1[id_][autore] = {}
        
        # Add etichette_semplici
        for possibili_etichette in etichette_semplici:
            d1[id_][autore][possibili_etichette] = df["Domanda1"].at[id_, possibili_etichette]
        
        # Add etichette_complesse
        for macroetichetta, possibili_etichette in etichette_complesse["Domanda1"].items():
            d1[id_][autore][macroetichetta] = set()
            for etichetta in possibili_etichette: 
                crocetta = df["Domanda1"].at[id_, etichetta]
                if etichetta.startswith("ETICHE"): #etichette aggiuntive
                    if crocetta != "":
                        for item in crocetta.split(","):
                            d1[id_][autore][macroetichetta].add(item.strip())
                else:
                    if "X" in crocetta:
                        d1[id_][autore][macroetichetta].add(etichetta)
                    elif "dedo" in crocetta:
                        d1[id_][autore][macroetichetta].add(etichetta+"-dedot")
                    elif "?" in crocetta:
                        d1[id_][autore][macroetichetta].add(etichetta+"-???????")
                    elif crocetta.isdigit():
                        d1[id_][autore][macroetichetta].add(crocetta + etichetta)

d2 = {}

for id_ in ids:
    d2[id_] = {}
    for autore, df in dataframes.items():
        d2[id_][autore] = {}
        
        # Add etichette_complesse
        for macroetichetta, possibili_etichette in etichette_complesse["Domanda 2"].items():
            d2[id_][autore][macroetichetta] = set()
            for etichetta in possibili_etichette: 
                crocetta = df["Domanda 2"].at[id_, etichetta]
                if etichetta.startswith("ETICHE"): #etichette aggiuntive
                    if crocetta != "":
                        for item in crocetta.split(","):
                            d2[id_][autore][macroetichetta].add(item.strip())
                else:
                    if "X" in crocetta:
                        d2[id_][autore][macroetichetta].add(etichetta)
                    elif "dedo" in crocetta:
                        d2[id_][autore][macroetichetta].add(etichetta+"-dedot")
                    elif "?" in crocetta:
                        d2[id_][autore][macroetichetta].add(etichetta+"-???????")


d3 = {}

for id_ in ids:
    d3[id_] = {}
    for autore, df in dataframes.items():
        d3[id_][autore] = {}
        
        # Add etichette_complesse
        for macroetichetta, possibili_etichette in etichette_complesse["Domanda 3"].items():
            d3[id_][autore][macroetichetta] = set()
            for etichetta in possibili_etichette: 
                crocetta = df["Domanda 3"].at[id_, etichetta]
                if etichetta.startswith("ETICHE"): #etichette aggiuntive
                    if crocetta != "":
                        for item in crocetta.split(","):
                            d3[id_][autore][macroetichetta].add(item.strip())
                else:
                    if "X" in crocetta:
                        d3[id_][autore][macroetichetta].add(etichetta)
                    elif "dedo" in crocetta:
                        d3[id_][autore][macroetichetta].add(etichetta+"-dedot")
                    elif "?" in crocetta:
                        d3[id_][autore][macroetichetta].add(etichetta+"-???????")

#Print the resulting dictionary for verification
#pprint.pprint(d1["654"]["ADZ"])

# print("Compare d1:")
# for id_ in ids:
#     a, b, c, d, e, f = autori_file.keys()
#     n = confronta(d1[id_][a], d1[id_][b], d1[id_][c], d1[id_][d], d1[id_][e], d1[id_][f])
#     if n != {}:
#         print(f"Difference for ID {id_}: {n}")

data = {"d1": d1, "d2": d2, "d3": d3}

etichette_dict = {}

for level, level_data in data.items():  # Iterate over d1, d2, d3
    etichette_dict[level] = {}
    for id_, autori in level_data.items():
        for autore, macroetichetta_dict in autori.items():
            for macroetichetta, etichette in macroetichetta_dict.items():
                if isinstance(etichette, set):
                    if macroetichetta not in etichette_dict[level]:
                        etichette_dict[level][macroetichetta] = set()
                    etichette_dict[level][macroetichetta].update(etichette)

#print("!!!!!!!!!!!!!!!!!!!!", etichette_dict)



#print("!!!!!!!!!!!!!!!!!!!! dizionario data", data)

# Function to count occurrences of all etichette in a nested dictionary
def count_etichette(data):
    counter = Counter()
    for level, level_data in data.items():  # Iterate over d1, d2, d3
        for id_, autori in level_data.items():
            for autore, macroetichetta_dict in autori.items():
                for macroetichetta, etichette in macroetichetta_dict.items():
                    if isinstance(etichette, set):
                        counter.update(etichette)
    return counter

# Get the count of all etichette as a dictionary
etichette_count = dict(count_etichette(data))
# Print the resulting dictionary for verification
# pprint.pprint(etichette_count)

# def count_etichette_advanced(data, etichette_count):
#     data_with_count = {}
#     for level, level_data in data.items():  # Iterate over d1, d2, d3
#         data_with_count[level] = {}
#         for id_, autori in level_data.items():
#             data_with_count[level][id_] = {}
#             for autore, macroetichetta_dict in autori.items():
#                 data_with_count[level][id_][autore] = {}
#                 for macroetichetta, etichette in macroetichetta_dict.items():
#                     if isinstance(etichette, set):
#                         data_with_count[level][id_][autore][macroetichetta] = {
#                             etichetta: etichette_count[etichetta] for etichetta in etichette
#                         }
#     return data_with_count

# Get the count of all etichette as a dictionary
#etichette_count_advanced = count_etichette_advanced(data, etichette_count)
# Print the resulting dictionary for verification
#pprint.pprint(etichette_count_advanced)


def count_etichette_as_tree(data, etichette_count):

    label_tree={}

    for level, level_data in data.items():  # Iterate over d1, d2, d3
        for id_, autori in level_data.items():
            for autore, macroetichetta_dict in autori.items():
                for macroetichetta, etichette in macroetichetta_dict.items():
                    if isinstance(etichette, set):
                        if level not in label_tree:
                            label_tree[level] = {}
                        if macroetichetta not in label_tree[level]:
                            label_tree[level][macroetichetta] = {}
                        for etichetta in etichette:
                            label_tree[level][macroetichetta][etichetta] = etichette_count[etichetta]
                        
    return label_tree
# Get the count of all etichette as a dictionary
label_tree = count_etichette_as_tree(data, etichette_count)
# Print the resulting dictionary for verification
# pprint.pprint(etichette_count)


print(f"\n\n========== Per ogni domanda, categoria, etichetta: numero di occorrenze e numero normalizzato (/{NUMAUTORI})")
for level, macroetichetta_dict in label_tree.items():
    print(level.upper())
    for macroetichetta, etichette in macroetichetta_dict.items():
        print(f"  {macroetichetta}")
        for etichetta, count in sorted(etichette.items(), key=lambda x: x[1], reverse=True):
            normalized_count = count / NUMAUTORI
            print(f"    {etichetta:<70} {count:>5} {normalized_count:>10.2f}")





# Calcolo della "problematicità" delle etichette
def calcola_problematicita(data):
    problematicita = {}
    # for level, level_data in data.items():  # Iterate over d1, d2, d3
    #     for id_, autori in level_data.items():
    #         for autore, macroetichetta_dict in autori.items():
    #             for macroetichetta, etichette in macroetichetta_dict.items():
    #                 if isinstance(etichette, set):
    #                     for etichetta in etichette:
    #                         if etichetta not in problematicita:
    #                             problematicita[etichetta] = {"R": 0}
    #                         problematicita[etichetta]["R"] += 1




    for domanda in etichette_dict:
        for macroetichetta in etichette_dict[domanda]:
            for etichetta in etichette_dict[domanda][macroetichetta]:        
                problematicita[etichetta] = {"R": 0}
                d1 = data[domanda]
                for id_ in d1:
                    conta = False
                    for autore in d1[id_]:
                        if etichetta in d1[id_][autore][macroetichetta]:
                            conta = True
                    if conta:
                        problematicita[etichetta]["R"] += 1 
    
    #print(problematicita)

    for etichetta, indici in problematicita.items():
        problematicita[etichetta]["PR_L"] = 0
        for domanda in data:
            for i in data[domanda]:
                C_i = 0
                for autore in data[domanda][i]:
                    for macroetichetta in data[domanda][i][autore]:
                        if etichetta in data[domanda][i][autore][macroetichetta]:
                            C_i += 1
                if C_i > 0:
                    R_i = NUMAUTORI - C_i
                    problematicita[etichetta]["PR_L"] += R_i
        problematicita[etichetta]["PR_L"] /= indici["R"]

    return problematicita

# Calcolo della problematicità
etichette_problematicita = calcola_problematicita(data)

# Stampa dei risultati

print("\n\n========== Problematicità: Etichetta, R, PR_L in ordine decrescente di probematicità")
for etichetta, values in sorted(etichette_problematicita.items(), key=lambda x: x[1]["PR_L"], reverse=True):
    print(f"{etichetta:<70} {values['R']:>5} {values['PR_L']:>10.2f}")

# Problematicità per ogni risposta separatamente per d1, d2, d3
for level in data.keys():  # Iterate over d1, d2, d3
    print(f"\n\n========== Problematicità per {level.upper()}")
    for T in ids:
        Q = set()
        for autore in autori_file.keys():
            for macroetichetta, etichette in data[level][T][autore].items():
                Q.update(etichette)

        Q = list(Q)
        problematicita_T = 0

        for etichetta_i in Q:
            C_i = sum(
                1 for autore in autori_file.keys()
                if any(etichetta_i in data[level][T][autore][macroetichetta]
                       for macroetichetta in data[level][T][autore])
            )
            T_i = NUMAUTORI - C_i
            problematicita_T += T_i

        PR_T = problematicita_T / len(Q) if Q else None
        print(f"Risposta {T}: Problematicità {PR_T}")


print(f"\n\n========== Per ogni domanda, categoria, etichetta in alfabetico: numero di occorrenze e problematicità")
for level, macroetichetta_dict in label_tree.items():
    print(level.upper())
    for macroetichetta, etichette in macroetichetta_dict.items():
        print(f"  {macroetichetta}")
        for etichetta, count in sorted(etichette.items(), key=lambda x: x[0].lower(), reverse=False):
            normalized_count = count / NUMAUTORI
            print(f"    {etichetta:<70} {count:>5} {etichette_problematicita[etichetta]['PR_L']:>10.2f}")