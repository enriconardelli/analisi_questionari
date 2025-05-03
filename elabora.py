import pandas as pd
import openpyxl
import pprint
import sys
from collections import Counter

def stampa_su_file_OFF(output_file):
    output_file.close()
    # Reset standard output back to console
    sys.stdout = sys.__stdout__

def stampa_su_file_ON(nome_file):
    output_file = open(nome_file, "w")
    # Redirect standard output to the file
    sys.stdout = output_file
    return output_file

def print_etichette_complesse_as_tree(etichette_complesse):
    # Stampa il dizionario etichette_complesse ricevuto in input
    # in forma di albero, ordinato per domanda, categoria e etichetta
    for domanda, keys in etichette_complesse.items():
        print(domanda)
        for key, values in keys.items():
            print(f"  {key}")
            sorted_values = sorted(values, key=lambda x: x.lower()) # Sort values alphabetically
            for value in sorted_values:
                print(f"    {value}")

def confronta_etichette_usate():
    # Costruisce per ogni autore un dizionario con le etichette usate
    # per ogni domanda e categoria e poi lo contronta con il dizionario
    # dell'autore precedente per vedere se ci sono differenze e stampare
    # le spcifiche differenze indicando le etichette in più e in meno
    # rispetto all'autore precedente
    output_file = stampa_su_file_ON("etichette_complesse_tutti_autori.txt")
    etichette_complesse_prev = {}
    sigla_autore_prev = None
    for sigla_autore, file_name in autori_file.items():
        etichette_complesse = {}
        file_path = f"{file_name}"
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for domanda in domande:
            sheet = wb[domanda]
            etichette_complesse[domanda] = {}
            lista_categorie = list(sheet.iter_rows(min_row=1, max_row=1, min_col=3, values_only=True))[0] # tutti i nomi delle categorie
            lista_etichette = list(sheet.iter_rows(min_row=2, max_row=2, min_col=3, values_only=True))[0] # tutti i nomi delle etichette
            categoria_corrente = None
            for col, categoria in enumerate(lista_categorie):
                if categoria:  # Se c'è un nome di categoria nella prima riga
                    categoria_corrente = categoria
                    etichette_complesse[domanda][categoria_corrente] = []
                if categoria_corrente and lista_etichette[col]:  # Add values under the current key
                    if lista_etichette[col] not in etichette_complesse[domanda][categoria_corrente]:
                        etichette_complesse[domanda][categoria_corrente].append(lista_etichette[col])
        print(f"\n\nCERCA EVENTUALI DIFFERENZE COL PRECEDENTE ETICHETTE COMPLESSE PER {sigla_autore} rispetto a {sigla_autore_prev}:\n")
        if etichette_complesse_prev:
            for domanda in etichette_complesse:
                for categoria in etichette_complesse[domanda]:
                    if categoria not in etichette_complesse_prev[domanda]:
                        print(f"New category in {domanda}: {categoria} --- {sigla_autore} rispetto a {sigla_autore_prev}")
                    else:
                        prev_values = set(etichette_complesse_prev[domanda][categoria])
                        curr_values = set(etichette_complesse[domanda][categoria])
                        if prev_values != curr_values:
                            prev_difference_curr = prev_values.difference(curr_values)
                            curr_difference_prev = curr_values.difference(prev_values)
                            print(f"Difference in {domanda}, category {categoria} --- {sigla_autore} rispetto a {sigla_autore_prev}:")
                            print(f"  add to {sigla_autore}: {prev_difference_curr}")
                            print(f"  add to {sigla_autore_prev}: {curr_difference_prev}")
                            # print(f"  Previous: {prev_values}")
                            # print(f"  Current: {curr_values}")
        etichette_complesse_prev = etichette_complesse
        sigla_autore_prev = sigla_autore
    stampa_su_file_OFF(output_file)

def ricerca_nuove_etichette(dataframes, etichette_semplici, etichette_complesse):
    # Funzione per cercare le nuove etichette inserite dagli autori nella colonna ETICHETTE AGGIUNTE
    # e costruire il dizionario nuove_etichette con le nuove etichette trovate
    # ed eventualmente stamparlo su file. se il valore per una certa categorie è set() 
    # significa che per quella categoria non ci sono nuove etichette
    # la funzione restituisce i dizionari d1, d2, d3 con le etichette usate
    # nelle risposte a tutte le domande da tutti gli autori

    # Get the IDs from the first dataframe (all dataframes have the same IDs)
    ids = dataframes["ADZ"]["Domanda1"].index

    # output_file = stampa_su_file_ON("nuove_etichette.txt")
    nuove_etichette = {}

    d1 = {}
    nuove_etichette["Domanda1"] = {}
    for id_ in ids:
        d1[id_] = {}
        nuove_etichette["Domanda1"][id_] = {}

        for autore, df in dataframes.items():
            d1[id_][autore] = {}
            nuove_etichette["Domanda1"][id_][autore] = {}
            
            # Add etichette_semplici (ci sono solo in Domanda1)
            for etichetta in etichette_semplici:
                d1[id_][autore][etichetta] = df["Domanda1"].at[id_, etichetta]
            
            # Add etichette_complesse
            for categoria, etichette in etichette_complesse["Domanda1"].items():
                d1[id_][autore][categoria] = set()
                nuove_etichette["Domanda1"][id_][autore][categoria] = {}
                for etichetta in etichette: 
                    crocetta = df["Domanda1"].at[id_, etichetta]
                    if etichetta.startswith("ETICHE"): #etichette aggiuntive
                        if crocetta != "":
                            for item in crocetta.split(","):
                                d1[id_][autore][categoria].add(item.strip())
                                nuove_etichette["Domanda1"][id_][autore][categoria] = item.strip()
                                # print(autore, id_, categoria, etichetta, item.strip())
                    else:
                        if "X" in crocetta:
                            d1[id_][autore][categoria].add(etichetta)
                        elif "dedo" in crocetta:
                            d1[id_][autore][categoria].add(etichetta+"-dedot")
                        elif "?" in crocetta:
                            d1[id_][autore][categoria].add(etichetta+"-???????")
                        elif crocetta.isdigit():
                            d1[id_][autore][categoria].add(crocetta + etichetta)

    d2 = {}
    nuove_etichette["Domanda 2"] = {}
    for id_ in ids:
        d2[id_] = {}
        nuove_etichette["Domanda 2"][id_] = {}
        for autore, df in dataframes.items():
            d2[id_][autore] = {}
            nuove_etichette["Domanda 2"][id_][autore] = {}
            
            # Add etichette_complesse
            for categoria, etichette in etichette_complesse["Domanda 2"].items():
                d2[id_][autore][categoria] = set()
                nuove_etichette["Domanda 2"][id_][autore][categoria] = set()
                for etichetta in etichette: 
                    crocetta = df["Domanda 2"].at[id_, etichetta]
                    if etichetta.startswith("ETICHE"): #etichette aggiuntive
                        if crocetta != "":
                            for item in crocetta.split(","):
                                d2[id_][autore][categoria].add(item.strip())
                                nuove_etichette["Domanda 2"][id_][autore][categoria].add(item.strip())
                                # print(autore, id_, categoria, etichetta, item.strip())
                    else:
                        if "X" in crocetta:
                            d2[id_][autore][categoria].add(etichetta)
                        elif "dedo" in crocetta:
                            d2[id_][autore][categoria].add(etichetta+"-dedot")
                        elif "?" in crocetta:
                            d2[id_][autore][categoria].add(etichetta+"-???????")

    d3 = {}
    nuove_etichette["Domanda 3"] = {}
    for id_ in ids:
        d3[id_] = {}
        nuove_etichette["Domanda 3"][id_] = {}
        for autore, df in dataframes.items():
            d3[id_][autore] = {}
            nuove_etichette["Domanda 3"][id_][autore] = {}

            # Add etichette_complesse
            for categoria, etichette in etichette_complesse["Domanda 3"].items():
                d3[id_][autore][categoria] = set()
                nuove_etichette["Domanda 3"][id_][autore][categoria] = set()
                for etichetta in etichette: 
                    crocetta = df["Domanda 3"].at[id_, etichetta]
                    if etichetta.startswith("ETICHE"): #etichette aggiuntive
                        if crocetta != "":
                            for item in crocetta.split(","):
                                d3[id_][autore][categoria].add(item.strip())
                                nuove_etichette["Domanda 3"][id_][autore][categoria].add(item.strip())
                                # print(autore, id_, categoria, etichetta, item.strip())
                    else:
                        if "X" in crocetta:
                            d3[id_][autore][categoria].add(etichetta)
                        elif "dedo" in crocetta:
                            d3[id_][autore][categoria].add(etichetta+"-dedot")
                        elif "?" in crocetta:
                            d3[id_][autore][categoria].add(etichetta+"-???????")

    # pprint.pprint(nuove_etichette)
    # stampa_su_file_OFF(output_file)

    return d1, d2, d3

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

########################################################################
# Legge in risposte[autore][domanda] i file di Excel
# con la classificazione delle risposte di ogni autore a ogni domanda
########################################################################

risposte = {}

for sigla_autore, file_name in autori_file.items():
    file_path = f"{file_name}"
    risposte[sigla_autore] = {}
    for i, domanda in enumerate(domande, start=0):
        df = pd.read_excel(file_path, sheet_name=i, skiprows=1, dtype=str)
        df.fillna('', inplace=True)  # Replace NaN with empty strings
        df.set_index(df.columns[0], inplace=True)  # Set the first column (ID) as the primary key
        risposte[sigla_autore][domanda] = df

etichette_semplici = ["ore sett.", "pacchetto ore", "altro tempo", "classe"]

########################
#
# confronta_etichette_usate()
#
########################

########################################################################
# Generazione in etichette_complesse di tutte le etichette usate nelle risposte 
# assumendo che tutti gli autori abbiano usato lo stesso insieme di
# etichette: questa verifica può essere fatta con la funzione
# confronta_etichette_usate 
########################################################################
a_file_path = "questionario-insegnanti-campione-60-ADZ.xlsx"
wb = openpyxl.load_workbook(a_file_path, data_only=True)
etichette_complesse = {}
for domanda in domande:
    sheet = wb[domanda]
    etichette_complesse[domanda] = {}
    lista_categorie = list(sheet.iter_rows(min_row=1, max_row=1, min_col=3, values_only=True))[0] # tutti i nomi delle categorie
    lista_etichette = list(sheet.iter_rows(min_row=2, max_row=2, min_col=3, values_only=True))[0] # tutti i nomi delle etichette
    categoria_corrente = None
    for col, categoria in enumerate(lista_categorie):
        if categoria:  # Se c'è un nome di categoria nella prima riga
            categoria_corrente = categoria
            etichette_complesse[domanda][categoria_corrente] = []
        if categoria_corrente and lista_etichette[col]:  # Add values under the current key
            if lista_etichette[col] not in etichette_complesse[domanda][categoria_corrente]:
                etichette_complesse[domanda][categoria_corrente].append(lista_etichette[col])

# output_file = stampa_su_file_ON("categorie+etichette_60_risposte.txt")
# print_etichette_complesse_as_tree(etichette_complesse)
# stampa_su_file_OFF(output_file)

########################################################################
# Creazione dei dizionari d1, d2, d3 con tutte le etichette usate 
# nelle risposte alle domande d1, d2, d3 da tutti gli autori
# che contengono anche le eventuali etichette nuove, se sono
# state aggiunte dagli autori nella colonna ETICHETTE AGGIUNTE
########################################################################

d1, d2, d3 = ricerca_nuove_etichette(risposte, etichette_semplici, etichette_complesse)


########################################################################
# Creazione di un unico dizionario etichette_dict[domanda][categoria]
# che contiene tutte le etichette usate nelle risposte
# a tutte le domande da tutti gli autori
########################################################################
dizionari_risposte = {"d1": d1, "d2": d2, "d3": d3}

etichette_dict = {}

for domanda, risposte_alla_domanda in dizionari_risposte.items():  # Iterate over d1, d2, d3
    etichette_dict[domanda] = {}
    for id_, autori in risposte_alla_domanda.items():
        for autore, categorie in autori.items():
            for categoria, etichette in categorie.items():
                if isinstance(etichette, set):
                    if categoria not in etichette_dict[domanda]:
                        etichette_dict[domanda][categoria] = set()
                    etichette_dict[domanda][categoria].update(etichette)
# output_file = stampa_su_file_ON("etichette_dict.txt")
# print("etichette_dict ", etichette_dict)
# stampa_su_file_OFF(output_file)


########################################################################
# Conteggio assoluto e normalizzato sugli autori
# delle occorrenze di ogni etichetta
########################################################################

# Function to count occurrences of all etichette in a nested dictionary
def count_etichette(dizionari_risposte):
    counter = Counter()
    for domanda, risposte_alla_domanda in dizionari_risposte.items():  # Iterate over d1, d2, d3
        for id_, autori in risposte_alla_domanda.items():
            for autore, categorie in autori.items():
                for categoria, etichette in categorie.items():
                    if isinstance(etichette, set):
                        counter.update(etichette)
    return counter

# Costruisce dizionario etichette_count[etichetta] che contiene le occorrenze di ogni etichetta
etichette_count = dict(count_etichette(dizionari_risposte))
# Print the resulting dictionary for verification
# output_file = stampa_su_file_ON("etichette_count.txt")
# pprint.pprint(etichette_count)
# stampa_su_file_OFF(output_file)

def count_etichette_as_tree(dizionari_risposte, etichette_count):
    label_tree={}
    for domanda, risposte_alla_domanda in dizionari_risposte.items():  # Iterate over d1, d2, d3
        for id_, autori in risposte_alla_domanda.items():
            for autore, categorie in autori.items():
                for categoria, etichette in categorie.items():
                    if isinstance(etichette, set):
                        if domanda not in label_tree:
                            label_tree[domanda] = {}
                        if categoria not in label_tree[domanda]:
                            label_tree[domanda][categoria] = {}
                        for etichetta in etichette:
                            label_tree[domanda][categoria][etichetta] = etichette_count[etichetta]
    return label_tree

# Costruisce dizionario etichette_count_as_tree[domanda][categoria][etichetta] 
# che contiene le occorrenze di ogni etichetta per ogni domanda e categoria
etichette_count_as_tree = count_etichette_as_tree(dizionari_risposte, etichette_count)
# Print the resulting dictionary for verification
# output_file = stampa_su_file_ON("etichette_count_as_tree.txt")
# pprint.pprint(label_tree)
# stampa_su_file_OFF(output_file)

output_file = stampa_su_file_ON("etichette_ordinate_conteggio.txt")
print(f"\n\n========== Per ogni domanda, categoria, etichetta: numero di occorrenze e numero normalizzato (/{NUMAUTORI})")
for domanda, categorie in etichette_count_as_tree.items():
    print(domanda.upper())
    for categoria, etichette in categorie.items():
        print(f"  {categoria}")
        for etichetta, count in sorted(etichette.items(), key=lambda x: x[1], reverse=True):
            normalized_count = count / NUMAUTORI
            print(f"    {etichetta:<60} {count:>5} {normalized_count:>10.2f}")
stampa_su_file_OFF(output_file)

output_file = stampa_su_file_ON("etichette_ordinate_alfabeticamente.txt")
print(f"\n\n========== Per ogni domanda, categoria, etichetta: numero di occorrenze e numero normalizzato (/{NUMAUTORI})")
for domanda, categorie in etichette_count_as_tree.items():
    print(domanda.upper())
    for categoria, etichette in categorie.items():
        print(f"  {categoria}")
        for etichetta, count in sorted(etichette.items(), key=lambda x: x[0].lower()):
            normalized_count = count / NUMAUTORI
            print(f"    {etichetta:<60} {count:>5} {normalized_count:>10.2f}")
stampa_su_file_OFF(output_file)


########################################################################
# Calcolo della problematicità delle etichette
########################################################################
def calcola_problematicita(dizionari_risposte):
    problematicita = {}
    for domanda in etichette_dict:
        for categoria in etichette_dict[domanda]:
            for etichetta in etichette_dict[domanda][categoria]:        
                problematicita[etichetta] = {"R": 0}
                risposte = dizionari_risposte[domanda]
                for id_ in risposte:
                    conta = False
                    for autore in risposte[id_]:
                        if etichetta in risposte[id_][autore][categoria]:
                            conta = True
                    if conta:
                        problematicita[etichetta]["R"] += 1 
    for etichetta, indici in problematicita.items():
        problematicita[etichetta]["PR_L"] = 0
        for domanda in dizionari_risposte:
            for i in dizionari_risposte[domanda]:
                C_i = 0
                for autore in dizionari_risposte[domanda][i]:
                    for categoria in dizionari_risposte[domanda][i][autore]:
                        if etichetta in dizionari_risposte[domanda][i][autore][categoria]:
                            C_i += 1
                if C_i > 0:
                    R_i = NUMAUTORI - C_i
                    problematicita[etichetta]["PR_L"] += R_i
        problematicita[etichetta]["PR_L"] /= indici["R"]
    return problematicita

etichette_problematicita = calcola_problematicita(dizionari_risposte)

# Stampa della problematicità delle etichette in ordine decrescente 
stampa_su_file_ON("problematicita_etichette.txt")
print("\n\n========== Problematicità: Etichetta, R, PR_L in ordine decrescente di probematicità")
for etichetta, values in sorted(etichette_problematicita.items(), key=lambda x: x[1]["PR_L"], reverse=True):
    print(f"{etichetta:<60} {values['R']:>5} {values['PR_L']:>10.2f}")
stampa_su_file_OFF(output_file)

# Stampa delle problematicità per ogni risposta separatamente
stampa_su_file_ON("problematicita_risposte.txt")
# Problematicità per ogni risposta separatamente per d1, d2, d3
for domanda in dizionari_risposte.keys():  # Iterate over d1, d2, d3
    print(f"\n\n========== Problematicità per {domanda.upper()}")
    for T in ids:
        Q = set()
        for autore in autori_file.keys():
            for macroetichetta, etichette in dizionari_risposte[domanda][T][autore].items():
                Q.update(etichette)

        Q = list(Q)
        problematicita_T = 0

        for etichetta_i in Q:
            C_i = sum(
                1 for autore in autori_file.keys()
                if any(etichetta_i in dizionari_risposte[domanda][T][autore][macroetichetta]
                       for macroetichetta in dizionari_risposte[domanda][T][autore])
            )
            T_i = NUMAUTORI - C_i
            problematicita_T += T_i

        PR_T = problematicita_T / len(Q) if Q else None
       
        if (PR_T != 0):
            print(f"Risposta {T}: Problematicità {PR_T:.2f}")
stampa_su_file_OFF(output_file)
