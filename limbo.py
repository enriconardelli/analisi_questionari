import pprint
import sys

def count_etichette_advanced(data, etichette_count):
    data_with_count = {}
    for level, level_data in data.items():  # Iterate over d1, d2, d3
        data_with_count[level] = {}
        for id_, autori in level_data.items():
            data_with_count[level][id_] = {}
            for autore, macroetichetta_dict in autori.items():
                data_with_count[level][id_][autore] = {}
                for macroetichetta, etichette in macroetichetta_dict.items():
                    if isinstance(etichette, set):
                        data_with_count[level][id_][autore][macroetichetta] = {
                            etichetta: etichette_count[etichetta] for etichetta in etichette
                        }
    return data_with_count

# Get the count of all etichette as a dictionary
etichette_count_advanced = count_etichette_advanced(dizionari_risposte, etichette_count)
# Print the resulting dictionary for verification
output_file = stampa_su_file_ON("etichette_count_advanced.txt")
pprint.pprint(etichette_count_advanced)
stampa_su_file_OFF(output_file)
