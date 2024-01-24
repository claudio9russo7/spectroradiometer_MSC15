
import xlsxwriter
import pandas as pd
import glob
import os


def elaboratore_multispettrale(percorso_raw_data, percorso_elab_data):
    nomi = []
    for file in os.listdir(percorso_raw_data):
        nomi.append(int(file[9:-7]))

    percorsi = glob.glob(f"{percorso_raw_data}/*.xls")
    results = []
# Oltre al traspose, con .iloc() mi prendo la riga 0 e rinomino le colonne, dopodiché
# prendo la riga uno dove sono i valori che mi interessano con df[1:]
    for path in percorsi:
        df = pd.read_excel(path)
        df = df.transpose()
        df.columns = df.iloc[0]
        df = df[1:]
        results.append(df)

    name_results = dict(sorted(zip(nomi, results)))
    final = {str(key): value for key, value in name_results.items()}


    for key, dataframe in final.items():
        dataframe["Nome File"] = key
        dataframe["Red"] = dataframe.iloc[0, 306:417].sum()
        dataframe["Far Red"] = dataframe.iloc[0, 416:528].sum()
        dataframe["Red:Far Red"] = dataframe["Red"] / dataframe["Far Red"]
        dataframe.insert(0, "Nome File", dataframe.pop("Nome File"))
        dataframe.insert(9, "Red", dataframe.pop("Red"))
        dataframe.insert(10, "Far Red", dataframe.pop("Far Red"))
        dataframe.insert(11, "Red:Far Red", dataframe.pop("Red:Far Red"))

    finalissimo_lista = list(final.values())
    lista_di_dataframe = []
    for element in finalissimo_lista:
        df = pd.DataFrame(element)
        lista_di_dataframe.append(df)

    columns = lista_di_dataframe[0].columns
    lenght = len(lista_di_dataframe[0].columns)

    for e in lista_di_dataframe:
        e.columns = range(lenght)

    nome_file = input("Come vuoi chiamare il file finale?\n")
    df_finale = pd.concat(lista_di_dataframe)
    df_finale.columns = columns
    with pd.ExcelWriter(f"{percorso_elab_data}/{nome_file}.xlsx", engine='xlsxwriter',
                        engine_kwargs={'options': {'strings_to_numbers': True}}) as writer:
        df_finale.to_excel(writer, sheet_name="All index")
        for keys, dataframe in final.items():
            dataframe.columns = columns
            dataframe.to_excel(writer, sheet_name=keys)

    print("Ho finito di elaborare!")


