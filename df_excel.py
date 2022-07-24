import pandas as pd


def lerexcel(arquivo, aba):
    df_excel = pd.read_excel(r'{}'.format(arquivo), sheet_name=f"{aba}")

    return df_excel