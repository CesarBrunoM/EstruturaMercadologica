import pandas as pd
from tkinter import filedialog


class ArquivoExcel(object):
    def buscararquivo(self, text_caminhoarquivo):
        '''Busca o arquivo excel no sistema.
        filetypes=(("Arquivo csv", ".csv"), ("Arquivo xlsx", ".xlsx"), ("Arquivo xls", ".xls"))'''
        planilha = filedialog.askopenfile(mode='r', title="Selecione o arquivo com a departamentalização",
                                          filetypes=(
                                              ("Arquivo xlsx", ".xlsx"), ("Arquivo xls", ".xls")))

        caminho = getattr(planilha, "name")

        text_caminhoarquivo.insert(0, caminho)

    def lerexcel(self, dirarquivo, aba):
        df_excel = pd.read_excel(r'{}'.format(dirarquivo), sheet_name=f"{aba}")

        return df_excel
