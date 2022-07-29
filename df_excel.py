import pandas as pd
from tkinter import filedialog


class ArquivoExcel(object):

    def __init__(self, text_caminhoarquivo):
        '''Busca o arquivo excel no sistema.
        filetypes=(("Arquivo csv", ".csv"), ("Arquivo xlsx", ".xlsx"), ("Arquivo xls", ".xls"))'''
        self.planilha = filedialog.askopenfile(mode='r', title="Selecione o arquivo com a departamentalização",
                                               filetypes=(
                                                   ("Arquivo xlsx", ".xlsx"), ("Arquivo xls", ".xls")))

        caminho = getattr(self.planilha, "name")

        text_caminhoarquivo.insert(0, caminho)

    def lerexcel(self, arquivo, aba):
        df_excel = pd.read_excel(r'{}'.format(arquivo), sheet_name=f"{aba}")

        return df_excel
