from tkinter import filedialog


class ArquivoExcel(object):

    def buscararquivo(self, textdir):
        '''Busca o arquivo excel no sistema.
        filetypes=(("Arquivo csv", ".csv"), ("Arquivo xlsx", ".xlsx"), ("Arquivo xls", ".xls"))'''
        self.planilha = filedialog.askopenfile(mode='r', title="Selecione o arquivo com a departamentalização",
                                               filetypes=(
                                                   ("Arquivo xlsx", ".xlsx"), ("Arquivo xls", ".xls")))

        caminho = getattr(self.planilha, "name")

        textdir.insert(0, caminho)