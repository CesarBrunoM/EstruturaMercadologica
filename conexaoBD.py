import pandas as pd
import pyodbc
from tkinter import messagebox


class ConexaoBanco(object):
    _banco = None

    def __init__(self, nomeserver, bancodados):
        nomeserver = nomeserver.strip()
        bancodados = bancodados.strip()

        self._banco = pyodbc.connect(
            "Driver={SQL Server};"
            f"Server={nomeserver};"
            f"Database={bancodados};"
            "UID=SA;"
            "PWD=usuarioteste;")

    def manipularbd(self, scriptsql):
        try:
            cursor = self._banco.cursor()
            cursor.execute(scriptsql)
            cursor.close()
            self._banco.commit()
        except ConnectionError:
            return False
        return True

    def consultarbd(self, scriptsql):
        dados = None
        try:
            cursor = self._banco.cursor()
            cursor.execute(scriptsql)
            dados = cursor.fetchall()
        except ConnectionError:
            return None
        return dados

    def fecharbd(self):
        self._banco.close()

    def consultaqtddados(self, nometabela):
        '''Realiza uma consulta no banco de dados com o parametro NOMETABELA retornando a
        quantidade de registros na tabela'''
        script = f'SELECT * FROM {nometabela}'
        dados = self.consultarbd(script)
        dados = pd.DataFrame(dados)
        qtd = len(dados.index)
        return qtd

    def testarconexao(self):
        '''Testa a conexão do banco de dados utilizando a função CURSOR()'''
        script = "SELECT DB_NAME();"
        bd = self.consultarbd(script)
        bdtupla = bd[0][0]
        messagebox.showinfo(title="Teste conexão", message=f"Conectado ao banco de dados {bdtupla}.")

    def deletarestrutura(self, textstatus):
        '''Deleta as tabelas de departamentos, grupos e subgrupos do banco de dados'''
        nometabela = ['SUB_GRUPOS', 'GRUPOS', 'DEPTO']
        script = f'DELETE FROM '
        for nome in nometabela:
            qtd = self.consultaqtddados(nome)
            if qtd > 0:
                self.manipularbd(script + nome)
                messagebox.showinfo(title="Sucesso",
                                    message=f"Dados da tabela {nome} deletados. \n{qtd} registros deletados.")
                textstatus.insert("1.0",
                                  f'foram deletados {qtd} registros da tabela {nome}.'
                                  f'\n==========================================================\n')
            else:
                messagebox.showwarning(title="Aviso", message=f"Não existe dados na tabela {nome}.")
                textstatus.insert("1.0",
                                  f'Nenhuma exclusão necessária na tabela {nome}.'
                                  f'\n==========================================================\n')

    def insertdepto(self, depto_df, codloja, textstatus):
        '''Realiza a leitura do arquivo excel buscando a aba DEPTO,
        apos faz o tratamento dos dados e insere as informações no banco de dados.'''
        count = 0
        scriptdepto = '''Insert into DEPTO (CODIGO, NOME, DATA,CODLOJA, SEQUENCIA) VALUES ('''
        consulta = self.consultaqtddados('DEPTO')

        if consulta == 0:
            for i, codigo in enumerate(depto_df['CODIGO']):
                depto = depto_df.loc[i, 'DEPARTAMENTO'].replace("'", "").strip().upper()

                df_dados = str(
                    codigo) + ', \'' + depto + '\'' + ',' + ' GETDATE() ' + f' ,{codloja} ' + ' ,' + str(
                    codigo) + ')'
                script = scriptdepto + df_dados
                count += 1
                self.manipularbd(script)
                textstatus.insert("1.0", f"Departamento {depto} inserido com sucesso.\n")

            textstatus.insert("1.0",
                              f"Foram inseridos {count} departamentos."
                              f"\n==========================================================\n")
        elif consulta > 0:
            messagebox.showwarning(title="Aviso",
                                   message=f"Necessário excluir os dados antes da inserção de valores.")

    def insertgrupo():
        '''Realiza a leitura do arquivo excel buscando a aba GRUPO,
        apos faz o tratamento dos dados e insere as informações no banco de dados.'''
        count = 0
        cursorbanco = cursor()

        try:
            arquivo = text_caminhoarquivo.get()
            grupo_df = lerexcel(arquivo, 'GRUPOS')

            for i, codigo in enumerate(grupo_df['CODIGO']):
                grupo = grupo_df.loc[i, 'GRUPO'].replace("'", "").strip().upper()
                coddep = grupo_df.loc[i, 'COD_DEPARTAMENTO']

                datagrup = str(codigo) + ', \'' + grupo + '\', ' + str(
                    coddep) + ', GETDATE(), ' + f' {codloja}, ' + str(
                    codigo) + ')'
                query = scriptGrupo + datagrup
                count += 1
                cursorbanco.execute(query)
                cursorbanco.commit()
                text_status.configure(state='normal')
                text_status.insert("1.0", f"Grupo {grupo} inserido com sucesso.\n")

            text_status.insert("1.0",
                               f"Foram inseridos {count} grupos."
                               f"\n==========================================================\n")
            text_status.configure(state='disabled')

        except FileNotFoundError:
            messagebox.showerror(title="Falta de dados para o comando",
                                 message="Validar conexão e arquivo selecionados.")

    def insertsubg():
        '''Realiza a leitura do arquivo excel buscando a aba SUBG,
        apos faz o tratamento dos dados e insere as informações no banco de dados.'''
        count = 0
        cursorbanco = cursor()

        try:
            arquivo = text_caminhoarquivo.get()
            subg_df = lerexcel(arquivo, 'SUB_GRUPOS')

            for i, codigo in enumerate(subg_df['CODIGO']):
                subgrupo = subg_df.loc[i, 'SUBGRUPO'].replace("'", "").strip().upper()
                codgrup = subg_df.loc[i, 'COD_GRUPO']

                datasubg = str(codigo) + f', {codloja}, ' + '\'' + subgrupo + '\', ' + str(
                    codgrup) + ', ' + ' GETDATE(), ' + str(
                    codigo) + ')'

                query = scriptSubg + datasubg
                count += 1
                cursorbanco.execute(query)
                cursorbanco.commit()
                text_status.configure(state='normal')
                text_status.insert("1.0", f"SubGrupo {subgrupo} inserindo com sucesso.\n")

            text_status.insert("1.0",
                               f"Foram inseridos {count} subgrupos.\n==========================================================\n")
            text_status.configure(state='disabled')

        except FileNotFoundError:
            messagebox.showerror(title="Falta de dados para o comando",
                                 message="Validar conexão e arquivo selecionados.")

    def ajustproduto():
        count = 0
        cursorbanco = cursor()
        try:
            arquivo = text_caminhoarquivo.get()
            produtos_df = lerexcel(arquivo, 'BASE_PRODUTO')

            for i, codigo in enumerate(produtos_df['CODIGO']):
                cod_subg = produtos_df.loc[i, 'COD_SUBG']
                produto = produtos_df.loc[i, 'PRODUTO']
                script = f'UPDATE PRODUTOS SET SUBG = {cod_subg} WHERE CODIGO = {codigo}'
                count += 1
                cursorbanco.execute(script)
                cursorbanco.commit()
                text_status.configure(state='normal')
                text_status.insert('1.0', f'Produto {produto} alterado para subgrupo de codigo {cod_subg}.\n')
            text_status.insert('1.0',
                               f'Foram alterados {count} produtos.\n==========================================================\n')
            text_status.configure(state='disabled')

        except FileNotFoundError:
            messagebox.showerror(title="Falta de dados para o comando",
                                 message="Validar conexão e arquivo selecionados.")


scriptGrupo = '''Insert into GRUPOS (CODIGO, NOME, CODDEP, DATA, CODLOJA, SEQUENCIA) VALUES ('''
scriptSubg = '''Insert into SUB_GRUPOS (CODIGO, CODLOJA, NOME, CODGRU, DATA, SEQUENCIA) VALUES ('''
