from tkinter import *
from tkinter import messagebox, scrolledtext, filedialog
import pyodbc
from df_excel import lerexcel
from conexaoBD import conectar, scriptDepto, scriptGrupo, scriptSubg
import pandas as pd

codloja = 1


def cursor():
    '''Criar o cursor para acesso ao banco de dados utilizando os parametros passados pelo usuario.'''
    nome_servidor = text_nomeserver.get()
    banco_dados = text_nomebanco.get()

    conexao_bd = conectar(nome_servidor, banco_dados)

    conexao = pyodbc.connect(conexao_bd)
    cursor_bd = conexao.cursor()
    return cursor_bd


def consultaqtddados(nometabela):
    '''Realiza uma consulta no banco de dados com o parametro NOMETABELA retornando a quantidade de registros na tabela'''
    # MELHORAR METODO PARA SABER QUANTOS REGISTROS EXISTEM NA TABELA COM PANDAS
    cursorbanco = cursor()
    cursorbanco.execute(f'SELECT * FROM {nometabela}')
    dados = cursorbanco.fetchall()
    dados = pd.DataFrame(dados)
    qtd = len(dados.index)
    return qtd


def deletartabela(nometabela):
    '''Deleta dados da tabela passada pelo parametro NOMETABELA.'''
    cursorbanco = cursor()

    cursorbanco.execute(f'DELETE FROM DBO.{nometabela}')
    cursorbanco.commit()


def testarconexao():
    '''Testa a conexão do banco de dados utilizando a função CURSOR()'''
    try:
        cursorbanco = cursor()

        cursorbanco.execute("SELECT DB_NAME();")
        bd = cursorbanco.fetchall()
        bdtupla = bd[0][0]
        messagebox.showinfo(title="Teste conexão", message=f"Conectado ao banco de dados {bdtupla}.")
    except:
        messagebox.showerror(title="Teste conexão", message="Falha na conexão com o banco de dados.")


def buscararquivo():
    '''Busca o arquivo excel no sistema.
    filetypes=(("Arquivo csv", ".csv"), ("Arquivo xlsx", ".xlsx"), ("Arquivo xls", ".xls"))'''
    planilha = filedialog.askopenfile(mode='r', title="Selecione o arquivo com a departamentalização",
                                      filetypes=(
                                          ("Arquivo xlsx", ".xlsx"), ("Arquivo xls", ".xls")))

    caminho = getattr(planilha, "name")
    text_caminhoarquivo.delete(0, END)
    text_caminhoarquivo.insert(0, caminho)


def deletarestrutura():
    '''Deleta as tabelas de departamentos, grupos e subgrupos do banco de dados'''
    nometabela = ['SUB_GRUPOS', 'GRUPOS', 'DEPTO']

    try:
        for nome in nometabela:
            qtd = consultaqtddados(nome)
            if qtd > 0:
                deletartabela(nome)
                messagebox.showinfo(title="Sucesso",
                                    message=f"Dados da tabela {nome} deletados. \n{qtd} registros deletados.")
                text_status.insert("1.0", f"Foram deletados {qtd} registros da tabela {nome}.\n")
            else:
                messagebox.showwarning(title="Aviso", message=f"Não existe dados na tabela {nome}.")
    except:
        messagebox.showerror(title="Erro na conexão",
                             message="Não foi possivel acessar o banco de dados, verifique as informações de conexão.")


def insertdepto():
    '''Realiza a leitura do arquivo excel buscando a aba DEPTO,
    apos faz o tratamento dos dados e insere as informações no banco de dados.'''
    count = 0
    cursorbanco = cursor()
    consulta = consultaqtddados('DEPTO')
    if consulta <= 0:
        try:
            arquivo = text_caminhoarquivo.get()
            depto_df = lerexcel(arquivo, 'DEPARTAMENTOS')

            for i, codigo in enumerate(depto_df['CODIGO']):
                depto = depto_df.loc[i, 'DEPARTAMENTO'].replace("'", "").strip().upper()

                df_dados = str(codigo) + ', \'' + depto + '\'' + ',' + ' GETDATE() ' + f' ,{codloja} ' + ' ,' + str(
                    codigo) + ')'
                query = scriptDepto + df_dados
                count += 1
                cursorbanco.execute(query)
                cursorbanco.commit()
                text_status.insert("1.0", f"Departamento {depto} inserido com sucesso.\n")

            text_status.insert("1.0",
                               f"Foram inseridos {count} departamentos.\n==========================================================\n")
        except FileNotFoundError:
            messagebox.showerror(title="Falta de dados para o comando",
                                 message="Validar conexão e arquivo selecionados.")
    else:
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

            datagrup = str(codigo) + ', \'' + grupo + '\', ' + str(coddep) + ', GETDATE(), ' + f' {codloja}, ' + str(
                codigo) + ')'
            query = scriptGrupo + datagrup
            count += 1
            cursorbanco.execute(query)
            cursorbanco.commit()
            text_status.insert("1.0", f"Grupo {grupo} inserido com sucesso.\n")

        text_status.insert("1.0",
                           f"Foram inseridos {count} grupos.\n==========================================================\n")
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

            text_status.insert("1.0", f"SubGrupo {subgrupo} inserindo com sucesso.\n")

        text_status.insert("1.0",
                           f"Foram inseridos {count} subgrupos.\n==========================================================\n")
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

            text_status.insert('1.0', f'Produto {produto} alterado para subgrupo de codigo {cod_subg}.\n')
        text_status.insert('1.0',
                           f'Foram alterados {count} produtos.\n==========================================================\n')
    except FileNotFoundError:
        messagebox.showerror(title="Falta de dados para o comando",
                             message="Validar conexão e arquivo selecionados.")


window = Tk()

window.geometry("637x573")
window.configure(bg="#ffffff")
canvas = Canvas(
    window,
    bg="#ffffff",
    height=573,
    width=637,
    bd=0,
    highlightthickness=0,
    relief="ridge")
canvas.place(x=0, y=0)

background_img = PhotoImage(file=f"imagens/background.png")
background = canvas.create_image(
    318.5, 286.5,
    image=background_img)

entry0_img = PhotoImage(file=f"imagens/img_textBox0.png")
entry0_bg = canvas.create_image(
    315.0, 221.5,
    image=entry0_img)

text_caminhoarquivo = Entry(
    bd=0,
    bg="#ffffff",
    highlightthickness=0)

text_caminhoarquivo.place(
    x=30, y=209,
    width=570,
    height=23)

entry1_img = PhotoImage(file=f"imagens/img_textBox1.png")
entry1_bg = canvas.create_image(
    411.5, 90.5,
    image=entry1_img)

text_nomebanco = Entry(
    bd=0,
    bg="#ffffff",
    highlightthickness=0)

text_nomebanco.place(
    x=223, y=78,
    width=377,
    height=23)

entry2_img = PhotoImage(file=f"imagens/img_textBox2.png")
entry2_bg = canvas.create_image(
    411.5, 57.5,
    image=entry2_img)

text_nomeserver = Entry(
    bd=0,
    bg="#ffffff",
    highlightthickness=0)

text_nomeserver.place(
    x=223, y=45,
    width=377,
    height=23)

img0 = PhotoImage(file=f"imagens/img0.png")
btn_teste = Button(
    image=img0,
    borderwidth=0,
    highlightthickness=0,
    command=testarconexao,
    relief="flat")

btn_teste.place(
    x=463, y=115,
    width=137,
    height=28)

img1 = PhotoImage(file=f"imagens/img1.png")
btn_deletar = Button(
    image=img1,
    borderwidth=0,
    highlightthickness=0,
    command=deletarestrutura,
    relief="flat")

btn_deletar.place(
    x=31, y=256,
    width=179,
    height=25)

img2 = PhotoImage(file=f"imagens/img2.png")
btn_insertdepto = Button(
    image=img2,
    borderwidth=0,
    highlightthickness=0,
    command=insertdepto,
    relief="flat")

btn_insertdepto.place(
    x=31, y=287,
    width=179,
    height=25)

img3 = PhotoImage(file=f"imagens/img3.png")
btn_insertgrup = Button(
    image=img3,
    borderwidth=0,
    highlightthickness=0,
    command=insertgrupo,
    relief="flat")

btn_insertgrup.place(
    x=31, y=318,
    width=179,
    height=25)

img4 = PhotoImage(file=f"imagens/img4.png")
btn_insertsubg = Button(
    image=img4,
    borderwidth=0,
    highlightthickness=0,
    command=insertsubg,
    relief="flat")

btn_insertsubg.place(
    x=31, y=349,
    width=179,
    height=25)

img5 = PhotoImage(file=f"imagens/img5.png")
btn_ajusteprod = Button(
    image=img5,
    borderwidth=0,
    highlightthickness=0,
    command=ajustproduto,
    relief="flat")

btn_ajusteprod.place(
    x=31, y=380,
    width=179,
    height=25)

img6 = PhotoImage(file=f"imagens/img6.png")
btn_buscaarquivo = Button(
    image=img6,
    borderwidth=0,
    highlightthickness=0,
    command=buscararquivo,
    relief="flat")

btn_buscaarquivo.place(
    x=463, y=242,
    width=137,
    height=28)

entry3_img = PhotoImage(file=f"imagens/img_textBox3.png")
entry3_bg = canvas.create_image(
    315.0, 496.0,
    image=entry3_img)

text_status = scrolledtext.ScrolledText(
    bd=0,
    bg="#ffffff",
    highlightthickness=0)

text_status.place(
    x=30, y=458,
    width=570,
    height=74)

window.resizable(False, False)
window.mainloop()
