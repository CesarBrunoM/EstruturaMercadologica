from tkinter import *
from tkinter import messagebox, scrolledtext
from df_excel import ArquivoExcel
from conexaoBD import ConexaoBanco
import pandas as pd


def conexao():
    banco = text_nomebanco.get()
    server = text_nomeserver.get()
    try:
        bdconexao = ConexaoBanco(server, banco)
        return bdconexao
    except:
        messagebox.showerror(title="Erro de conexão", message="Falha na conexão com o banco de dados")


def dataframeexcel(tabela):
    arquivo = text_caminhoarquivo.get()
    dataframe = pd.read_excel(r'{}'.format(arquivo), sheet_name=f"{tabela}")
    return dataframe


def btntestarconexao():
    '''Criar o cursor para acesso ao banco de dados utilizando os parametros passados pelo usuario.'''
    try:
        conexao().testarconexao()
    except AttributeError:
        messagebox.showerror(title="Falha",
                             message="Não foi possivel realizar o teste.\nRevise os dados de conexão.")


def btnbuscaarquivo():
    text_caminhoarquivo.configure(state='normal')
    text_caminhoarquivo.delete(0, END)
    ArquivoExcel().buscararquivo(text_caminhoarquivo)
    text_caminhoarquivo.configure(state='disabled')


def btndeletarestrutura():
    text_status.configure(state='normal')
    try:
        conexao().deletarestrutura(text_status)
        text_status.insert("1.0",
                           "Processo de exclusão concluido.\n==========================================================\n")
    except:
        messagebox.showerror(title="Falha ao deletar os dados",
                             message="Não foi possivel encontrar os dados para exclusão.\nVerifique seu acesso ao banco.")
    text_status.configure(state='disabled')


def btninserirdepto():
    try:
        dataframe = dataframeexcel(tabela='DEPARTAMENTOS')
        text_status.configure(state='normal')
    except FileNotFoundError:
        messagebox.showerror(title='Falha ao ler arquivo', message='Arquivo excel não foi encontrado.')
    else:
        try:
            conexao().insertdepto(depto_df=dataframe, codloja=1, textstatus=text_status)
        except:
            messagebox.showerror(title="Falha de conexão",
                                 message="Verifique seu acesso ao banco de dados.")
    text_status.configure(state='disabled')


def btninserirgrupo():
    try:
        dataframe = dataframeexcel(tabela='GRUPOS')
        text_status.configure(state='normal')
    except FileNotFoundError:
        messagebox.showerror(title='Falha ao ler arquivo', message='Arquivo excel não foi encontrado.')
    else:
        try:
            conexao().insertgrupo(grupo_df=dataframe, codloja=1, textstatus=text_status)
        except:
            messagebox.showerror(title="Falha de conexão", message="Verifique seu acesso ao banco de dados.")
    text_status.configure(state='disabled')


def btninserirsubg():
    try:
        dataframe = dataframeexcel(tabela='SUB_GRUPOS')
        text_status.configure(state='normal')
    except FileNotFoundError:
        messagebox.showerror(title='Falha ao ler arquivo', message='Arquivo excel não foi encontrado.')
    else:
        try:
            conexao().insertsubg(subg_df=dataframe, codloja=1, textstatus=text_status)
        except:
            messagebox.showerror(title="Falha de conexão", message="Verifique seu acesso ao banco de dados.")
    text_status.configure(state='disabled')


def ajusteproduto():
    try:
        df_produto = dataframeexcel(tabela='BASE_PRODUTO')
        text_status.configure(state='normal')
    except FileNotFoundError:
        messagebox.showerror(title='Falha ao ler arquivo', message='Arquivo excel não foi encontrado.')
    else:
        try:
            conexao().ajustproduto(df_produto, text_status)
        except:
            messagebox.showerror(title="Falha de conexão", message="Verifique seu acesso ao banco de dados.")
    text_status.configure(state='disabled')


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
    highlightthickness=0,
    state='disabled')

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
    command=btntestarconexao,
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
    command=btndeletarestrutura,
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
    command=btninserirdepto,
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
    command=btninserirgrupo,
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
    command=btninserirsubg,
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
    command=ajusteproduto,
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
    command=btnbuscaarquivo,
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
    highlightthickness=0,
    state='disabled')

text_status.place(
    x=30, y=458,
    width=570,
    height=74)

window.resizable(False, False)
window.mainloop()
