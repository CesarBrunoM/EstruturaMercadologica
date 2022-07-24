def conectar(nomeServidor, bancoDados):
    nomeServidor = nomeServidor.strip().upper()
    bancoDados = bancoDados.strip().upper()

    conexao_bd = (
        "Driver={SQL Server};"
        f"Server={nomeServidor};"
        f"Database={bancoDados};"
        "UID=SA;"
        "PWD=usuarioteste;")

    return conexao_bd

scriptDepto = '''Insert into DEPTO (CODIGO, NOME, DATA,CODLOJA, SEQUENCIA) VALUES ('''
scriptGrupo = '''Insert into GRUPOS (CODIGO, NOME, CODDEP, DATA, CODLOJA, SEQUENCIA) VALUES ('''
scriptSubg = '''Insert into SUB_GRUPOS (CODIGO, CODLOJA, NOME, CODGRU, DATA, SEQUENCIA) VALUES ('''