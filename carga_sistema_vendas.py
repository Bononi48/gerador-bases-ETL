import pandas as pd
import pyodbc
from dotenv import load_dotenv
import os

load_dotenv()

DB_DRIVER = os.getenv("DB_DRIVER")
DB_SERVER = os.getenv("DB_SERVER")
DB_DATABASE = os.getenv("DB_DATABASE")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
PATH_XLSX = os.getenv("PATH_XLSX")

# Estabelecendo conexão com BD
dados_conexao = (
    f"Driver={DB_DRIVER};"
    f"Server={DB_SERVER};"
    f"Database={DB_DATABASE};"
    f"UID={DB_USER};"
    f"PWD={DB_PASSWORD};"
    "Encrypt=no;"
    "TrustServerCertificate=yes;"
                 )

conexao = pyodbc.connect(dados_conexao)

print('conexão bem sucedida')

cursor = conexao.cursor()

wb = pd.read_excel(rf'{PATH_XLSX}')

print('Inserindo dados na tabela')

for i, linha in enumerate(wb['VendedorID']):
    data = {
        'VendedorID': int(wb.loc[i, 'VendedorID']),
        'DataID': wb.loc[i, 'DataID'],
        'ProdutoID': wb.loc[i, 'ProdutoID'],
        'ClientID': wb.loc[i, 'ClientID'],
        'Qtde Vendas': wb.loc[i, 'Qtde Vendas'],
        'valor venda': str(wb.loc[i, 'valor venda']),
        'desconto': wb.loc[i, 'desconto']
    }

    for key, value in data.items():
        if pd.isna(value):
            data[key] = None

    valores = []

    for v in data.values():
        if hasattr(v, "item"):
            v = v.item()

        valores.append(v)

    comando = """INSERT INTO [dbo].[Fato_Vendas]
                ([VendedorID]
                ,[DataID]
                ,[ProdutoID]
                ,[ClienteID]
                ,[Quantidade_Venda]
                ,[ValorTotal]
                ,[Desconto])
                VALUES (?, ?, ?, ?, ?, ?, ?)"""

    cursor.execute(comando, tuple(valores))
    cursor.commit()

print('Processo de carga concluído')
