import pandas as pd
import random
from datetime import datetime, timedelta

def gerar_base_ficticia():
    nome_arq = str(input('Qual vai ser o nome do arquivo: '))
    nome_arq = nome_arq + '.xlsx'
    qtde_linhas = int(input("Quantas linhas a base deve ter? "))
    qtde_colunas = int(input("Quantas colunas deseja criar? "))

    colunas = {}


    for i in range(qtde_colunas):
        print(f"\n--- Coluna {i+1} ---")
        nome_coluna = input("Digite o nome da coluna: ")

        tipo = input(
            f"Escolha o tipo da coluna '{nome_coluna}':\n"
            "1 = Texto (valores possíveis)\n"
            "2 = Número (range ex: 10-500)\n"
            "3 = Data (range ex: 2022-01-01 ; 2022-12-31)\n"
            "Digite a opção: "
        )

        if tipo == "1":
            valores = input(
                f"Digite os valores possíveis para '{nome_coluna}' separados por vírgula: "
            )
            lista_valores = [v.strip() for v in valores.split(",")]
            colunas[nome_coluna] = ("texto", lista_valores)

        elif tipo == "2":
            faixa = input(
                f"Digite o range numérico para '{nome_coluna}' (ex: 1-1900): "
            )
            minimo, maximo = map(int, faixa.split("-"))
            colunas[nome_coluna] = ("numero", (minimo, maximo))

        elif tipo == "3":
            datas = input(
                f"Digite o range de datas (ex: 2022-01-01 ; 2022-12-31): "
            )
            data_ini, data_fim = [d.strip() for d in datas.split(";")]
            colunas[nome_coluna] = ("data", (data_ini, data_fim))

        else:
            print("Tipo inválido! Pulando coluna.\n")

    
    data = {}

    for coluna, config in colunas.items():
        tipo, valor = config

        if tipo == "texto":
            lista = valor
            data[coluna] = [random.choice(lista) for _ in range(qtde_linhas)]

        elif tipo == "numero":
            minimo, maximo = valor
            data[coluna] = [random.randint(minimo, maximo) for _ in range(qtde_linhas)]

        elif tipo == "data":
            ini, fim = valor
            dt_ini = datetime.strptime(ini, "%Y-%m-%d")
            dt_fim = datetime.strptime(fim, "%Y-%m-%d")
            delta = (dt_fim - dt_ini).days

            data[coluna] = [
                (dt_ini + timedelta(days=random.randint(0, delta))).date()
                for _ in range(qtde_linhas)
            ]

    # Criar DataFrame e salvar
    df = pd.DataFrame(data)
    df.to_excel(nome_arq, index=False)

    print("\nBase fictícia gerada com sucesso! → base_ficticia.xlsx")

# Executar programa
gerar_base_ficticia()
