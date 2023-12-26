# -*- coding: utf-8 -*-

import pandas as pd

# Nome do arquivo Excel
arquivo_excel = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Financeiro - Aplicacoes\\dados_aplicacoes.xlsx"

# Nome do arquivo CSV de saída
arquivo_csv = "C:\\Users\\Bulbe\\OneDrive - Bulbe\\Power Automate\\Financeiro - Aplicacoes\\dados_aplicacoes_csv.csv"


# Lista para armazenar os dados de todas as planilhas
dados_total = []

# Ler a planilha do Excel
df = pd.read_excel(arquivo_excel, usecols="A:J", header=0, engine='openpyxl')
# Adicionar os dados da planilha à lista total
dados_total.append(df)



# Concatenar os dados de todas as planilhas
dados_concatenados = pd.concat(dados_total)

# Salvar os dados concatenados em um arquivo CSV
dados_concatenados.to_csv(arquivo_csv, sep=";", index=False, encoding="utf-8-sig", decimal=',')