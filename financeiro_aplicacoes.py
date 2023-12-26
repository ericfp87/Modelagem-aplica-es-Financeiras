# -*- coding: utf-8 -*-

import pandas as pd
import os
import shutil
import xlrd

meses = [
    'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
    'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
]

diretorio_origem = r'G:\Drives compartilhados\Financeiro\01 - Pagamentos\2023'
diretorio_destino = r'C:\Users\Bulbe\OneDrive - Bulbe\Power Automate\Financeiro - Aplicacoes\arquivos'

for mes in range(1, 13):
    nome_mes = f"{mes:02d} - {meses[mes-1].upper()}"
    diretorio_mes = os.path.join(diretorio_origem, nome_mes, 'BULBE', '01 - Extratos Bancários')

    if os.path.exists(diretorio_mes):
        for arquivo in os.listdir(diretorio_mes):
            if arquivo.startswith('BULBE APLIC') and arquivo.endswith('.xls'):
                caminho_origem = os.path.join(diretorio_mes, arquivo)
                caminho_destino = os.path.join(diretorio_destino, arquivo)
                shutil.copy2(caminho_origem, caminho_destino)








diretorio = r'G:\Drives compartilhados\Analise de Dados\02 - Projetos\01 - BI\Financeiro Aplicações\arquivos'

# Lista todos os arquivos .xls no diretório
arquivos_xls = [arquivo for arquivo in os.listdir(diretorio) if arquivo.endswith('.xls')]

# # Loop para converter cada arquivo .xls em .xlsx
# for arquivo in arquivos_xls:
#     caminho_arquivo = os.path.join(diretorio, arquivo)
#     planilha = pd.read_excel(caminho_arquivo)  # Lê o arquivo .xls usando o pandas
#     novo_nome = os.path.splitext(arquivo)[0] + '.xlsx'  # Novo nome do arquivo com extensão .xlsx
#     novo_caminho = os.path.join(diretorio, novo_nome)
#     planilha.to_excel(novo_caminho, index=False)  # Salva a planilha em formato .xlsx

#     # Opcional: Remover o arquivo .xls original
#     # os.remove(caminho_arquivo)

# print('Conversão concluída.')








# # Definir o diretório dos arquivos .xls
# diretorio = r"G:\Drives compartilhados\Analise de Dados\02 - Projetos\01 - BI\Financeiro Aplicações\arquivos"

# # Obter a lista de arquivos .xls no diretório
# arquivos = [arq for arq in os.listdir(diretorio) if arq.endswith(".xls")]

# # Cabeçalho do arquivo .csv
# cabecalho = [
#     "Arquivo",
#     "Saldo anterior",
#     "Aplicações",
#     "Resgates/Recompras",
#     "Vencimentos",
#     "Rendim. acumulado",
#     "Saldo bruto final (**)","Impostos estimados",
#     "Saldo final (**)"]

# # Criar um DataFrame vazio
# df_final = pd.DataFrame(columns=cabecalho)

# # Loop para processar cada arquivo .xls
# for arquivo in arquivos:
#     # Caminho completo do arquivo
#     caminho_arquivo = os.path.join(diretorio, arquivo)
    
#     # Ler o arquivo .xls e extrair os dados usando a biblioteca xlrd
#     xls_workbook = xlrd.open_workbook(caminho_arquivo)
#     xls_sheet = xls_workbook.sheet_by_index(0)
#     valores = xls_sheet.row_values(1)
    
#     # Adicionar o nome do arquivo aos valores extraídos
#     valores.insert(0, arquivo)
    
#     # Adicionar os valores extraídos ao DataFrame final
#     df_final.loc[len(df_final)] = valores

# # Formatar os valores no DataFrame final
# df_final = df_final.applymap(lambda x: f"{x:,.2f}".replace(",", "_").replace(".", ",").replace("_", "."))

# # Salvar o DataFrame como arquivo .csv
# caminho_csv = os.path.join(diretorio, "dados.csv")
# df_final.to_csv(caminho_csv, sep=";", index=False)

# print("Arquivo CSV gerado com sucesso!")