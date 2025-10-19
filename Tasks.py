import json
import os
import threading
import queue
import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, Playwright, TimeoutError, expect
import warnings
import pyxlsb
import csv

import time
from datetime import date, timedelta

caminho_base = os.getcwd()


def Processar_Demandas(caminho_base=caminho_base):

    caminho_pasta = os.path.join(caminho_base,"Demanda")
    caminho_df_fornecedor = os.path.join(caminho_base,"Bases","DB Fornecedores.xlsx")
    df_DB_fornecedor = pd.read_excel(caminho_df_fornecedor)
    df_DB_fornecedor=df_DB_fornecedor[["CODIMS","CODSAP","UF","FANTAS"]]

    # Verifica se a pasta de demandas existe
    if not os.path.isdir(caminho_pasta):
        print(f"Aviso: A pasta '{caminho_pasta}' não foi encontrada.")
        return pd.DataFrame()

    # Lista para armazenar os DataFrames de cada arquivo processado
    lista_dfs = []
    df_temp = pd.DataFrame()

    # Percorre todos os arquivos na pasta de demandas
    for nome_arquivo in os.listdir(caminho_pasta):
        caminho_completo_arquivo = os.path.join(caminho_pasta, nome_arquivo)
        nome_arquivo_lower = nome_arquivo.lower()
        
        try:
            # --- MANTÉM A LÓGICA ORIGINAL PARA ARQUIVOS .TXT E .CSV ---
            if nome_arquivo_lower.endswith((".txt", ".csv")):
                dados_arquivo_atual = []
                with open(caminho_completo_arquivo, "r", encoding="utf-8", errors="ignore") as arquivo:
                    linhas_a_processar = arquivo.readlines()

                # Processa cada linha extraída do arquivo de texto
                for linha in linhas_a_processar:
                    if "AUTOMATIC" in linha:
                        continue

                    linha = linha.strip()

                    # A lógica de fatiamento requer um comprimento mínimo
                    if len(linha) >= 20:
                        try:
                            # Extrai os dados com base na posição dos caracteres
                            PN = linha[3:14]
                            SAP = linha[-20:-11]
                            quantidade = linha[-11:].replace("+", "")

                            # Adiciona os dados extraídos à lista deste arquivo
                            dados_arquivo_atual.append({
                                "PN": int(PN.strip()),
                                "SAP": int(SAP.strip()),
                                "QUANT": int(quantidade.strip()),
                            })
                        except (ValueError, IndexError):
                            # Ignora linhas que não seguem o formato esperado
                            continue

                # Se dados foram extraídos do arquivo, cria um DataFrame
                if dados_arquivo_atual:
                    df_temp = pd.DataFrame(dados_arquivo_atual)
                    lista_dfs.append(df_temp)


            # --- NOVA LÓGICA PARA PROCESSAR ARQUIVOS EXCEL (.XLS, .XLSX) ---
            elif nome_arquivo_lower.endswith((".xls", ".xlsx")):
                
                # Mapeamento dos nomes de coluna do arquivo Excel para os nomes desejados
                colunas_mapeamento = {
                    'DESENHO': 'PN',
                    'COD ORIGEM': 'SAP',
                    'ENTREGA SOLICITADA': 'QUANT'
                }

                # Lê o arquivo Excel
                df_excel = pd.read_excel(caminho_completo_arquivo)
                

                # Pega a lista de colunas que precisamos do arquivo original
                colunas_originais_necessarias = list(colunas_mapeamento.keys())

                # Verifica se todas as colunas necessárias existem no arquivo
                if not all(coluna in df_excel.columns for coluna in colunas_originais_necessarias):
                    print(f"Aviso: O arquivo '{nome_arquivo}' não contém todas as colunas necessárias e será ignorado.")
                    continue

                # 1. Seleciona apenas as colunas que nos interessam
                df_temp = df_excel[colunas_originais_necessarias].copy()
                
                # 2. Renomeia as colunas para o padrão final
                df_temp.rename(columns=colunas_mapeamento, inplace=True)
                df_temp = df_temp[df_temp["QUANT"] > 0]

                # 3. Adiciona o DataFrame processado à lista para concatenação posterior
                lista_dfs.append(df_temp)

        except Exception as e:
            print(f"Erro ao processar o arquivo '{nome_arquivo}': {e}")
            continue

    if not lista_dfs:
        print("Nenhum dado válido foi processado.")
        return pd.DataFrame()
    
    # Concatena todos os DataFrames da lista em um único DataFrame final
    df_final = pd.concat(lista_dfs, ignore_index=True)
    
    colunas_numericas = ["PN", "SAP", "QUANT"]
    for col in colunas_numericas:
        df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

    # Remove linhas onde a conversão numérica falhou (resultando em NaT/NaN)
    df_final.dropna(subset=colunas_numericas, inplace=True)

    df_unique_fornecedor = df_DB_fornecedor.drop_duplicates(subset=["CODSAP"], keep="first")

    # Convert the unique CODSAP column to integer and set as index
    codsap_map = df_unique_fornecedor.set_index("CODSAP")["FANTAS"]
    codsap_map_estado = df_unique_fornecedor.set_index("CODSAP")["UF"]

    # Map to df_final
    df_final["FORNECEDOR"] = df_final["SAP"].astype('Int64', errors='ignore').map(codsap_map)
    df_final["ESTADO"] = df_final["SAP"].astype('Int64', errors='ignore').map(codsap_map_estado)
    
    # Converte colunas para inteiro após remover os nulos
    for col in colunas_numericas:
        df_final[col] = df_final[col].astype(int)

    return df_final.to_excel("Resultados/Demandas_Total.xlsx",index=False)



def PROG_CARGO_LIFT():
    print("print cargo lift")




def PROG_CKD_MILK_RUN():
    print("PROG_CKD_MILK_RUN")

