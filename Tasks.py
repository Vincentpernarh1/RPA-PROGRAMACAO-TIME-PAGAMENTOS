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

import xlwings as xw
import time
from datetime import date, timedelta

caminho_base = os.getcwd()

def download_Demanda(page, url_order, q, username, password):
   
    # Processar_Demandas(q)
    try:
        # --- 1. Login and Initial Navigation ---
        q.put(("status", "Navigating to login page..."))
        page.goto(url_order, timeout=60000)
        q.put(("progress", 10))

        q.put(("status", "Performing login..."))
        page.get_by_role("textbox", name="User").fill(username)
        page.get_by_role("textbox", name="Password").fill(password)
        page.get_by_role("button", name="Log In").click()
        q.put(("status", "Login successful!"))
        q.put(("progress", 20))

     
        # --- 2. Navigate to the correct report section ---
        q.put(("status", "Navigating to the report section..."))
        page.locator("#ID_button").click()
        page.get_by_text("ELOG - Importar A8 Automatica").nth(1).click()
        q.put(("progress", 30))

        # --- 3. Date-based Search Loop ---
        current_date = date.today()
        records_found = False
        
        current_date = date.today()
        records_found = False
        
        while not records_found:
            date_str = current_date.strftime('%m/%d/%Y')
            q.put(("status", f"Searching for records on: {date_str}"))
            
            # Fill date fields and click search
            page.get_by_role("textbox", name="Data Inicial").fill(date_str)
            page.get_by_role("textbox", name="Data Final").fill(date_str)
            page.get_by_role("button", name="Pesquisar").click()

            try:
                # If it appears, the code continues. If not, it raises a TimeoutError.
                page.get_by_text("Total records:").click(timeout = 3000)
                
                # This code only runs if the expect() call above succeeds
                q.put(("status", f"Records found for {date_str}!"))
                records_found = True
                q.put(("progress", 45))

            except TimeoutError:
                # This code runs only if the locator was not visible after 60 seconds
                q.put(("status", f"No records found for {date_str}. Trying previous day."))
                current_date -= timedelta(days=1)
                time.sleep(2) # Small delay before trying again

        # --- 4. Row-wise TXT File Download ---
        q.put(("status", "Starting individual file downloads..."))
        
        # Define the base directory for downloads for the found date
        download_path_base = os.path.join(f"Demanda")
        os.makedirs(download_path_base, exist_ok=True)
        
        # This selector targets rows within the table's body to avoid header rows.
        data_rows = page.get_by_role("cell", name="Download")

        row_count = data_rows.count()
        q.put(("status", f"Found {row_count} files to download."))

        for i in range(1,row_count):
            
            row = data_rows.nth(i)
           
            with page.expect_download() as download_info:
               download_link =row.click()
            
            download = download_info.value
            
            # Construct the full path and save the file
            file_path = os.path.join(download_path_base, f"{i}_{download.suggested_filename}")
            
            if os.path.exists(file_path):
                os.remove(file_path)
            download.save_as(file_path)
            q.put(("status", f"Downloaded: {download.suggested_filename}"))

        q.put(("status", "All individual file downloads are complete."))

        q.put(("progress", 65))

        q.put(("status", "Deseja continuar com a transformação das bases?"))

        Processar_Demandas(q)


    except Exception as e:
        q.put(("status", f"An error occurred: {e}"))
        # You might want to add more specific error handling here


def Processar_Demandas(q):
   
    caminho_pasta = os.path.join(caminho_base, "Demanda")
    # output_path = "Resultados/Demandas_Total.xlsx"
    # demand_path = os.path.join(caminho_base,output_path)
    # Atualiza_PFEP(demand_path,q)

    caminho_df_fornecedor = os.path.join(caminho_base, "Bases", "DB Fornecedores.xlsx")
    df_DB_fornecedor = pd.read_excel(caminho_df_fornecedor)
    df_DB_fornecedor = df_DB_fornecedor[["CODIMS", "CODSAP", "UF", "FANTAS"]]

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
                    
                    # --- NOVO: Adiciona o identificador como False para TXT/CSV ---
                    df_temp['__highlight_sap'] = False 
                    
                    lista_dfs.append(df_temp)

            # --- LÓGICA MODIFICADA PARA PROCESSAR ARQUIVOS EXCEL (.XLS, .XLSX) ---
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

                # --- NOVO: Adiciona o identificador para linhas do Excel ---
                df_temp['__highlight_sap'] = True
                # --- FIM DA ADIÇÃO ---

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
    
    df_final['__highlight_sap'] = df_final['__highlight_sap'].fillna(False)

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

    condicao_estado = df_final['ESTADO'] != 'MG'
    condicao_sap = df_final['SAP'] != 800030982
    
    # Aplica AMBAS as condições. O .copy() evita o SettingWithCopyWarning
    df_final = df_final[condicao_estado & condicao_sap].copy()
      
    light_yellow = '#FFFFE0' 
    df_funilaria = pd.DataFrame(columns=['SAP', 'FORNECEDOR']) # Inicializa vazio

    
    if '__highlight_sap' in df_final.columns:
        funilaria_mask = (df_final['__highlight_sap'] == True)
        df_funilaria = df_final[funilaria_mask][['SAP', 'FORNECEDOR']].copy()
        df_funilaria.drop_duplicates(inplace=True)

        DF_Horarios = le_arquivo_horario() # Chama sua função
        
        if not DF_Horarios.empty:
            # 2. Cria o mapa: [Supplier Code] -> [Horário de Janela]
            #    Usa drop_duplicates para garantir que cada Supplier Code seja único
            mapa_horario_pronto = DF_Horarios.drop_duplicates(
                subset=['Supplier Code']
            ).set_index('Supplier Code')['Horário de Janela']

            # 3. Aplica o mapa ao df_funilaria. É muito mais rápido que 'apply'
            df_funilaria['HORÁRIO'] = df_funilaria['SAP'].map(mapa_horario_pronto)
        else:
            print("Aviso: DF_Horarios está vazio. A coluna 'HORÁRIO' não será preenchida.")
            df_funilaria['HORÁRIO'] = pd.NaT # Ou None
        
        # Lógica original de máscara para estilização
        highlight_mask = funilaria_mask.tolist()
        
        # 3. Remove a coluna auxiliar. Ela não é mais necessária.
        df_final = df_final.drop(columns=['__highlight_sap'])
    else:
        # Se a coluna nunca foi criada (só arquivos TXT), cria uma máscara vazia (tudo False)
        highlight_mask = [False] * len(df_final)
        print("Nenhum dado de Excel processado, a aba 'funilaria' estará vazia.")

    sap_styles = [
        f'background-color: {light_yellow}' if mask else ''
        for mask in highlight_mask
    ]

   
    styler = df_final.style

    try:
        styler = styler.apply(lambda _: sap_styles, subset=['SAP'])
    except Exception as e:
        print(f"Não foi possível aplicar o estilo à coluna 'SAP': {e}")
        
    output_path = "Resultados/Demandas_Total.xlsx"
    demand_path = os.path.join(caminho_base,output_path)

    
    try:
        # Cria um ExcelWriter para salvar ambos os DataFrames
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Salva a aba principal estilizada
            styler.to_excel(writer, sheet_name='Demandas_Total', index=False)
            
            # Salva a nova aba 'funilaria' (sem estilo)
            df_funilaria.to_excel(writer, sheet_name='funilaria', index=False)
            
        print(f"Arquivo salvo com sucesso com abas 'Demandas_Total' e 'funilaria' em: {output_path}")
        Atualiza_PFEP(demand_path,q)
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel com múltiplas abas: {e}")

        try:
            df_final.to_excel(output_path, index=False)
            print(f"Arquivo salvo (APENAS ABA PRINCIPAL, sem estilo) em: {output_path}")
            Atualiza_PFEP(demand_path,q)
        except Exception as e_fallback:
            print(f"Erro fatal ao salvar o arquivo: {e_fallback}")

    return

def le_arquivo_horario() :
   
    # Define as colunas de interesse
    colunas_horarios = ['Supplier Code', 'Horário de Janela']
    
    # Inicializa o DataFrame final como vazio. 
    DF_Horarios = pd.DataFrame(columns=colunas_horarios)
    
    # 1. Define o caminho para a pasta
    caminho_matriz_folder = os.path.join(caminho_base, '1 - MATRIZ')
    
    # 2. Verifica se a pasta existe
    if not os.path.isdir(caminho_matriz_folder):
        print(f"Aviso: A pasta de Matriz não foi encontrada em: {caminho_matriz_folder}")
    else:
        # 3. Encontra o nome do arquivo dinamicamente
        nome_arquivo_horarios_completo = None
        termo_busca = "horários e restrições" # Busca em minúsculo

        try:
            for nome_arquivo in os.listdir(caminho_matriz_folder):
                nome_lower = nome_arquivo.lower()
                # Verifica se o termo está no nome E se é um arquivo Excel
                if termo_busca in nome_lower and (nome_lower.endswith('.xlsx') or nome_lower.endswith('.xls')):
                    nome_arquivo_horarios_completo = os.path.join(caminho_matriz_folder, nome_arquivo)
                    print(f"Arquivo de horários encontrado: {nome_arquivo}")
                    break # Para no primeiro arquivo que encontrar
        except Exception as e:
            print(f"Erro ao listar arquivos na pasta Matriz: {e}")

        # 4. Se o arquivo foi encontrado, tenta ler os dados
        if nome_arquivo_horarios_completo:
            
            # --- CORREÇÃO: Define o nome literal da aba ---
            sheet_name_literal = "FIASA, CKD, MOPAR "
            
            try:
                
                DF_Horarios = pd.read_excel(
                    nome_arquivo_horarios_completo, 
                    sheet_name=sheet_name_literal, # Passa o nome literal da aba
                    usecols=colunas_horarios
                )
                
                print(f"Horários carregados da aba: '{sheet_name_literal}'")
                
                # Remove linhas que possam ter vindo com NaNs (ex: linhas em branco)
                DF_Horarios.dropna(subset=colunas_horarios, inplace=True)
                # print(f"Total de {len(DF_Horarios)} registros de horários carregados.")

            except FileNotFoundError:
                print(f"Erro: O arquivo '{nome_arquivo_horarios_completo}' não foi encontrado.")
            except ValueError as e:
                # Erro comum se as colunas ou a aba não existirem
                print(f"Erro ao ler colunas/aba: {e}")
                print(f"Verifique se a aba '{sheet_name_literal}' E as colunas {colunas_horarios} existem.")
            except Exception as e:
                print(f"Erro inesperado ao ler o arquivo Excel de horários: {e}")
                # (Pode acontecer se a aba não for encontrada)
                if sheet_name_literal not in str(e): # Evita msg duplicada
                     print(f"Verifique se a aba '{sheet_name_literal}' existe no arquivo.")
        
        else:
            if os.path.isdir(caminho_matriz_folder):
                 print(f"Aviso: Nenhum arquivo contendo '{termo_busca}' foi encontrado em {caminho_matriz_folder}")

    return DF_Horarios


def Atualiza_PFEP(path_demandas,q):
    q.put(("status", "Iniciando atualização do PFEP..."))
    caminho_pasta_pfep = os.path.join(caminho_base, '1 - MATRIZ')
    nome_pfep = None
    
    for nome in os.listdir(caminho_pasta_pfep):
        if ('PFEP 2024 DHL' in nome or 'PFEP 2025 DHL' in nome) and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            nome_pfep = os.path.join(caminho_pasta_pfep, nome)
            break
    
    if not nome_pfep:
        print("⚠️ Arquivo PFEP não encontrado.")
        return

    demand_folder = os.path.dirname(path_demandas)
    demand_file = os.path.basename(path_demandas)

    app = None
    wb = None
    wb_demandas = None

    q.put(("status", "Abrindo PFEP e atualizando dados..."))
    try:
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False
        
        app.api.AskToUpdateLinks = False  # 🔒 Prevents Excel popup

        q.put(("status", "Abrindo arquivo de demandas..."))
        wb_demandas = app.books.open(path_demandas)

        # Open PFEP safely, no popup or freeze
        wb = app.books.open(nome_pfep, update_links=False)
        wb.app.api.EnableEvents = True
        ws = wb.sheets['PFEP']
        # Processar_programacao(wb,q)
        q.put(("status", "Atualizando dados do PFEP..."))

        q.put(("status", "Atualizando fórmulas do PFEP..."))
        formula = (
            f'=IF(OFFSET($EL$1,ROW()-1,0)="Fora do Escopo",0,'
            f'IF(OFFSET($CG$1,ROW()-1,0)="YES",0,'
            f'SUMIFS(\'{demand_folder}\\[{demand_file}]Demandas_Total\'!$C:$C,'
            f'\'{demand_folder}\\[{demand_file}]Demandas_Total\'!$B:$B,'
            f'LEFT(OFFSET($V$1,ROW()-1,0),9),'
            f'\'{demand_folder}\\[{demand_file}]Demandas_Total\'!$A:$A,'
            f'OFFSET($C$1,ROW()-1,0))))'
        )

        ws.range('P5').formula = formula
        
        q.put(("status", "Executando macro de recálculo do PFEP..."))
        wb.app.macro("Calcula_Todo_PFEP")()
        q.put(("status", "Macro de recálculo finalizada."))
                        
        # print("Salvando PFEP...")
        q.put(("status", "Salvando PFEP atualizado..."))
        wb.save()
        # print("💾 PFEP atualizado com sucesso e recálculo executado.")
        q.put(("status", "PFEP atualizado com sucesso!"))


        if wb_demandas:
            wb_demandas.close()

        Processar_programacao(wb,q)

    except Exception as e:
        print(f"❌ Erro inesperado durante a atualização do PFEP: {e}")

    finally:
        if wb:
            wb.close()
        
        if app:
            app.display_alerts = True
            app.quit()





def Processar_programacao(pfep, q):

    q.put(("status", "Iniciando atualização da Programação FIASA..."))
    # Ensure 'caminho_base' is accessible
    global caminho_base 
    
    caminho_pasta_matriz = os.path.join(caminho_base, '1 - MATRIZ')
    nome_prog_fiasa = None
    
    # --- Locate Programação FIASA file ---
    for nome in os.listdir(caminho_pasta_matriz):
        if ('Programação FIASA - OFICIAL' in nome or '1. Programação FIASA - OFICIAL' in nome) and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            nome_prog_fiasa = os.path.join(caminho_pasta_matriz, nome)
            break
    
    if not nome_prog_fiasa:
        print("⚠️ Arquivo 'Programação FIASA - OFICIAL' não encontrado.")
        return

    # --- Sheets in PFEP workbook (already open) ---
    ws_pfep = pfep.sheets['PFEP']
    ws_supplier = pfep.sheets['Suppliers DB']
    q.put(("status", "Lendo dados do PFEP e Supplier DB..."))
    # --- Find last rows ---
    last_row_pfep = ws_pfep.range('C' + str(ws_pfep.cells.last_cell.row)).end('up').row
    last_row_supplier = ws_supplier.range('B' + str(ws_supplier.cells.last_cell.row)).end('up').row
    
    print(f"📊 PFEP last row: {last_row_pfep}")
    print(f"📊 Supplier DB last row: {last_row_supplier}")

    # --- Copy PFEP filtered data ---
    data_pfep = None
    q.put(("status", "Filtrando e copiando dados do PFEP..."))
    if ws_pfep.api.AutoFilterMode:
        ws_pfep.api.AutoFilterMode = False
        
    # print("Applying filter to PFEP (Column P <> 0,00)...") 
    q.put(("status", "Aplicando filtro no PFEP (Coluna P <> 0,00)..."))
    
    try:
        # 1️⃣ Define the xlwings range (includes header)
        filter_range = ws_pfep.range(f"A6:HA{last_row_pfep}")

        # 2️⃣ Apply the filter
        filter_range.api.AutoFilter(Field:=16, Criteria1:="<>0,00")

        # 3️⃣ Get visible rows after filtering
        visible_cells = filter_range.api.SpecialCells(12)  # xlCellTypeVisible = 12

        # 4️⃣ Combine values from visible "areas", skipping header
        data_pfep = []
        for i, area in enumerate(visible_cells.Areas):
            area_range = ws_pfep.range(area.Address)
            values = area_range.value

            if isinstance(values, list) and isinstance(values[0], list):
                # Skip first row of first area (header)
                if i == 0:
                    data_pfep.extend(values[1:])
                else:
                    data_pfep.extend(values)
            else:
                # Single row case (no sublists)
                if i > 0:  # skip header only on first block
                    data_pfep.append(values)

        print(f"✅ Copied {len(data_pfep)} visible rows (header excluded) from PFEP.")

            
    except Exception as e:
        print(f"⚠️ No visible data found in PFEP after filtering.")
        q.put(("status", "⚠️ Nenhum dado visível encontrado no PFEP após o filtro."))


    q.put(("status", "Lendo dados do Supplier DB...")) 
    if ws_pfep.api.AutoFilterMode:
        ws_pfep.api.AutoFilterMode = False
    
    # --- Copy Supplier DB data ---
    if ws_supplier.api.AutoFilterMode:
        print("Clearing filters from Supplier DB sheet...")
        ws_supplier.api.AutoFilterMode = False
    q.put(("status", "Copiando dados do Supplier DB..."))   
    data_supplier = ws_supplier.range(f"A6:AT{last_row_supplier}").value
    if data_supplier:
        print("✅ Copied Supplier DB data.")

    # --- Open Programação FIASA workbook ---
    print("Opening Programação FIASA...")
    q.put(("status", "Abrindo Programação FIASA..."))
    app_prog_fiasa = xw.App(visible=True, add_book=False) 
    app_prog_fiasa.display_alerts = False
    app_prog_fiasa.api.AskToUpdateLinks = False
    wb_fiasa = app_prog_fiasa.books.open(nome_prog_fiasa,update_links=False,read_only=False)
    q.put(("status", "Programação FIASA aberta."))
    try:
    
        ws_cola_pfep = wb_fiasa.sheets['COLAR PFEP']
        ws_cola_supplier = wb_fiasa.sheets['COLAR SUPPLIER']
        
        ws_cola_pfep.range('A2:HB15000').clear_contents()
        ws_cola_supplier.range('A2:AT5000').clear_contents()

        if data_pfep:
            print("Pasting PFEP data...")
            # 1️⃣ Paste values
            dest_range = ws_cola_pfep.range('A2').resize(len(data_pfep), len(data_pfep[0]))
            dest_range.value = data_pfep

            # 2️⃣ Copy formatting from row 5
            template_row = ws_cola_pfep.range('5:5')  # Entire row 5 as template
            format_range = ws_cola_pfep.range('A2').resize(len(data_pfep), dest_range.columns.count)
            template_row.api.Copy()
            format_range.api.PasteSpecial(Paste=-4122)  # xlPasteFormats
            ws_cola_pfep.api.Application.CutCopyMode = False

        else:
            print("⚠️ No PFEP data to paste.")

        q.put(("status", "Colando dados do Supplier DB na Programação FIASA..."))
        if data_supplier:
            print("Pasting Supplier DB data...")
            
            dest_range = ws_cola_supplier.range('A2').resize(len(data_supplier), len(data_supplier[0]))
            dest_range.value = data_supplier

            template_row = ws_cola_supplier.range('5:5')  # Entire row 5 as template
            format_range = ws_cola_supplier.range('A2').resize(len(data_supplier), dest_range.columns.count)
            template_row.api.Copy()
            format_range.api.PasteSpecial(Paste=-4122)  # xlPasteFormats
            ws_cola_supplier.api.Application.CutCopyMode = False

        else:
            print("⚠️ No Supplier data to paste.")
        q.put(("status", "Executando recálculo da Programação FIASA..."))
        print("Recalculating, saving, and closing Programação FIASA...")
        wb_fiasa.app.api.CalculateFullRebuild()

        q.put(("status", "Salvando Programação FIASA..."))
        wb_fiasa.save()

    finally:
        q.put(("status", "Finalizando atualização da Programação FIASA..."))
        wb_fiasa.close()
        app_prog_fiasa.quit()
        q.put(("status", "Programação FIASA atualizada com sucesso!"))




