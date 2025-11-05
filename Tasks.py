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
import math
import xlwings as xw
import time
from datetime import date, timedelta
import traceback

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


        # if wb_demandas:
        #     wb_demandas.close()

        Processar_programacao(wb_demandas,wb,q)

    except Exception as e:
        print(f"❌ Erro inesperado durante a atualização do PFEP: {e}")

    finally:
        
        pass
        

def Processar_programacao(wb_demandas,pfep, q):

    q.put(("status", "Iniciando atualização da Programação FIASA..."))
    # Ensure 'caminho_base' is accessible
    global caminho_base 
    
    caminho_pasta_matriz = os.path.join(caminho_base, '1 - MATRIZ')
    nome_prog_fiasa = None
    cargolift_sp_Supplier = None
    cargolift_sp_PFEP = None

    
    # --- Locate Programação FIASA file ---
    for nome in os.listdir(caminho_pasta_matriz):
        if ('Programação FIASA - OFICIAL' in nome or '1. Programação FIASA - OFICIAL' in nome) and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            nome_prog_fiasa = os.path.join(caminho_pasta_matriz, nome)

        if ('Cargolift SP - PFEP' in nome or 'Cargolift SP - PFEP' in nome) and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            cargolift_sp_PFEP = os.path.join(caminho_pasta_matriz, nome)

        if ('Cargolift SP - Suppliers' in nome or 'Cargolift SP - Suppliers' in nome) and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            cargolift_sp_Supplier = os.path.join(caminho_pasta_matriz, nome)
            
    
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
        

        q.put(("status", "Fechando PFEP de Arquivo de demandas..."))
       

        q.put(("status", "PFEP fechado."))

        suppliers_carrier = {
            800006524: "CARGOLIFT",
            800006517: "CARGOLIFT",
            800000656: "CARGOLIFT",
            800046898: "CARGOLIFT",
            800027567: "CARGOLIFT",
            800033665: "CARGOLIFT",
            800046464: "CARGOLIFT",
            800030982: "CARGOLIFT",
            800005848: "CARGOLIFT"
        }

        suppliers_fiasa = {
            800033665: "CARGOLIFT",
        }

        q.put(("status", "Atualizando Carrier na Programação FIASA..."))

        # Worksheet
        ws_Sup_db_corrier = wb_fiasa.sheets['Suppliers DB']
        if ws_Sup_db_corrier.api.AutoFilterMode:
            ws_Sup_db_corrier.api.AutoFilterMode = False

        # Find last used row in column C
        last_row = ws_Sup_db_corrier.range('C' + str(ws_Sup_db_corrier.cells.last_cell.row)).end('up').row

        q.put(("status", f"Processando {last_row - 1} linhas para atualização de Carrier..."))
        # Get data
        supplier_codes = ws_Sup_db_corrier.range(f'C2:C{last_row}').value
        fca_values = ws_Sup_db_corrier.range(f'D2:D{last_row}').value

        # Loop through rows
        for i, code in enumerate(supplier_codes, start=2):
            if code is None:
                continue

            try:
                code_int = int(code)
            except:
                continue

            # 1️⃣ Check in main carrier dict
            if code_int in suppliers_carrier:
                carrier_value = suppliers_carrier[code_int]

                # 2️⃣ If in FIASA dict, only update if column D == "FCA"
                if code_int in suppliers_fiasa:
                    if "FCA" in str(fca_values[i - 2]).strip().upper() :
                        ws_Sup_db_corrier.range(f'AB{i}').value = carrier_value
                else:
                    ws_Sup_db_corrier.range(f'AB{i}').value = carrier_value
        q.put(("status", "Carrier atualizado na Programação FIASA."))
        

        q.put(("status", "Salvando Programação FIASA..."))
        wb_fiasa.save()

        progrma_cargolift (cargolift_sp_PFEP,cargolift_sp_Supplier,wb_fiasa,q,pfep,wb_demandas)

    finally:
        q.put(("status", "Finalizando atualização da Programação FIASA..."))
        wb_fiasa.close()
        app_prog_fiasa.quit()
        q.put(("status", "Programação FIASA atualizada com sucesso!"))


def progrma_cargolift(arquivo_cargolift_sp_PFEP, arquivo_cargolift_sp_Supplier, wb_fiasa, q,pfep,wb_demandas):



    q.put(("status", "Iniciando atualização da Programação Cargolift SP..."))

    ws_pfep = wb_fiasa.sheets['PFEP']
    ws_supplier_db = wb_fiasa.sheets['Suppliers DB']

    # Remove any active filters before we start
    for ws in [ws_pfep, ws_supplier_db]:
        try:
            if ws.api.FilterMode:
                ws.api.ShowAllData()
        except:
            pass
        ws.api.AutoFilterMode = False

    # ========== PFEP ==========
    q.put(("status", "Filtrando PFEP por CARGOLIFT..."))
    last_row_pfep = ws_pfep.range('A' + str(ws_pfep.cells.last_cell.row)).end('up').row

    # PFEP columns go to AQ (43)
    filter_range_pfep = ws_pfep.range(f"A1:AQ{last_row_pfep}")
    filter_range_pfep.api.AutoFilter(Field:=43, Criteria1:="CARGOLIFT")

    # Collect visible data
    try:
        visible_cells_pfep = filter_range_pfep.api.SpecialCells(12)  # xlCellTypeVisible
        data_pfep = []
        for i, area in enumerate(visible_cells_pfep.Areas):
            area_range = ws_pfep.range(area.Address)
            values = area_range.value
            if isinstance(values, list) and isinstance(values[0], list):
                if i == 0:
                    data_pfep.extend(values[1:])  # skip header
                else:
                    data_pfep.extend(values)
            else:
                if i > 0:
                    data_pfep.append(values)
        print(f"✅ Copied {len(data_pfep)} visible rows from PFEP.")
    except Exception:
        data_pfep = []
        print("⚠️ Nenhum dado visível encontrado em PFEP com filtro 'CARGOLIFT'.")

    # Clear filter
    try:
        if ws_pfep.api.FilterMode:
            ws_pfep.api.ShowAllData()
    except:
        pass
    ws_pfep.api.AutoFilterMode = False

    # ========== SUPPLIER DB ==========
    q.put(("status", "Filtrando Supplier DB por CARGOLIFT..."))
    last_row_supplier = ws_supplier_db.range('A' + str(ws_supplier_db.cells.last_cell.row)).end('up').row

    # Supplier DB columns go to AJ (36)
    filter_range_supplier = ws_supplier_db.range(f"A1:AJ{last_row_supplier}")
    filter_range_supplier.api.AutoFilter(Field:=28, Criteria1:="CARGOLIFT")  # AB = 28

    try:
        visible_cells_supplier = filter_range_supplier.api.SpecialCells(12)
        data_supplier = []
        for i, area in enumerate(visible_cells_supplier.Areas):
            area_range = ws_supplier_db.range(area.Address)
            values = area_range.value
            if isinstance(values, list) and isinstance(values[0], list):
                if i == 0:
                    data_supplier.extend(values[1:])
                else:
                    data_supplier.extend(values)
            else:
                if i > 0:
                    data_supplier.append(values)
        print(f"✅ Copied {len(data_supplier)} visible rows from Supplier DB.")
    except Exception:
        data_supplier = []
        print("⚠️ Nenhum dado visível encontrado em Supplier DB com filtro 'CARGOLIFT'.")

    try:
        if ws_supplier_db.api.FilterMode:
            ws_supplier_db.api.ShowAllData()
    except:
        pass
    ws_supplier_db.api.AutoFilterMode = False

    # ========== DESTINATION FILES ==========
    q.put(("status", "Abrindo arquivos de destino (Cargolift SP)..."))
    # ✅ Single Excel instance for both destination files
    app_cargolift = xw.App(visible=True, add_book=False)
    app_cargolift.display_alerts = False
    app_cargolift.api.AskToUpdateLinks = False

    wb_cargolift_sp_PFEP = app_cargolift.books.open(
        arquivo_cargolift_sp_PFEP, update_links=False, read_only=False
    )
    wb_cargolift_sp_Supplier = app_cargolift.books.open(
        arquivo_cargolift_sp_Supplier, update_links=False, read_only=False
    )

    
    ws_dest_pfep = wb_cargolift_sp_PFEP.sheets['PFEP']
    ws_dest_supplier = wb_cargolift_sp_Supplier.sheets['Cargolift SP - Suppliers DB Wk ']

    # Clear destination before pasting
    ws_dest_pfep.range('A3').expand().clear_contents()
    ws_dest_supplier.range('A3').expand().clear_contents()

    # Paste PFEP
    if data_pfep:
        print("📋 Colando dados PFEP...")
        ws_dest_pfep.range('A3').value = data_pfep

    # Paste Supplier DB
    if data_supplier:
        print("📋 Colando dados Supplier DB...")
        ws_dest_supplier.range('A3').value = data_supplier
    
    # Save and close
    wb_cargolift_sp_PFEP.save()
    wb_cargolift_sp_Supplier.save()
    Corregir_peso_e_valor(wb=wb_cargolift_sp_Supplier,demandas_path =wb_demandas ,q = q,pfep_source = pfep)

    wb_cargolift_sp_PFEP.close()
    wb_cargolift_sp_Supplier.close()
    app_cargolift.quit()

    q.put(("status", "✅ Atualização da Programação Cargolift SP concluída!"))
    print("✅ Atualização concluída com sucesso!")





# --- 1. FUNÇÃO DE NORMALIZAÇÃO CORRIGIDA ---
def normalize_value(value):
    """
    Converte valor para string, remove espaços, coloca em maiúsculo
    e remove '.0' do final (crucial para PNs lidos como float).
    """
    if value is None:
        return ""
    # Converte para string, remove espaços
    s = str(value).strip().upper()
    
    # Remove ".0" de PNs que o Excel leu como float
    if s.endswith('.0'):
        return s[:-2] # Remove os dois últimos caracteres ('.0')
        
    return s




def Corregir_peso_e_valor(q, wb=None, demandas_path=None, pfep_source=None):
    
    # Renomeia variáveis para clareza
    wb_cargolift = wb
    wb_demandas = demandas_path
    wb_pfep = pfep_source
    
    global caminho_base 
    global normalize_value 

    q.put(("status",f"--- 🚀 Iniciando a função Corregir_peso_e_valor ---"))
    q.put(("status",f"Workbook Alvo: {wb_cargolift.name}"))
    q.put(("status",f"Workbook Demandas: {wb_demandas.name}"))
    q.put(("status",f"Workbook PFEP: {wb_pfep.name}"))

    try:
        # --- 1. Obter dados do 'wb_demandas' (PN e SAP) ---
        q.put(("status", "Etapa 1: Lendo dados do 'Demandas'..."))
        sheet_demandas = wb_demandas.sheets.active
        df_demandas = sheet_demandas.range('A1').expand().options(pd.DataFrame, index=False, header=True).value
        
        if 'PN' not in df_demandas.columns or 'SAP' not in df_demandas.columns:
            q.put(("status", "ERRO: Colunas 'PN' ou 'SAP' não encontradas no arquivo Demandas."))
            return

        df_demandas['PN_norm'] = df_demandas['PN'].apply(normalize_value)
        df_demandas['SAP_norm'] = df_demandas['SAP'].apply(normalize_value)
        q.put(("status", f"Ok. Encontrados {len(df_demandas)} registros no 'Demandas'."))

        # --- 2. Obter dados do 'wb_pfep' (Part Number RTM) ---
        q.put(("status","Etapa 2: Lendo dados do 'PFEP'..."))
        sheet_pfep = wb_pfep.sheets.active
        header_range = sheet_pfep.range('A6').expand('right')
        header_values = header_range.value
        
        pn_col_index = None # Índice 0-based
        if isinstance(header_values, list):
            try:
                pn_col_index = header_values.index('Part Number RTM')
            except ValueError:
                q.put(("status", "ERRO: Coluna 'Part Number RTM' não encontrada na linha 6 do 'PFEP'."))
                return
        else:
            q.put(("status", "ERRO: Não foi possível ler os cabeçalhos da linha 6 do 'PFEP'."))
            return

        pn_col_letter = xw.utils.col_name(pn_col_index + 1)
        q.put(("status",f"Ok. Coluna 'Part Number RTM' encontrada em: {pn_col_letter}6 (Índice 0-based: {pn_col_index})"))
        
        last_row_pfep = sheet_pfep.range(f'{pn_col_letter}7').end('down').row
        pfep_pns_raw = sheet_pfep.range(f'{pn_col_letter}7:{pn_col_letter}{last_row_pfep}').value
        
        pfep_pns_set = {normalize_value(pn) for pn in pfep_pns_raw if pn is not None}
        q.put(("status",f"Ok. Encontrados {len(pfep_pns_set)} PNs únicos no 'PFEP'."))
        
        if pfep_pns_set:
             q.put(("status",f" 	-> Debug: Exemplo de PN normalizado do PFEP: '{next(iter(pfep_pns_set))}'"))

        # --- 3. "XLOOKUP": Encontrar PNs que estão no 'Demandas' mas não no 'PFEP' ---
        q.put(("status","Etapa 3: Cruzando dados (PNs Demandas x PFs PFEP)..."))
        
        missing_pns_mask = ~df_demandas['PN_norm'].isin(pfep_pns_set)
        df_missing_pns = df_demandas[missing_pns_mask]
        
        found_pns_mask = ~missing_pns_mask
        found_pns_count = found_pns_mask.sum()
        missing_pns_count = len(df_missing_pns)
        
        q.put(("status", f"Ok. PNs de 'Demandas' ENCONTRADOS no 'PFEP': {found_pns_count}"))
        q.put(("status", f"Ok. PNs de 'Demandas' NÃO ENCONTRADOS no 'PFEP': {missing_pns_count}"))

        if missing_pns_count > 0 and not df_missing_pns.empty:
            q.put(("status", f" 	-> Exemplo de PN NÃO encontrado: '{df_missing_pns.iloc[0]['PN_norm']}' (Original: '{df_missing_pns.iloc[0]['PN']}')"))
        if found_pns_count > 0:
            df_found = df_demandas[found_pns_mask]
            if not df_found.empty:
                q.put(("status", f" 	-> Exemplo de PN ENCONTRADO: '{df_found.iloc[0]['PN_norm']}' (Original: '{df_found.iloc[0]['PN']}')"))

        if df_missing_pns.empty:
            q.put(("status","Ok. Nenhum PN do 'Demandas' está faltando no 'PFEP'. Encerrando."))
            q.put(("status","--- ✅ Processo concluído (sem alterações) ---"))
            return



        if wb_demandas :
                wb_demandas.close()

        if  wb_pfep :
             wb_pfep.close()



            
        q.put(("status", f"Ok. {len(df_missing_pns)} PNs faltantes serão filtrados pelo SAP..."))

        # --- 4. Filtrar pelos SAPs do JSON ---
        q.put(("status", "Etapa 4: Filtrando PNs faltantes pelo JSON 'Forncedores_Responsavel.json'..."))
        json_path = os.path.join(caminho_base,"Bases","Forncedores_Responsavel.json")
        
        if not os.path.exists(json_path):
            q.put(("status", f"ERRO: Arquivo JSON '{json_path}' não encontrado."))
            return

        with open(json_path, 'r', encoding='utf-8') as f:
            fornecedores_data = json.load(f)
        
        valid_saps_from_json = {normalize_value(sap): data for sap, data in fornecedores_data.items()}
        q.put(("status",f"Ok. Carregados {len(valid_saps_from_json)} SAPs do arquivo JSON."))

        saps_to_update_mask = df_missing_pns['SAP_norm'].isin(valid_saps_from_json.keys())
        df_final_list = df_missing_pns[saps_to_update_mask]

        if df_final_list.empty:
            q.put(("status","Ok. Nenhum dos PNs faltantes corresponde a um SAP válido no JSON. Encerrando."))
            q.put(("status","--- ✅ Processo concluído (sem alterações) ---"))
            return
        
        unique_saps_to_update = df_final_list['SAP_norm'].unique()
        q.put(("status", f"Ok. Processando atualizações para {len(unique_saps_to_update)} SAP(s) único(s)."))
        q.put(("status", f" 	-> SAPs para atualizar: {list(unique_saps_to_update)}"))


        # --- 5 & 6. Atualizar o 'wb_cargolift' ---
        q.put(("status","Etapa 5/6: Iniciando atualizações no 'Cargolift'..."))
        
        sheet_target = None
        target_sheet_name_fragment = "SUPPLIERS DB WK" 
        
        try:
            sheet_names_list = [sheet.name for sheet in wb_cargolift.sheets]
            q.put(("status", f"Planilhas encontradas no '{wb_cargolift.name}': {sheet_names_list}"))

            for sheet in wb_cargolift.sheets:
                normalized_sheet_name = sheet.name.strip().upper() 
                if target_sheet_name_fragment in normalized_sheet_name:
                    sheet_target = sheet
                    q.put(("status", f"SUCESSO: Planilha de atualização encontrada: '{sheet.name}'"))
                    break 

            if sheet_target is None:
                q.put(("status", f"ERRO: Nenhuma planilha encontrada no workbook '{wb_cargolift.name}' que contenha o nome '{target_sheet_name_fragment}'."))
                return
                
        except Exception as e:
            q.put(("status", f"ERRO: Falha ao tentar encontrar a planilha de atualização. Detalhe: {e}"))
            return

        # --- ⬇️⬇️⬇️ LÓGICA DE DIAS REMOVIDA (Conforme solicitado) ⬇️⬇️⬇️ ---
        m3_cols = ['I', 'J', 'K', 'L', 'M', 'N'] # Mon, Tue, Wed, Thu, Fri, Sat
        kg_cols = ['P', 'Q', 'R', 'S', 'T', 'U'] # Mon, Tue, Wed, Thu, Fri, Sat
        # --- ⬆️⬆️⬆️ FIM DA REMOÇÃO ⬆️⬆️⬆️ ---


        q.put(("status",f"Mapeando 'Supplier Code' da planilha '{sheet_target.name}' para otimização..."))
        last_row_target = sheet_target.range('C1').end('down').row
        sap_codes_raw = sheet_target.range(f'C2:C{last_row_target}').value
        
        sap_to_row_map = {}
        if isinstance(sap_codes_raw, list):
            for i, sap in enumerate(sap_codes_raw):
                sap_norm = normalize_value(sap) 
                if sap_norm:
                    sap_to_row_map[sap_norm] = i + 2
        else: # Caso de uma única célula
            sap_norm = normalize_value(sap_codes_raw)
            if sap_norm:
                sap_to_row_map[sap_norm] = 2

        q.put(("status",f"Ok. {len(sap_to_row_map)} SAPs mapeados na planilha 'Cargolift'."))

        if sap_to_row_map:
            exemplo_sap_mapeado = next(iter(sap_to_row_map.keys()))
            q.put(("status", f" 	-> Debug: Exemplo de SAP normalizado do 'Cargolift' (Coluna C): '{exemplo_sap_mapeado}'"))

        updates_made = 0
        for sap_norm in unique_saps_to_update:
            if sap_norm in sap_to_row_map:
                row_number = sap_to_row_map[sap_norm]
                
                json_data_for_sap = valid_saps_from_json[sap_norm]
                m3_to_add = json_data_for_sap.get("M3", 0)
                kg_to_add = json_data_for_sap.get("Kg", 0)
                
                # --- ⬇️⬇️⬇️ NOVA LÓGICA DE ATUALIZAÇÃO ⬇️⬇️⬇️ ---
                target_m3_col = None
                target_kg_col = None
                current_m3 = 0
                current_kg = 0

                try:
                    # Lê os valores da semana inteira para esta linha
                    m3_values = sheet_target.range(f'I{row_number}:N{row_number}').value
                    kg_values = sheet_target.range(f'P{row_number}:U{row_number}').value
                except Exception as e:
                    q.put((f"ERRO: Falha ao ler dados da linha {row_number}. {e}"))
                    continue # Pula para o próximo SAP

                # Encontra a primeira coluna de M3 com valor > 0
                if isinstance(m3_values, list):
                    for i, val in enumerate(m3_values):
                        if val is not None:
                            try:
                                if float(val) > 0:
                                    target_m3_col = m3_cols[i] # Ex: 'I'
                                    current_m3 = float(val)
                                    break # Para no primeiro que encontrar
                            except ValueError:
                                continue # Ignora valores não numéricos (texto)
                
                # Encontra a primeira coluna de Kg com valor > 0
                if isinstance(kg_values, list):
                    for i, val in enumerate(kg_values):
                        if val is not None:
                            try:
                                if float(val) > 0:
                                    target_kg_col = kg_cols[i] # Ex: 'P'
                                    current_kg = float(val)
                                    break # Para no primeiro que encontrar
                            except ValueError:
                                continue # Ignora valores não numéricos

                # Se encontrou colunas de planejamento para M3 e Kg, atualiza
                if target_m3_col and target_kg_col:
                    new_m3 = current_m3 + m3_to_add
                    new_kg = current_kg + kg_to_add
                    
                    # Escreve de volta nas colunas específicas que encontrou
                    sheet_target.range(f'{target_m3_col}{row_number}').value = new_m3
                    sheet_target.range(f'{target_kg_col}{row_number}').value = new_kg
                    
                    q.put(("status", f" 	-> SUCESSO: SAP {sap_norm} (Linha {row_number}): Col M3 '{target_m3_col}' ({current_m3} + {m3_to_add} = {new_m3}). Col Kg '{target_kg_col}' ({current_kg} + {kg_to_add} = {new_kg})."))
                    updates_made += 1
                else:
                    # Se não encontrou nenhuma coluna com valor > 0
                    q.put(("status", f" 	-> AVISO: SAP '{sap_norm}' (Linha {row_number}): Nenhum valor de planejamento (>0) encontrado nas colunas I-N e P-U. Pulando."))
                # --- ⬆️⬆️⬆️ FIM DA NOVA LÓGICA ⬆️⬆️⬆️ ---
                
            else:
                q.put(("status", f" 	-> AVISO: SAP '{sap_norm}' (da lista de PNs faltantes) não foi encontrado na coluna 'Supplier Code' do 'Cargolift'. Pulando."))
        
        q.put(("status",f"--- ✅ Processo concluído. {updates_made} atualizações de SAP realizadas. ---"))






        # calling the cut and copy fucntion to cut and copy data form each programme to the cargo lift.
        # this is the final part of the fuction 



        q.put(("status","-----------REABRINDO OS ARQUIVOS PARA COMEÇAR COPIA E COLAR-------------"))
        q.put(("progress",70))
        
        # Copiar_planejamentos_para_cargolift(wb_cargolift = wb_cargolift,q=q) 
        # wb_cargolift_sp_Supplier.close()
    except Exception as e:
        q.put(("status",f"--- ❌ ERRO CRÍTICO na função ---"))
        q.put(("status", f"Erro: {e}"))
        q.put(("status",f"Traceback: {traceback.format_exc()}"))






def Copiar_planejamentos_para_cargolift_Arquivos(wb_cargolift = None,q = None) :

    q.put(("status","Inicializando a funcção de copia e colar os dados"))

    caminho_pasta_matriz = os.path.join(caminho_base, '1 - MATRIZ')
    caminho_pasta_programacoes = os.path.join(caminho_base, 'Planilhas_Recebidos')

    nome_prog_fiasa = None
    cargolift_sp_Supplier = None
    cargolift_sp_PFEP = None
    Programacao_FPT_Sul =  None

    #-------------- Programações  -------------

    Cargolift_prog_PR =  None
    Cargolift_prog_FIAPE = None
    cargolift_prog_FPT =  None
    cargolift_prog_MOPAR =  None
    cargolift_sp_Part_Number_MOPAR =  None
    Programacao_CKD =  None

    cargolift_sp_Embalagem_FPT =  None
    cargolift_sp_Embalagem_2_Atualizada = None
    CGLFT_SP_Embalagem_FIAPE = None

    
    # --- Locate Programação FIASA file ---
    for nome in os.listdir(caminho_pasta_matriz):
        if ('Programação FIASA - OFICIAL' in nome or '1. Programação FIASA - OFICIAL' in nome) and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            nome_prog_fiasa = os.path.join(caminho_pasta_matriz, nome)

        if ('Cargolift SP - PFEP' in nome or 'Cargolift SP - PFEP' in nome) and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            cargolift_sp_PFEP = os.path.join(caminho_pasta_matriz, nome)

        if wb_cargolift == None :
            if ('Cargolift SP - Suppliers' in nome or 'Cargolift SP - Suppliers' in nome) and nome.endswith(('.xlsm', '.xls', '.xlsx')):
                cargolift_sp_Supplier = os.path.join(caminho_pasta_matriz, nome)
        
        if ('PROGRAMAÇÃO SUL' in nome or 'PROGRAMAÇÃO SUL' in nome) and nome.endswith(('.xlsm', '.xls', '.xlsx')):
            Programacao_FPT_Sul= os.path.join(caminho_pasta_matriz, nome)


        
    # *** CORRECTION: Changed 'nome' to 'prog' in this loop to correctly find files ***
    for prog in os.listdir(caminho_pasta_programacoes):
        if ('Cargolift SP - PFEP  Porto Real' in prog or 'Cargolift SP - PFEP  Porto Real' in prog) and prog.endswith(('.xlsm', '.xls', '.xlsx')):
            Cargolift_prog_PR = os.path.join(caminho_pasta_programacoes, prog)

        if ('Cargolift SP - FIAPE ' in prog or 'Cargolift SP - FIAPE ' in prog) and prog.endswith(('.xlsm', '.xls', '.xlsx')):
            Cargolift_prog_FIAPE = os.path.join(caminho_pasta_programacoes, prog)

        if ('FPT BT' in prog or 'FPT BT' in prog) and prog.endswith(('.xlsm', '.xls', '.xlsx')):
            cargolift_prog_FPT = os.path.join(caminho_pasta_programacoes, prog)

        if ('Cargolift SP Fornecedor MOPAR' in prog or 'Cargolift SP Fornecedor MOPAR' in prog) and prog.endswith(('.xlsm', '.xls', '.xlsx')):
            cargolift_prog_MOPAR = os.path.join(caminho_pasta_programacoes, prog)

        if ('Cargolift SP Part Number' in prog or 'Cargolift SP Part Number' in prog) and prog.endswith(('.xlsm', '.xls', '.xlsx')):
            cargolift_sp_Part_Number_MOPAR = os.path.join(caminho_pasta_programacoes, prog)

        if ('PROGRAMAÇÃO CKD' in prog or 'PROGRAMAÇÃO CKD' in prog) and prog.endswith(('.xlsm', '.xls', '.xlsx')):
            Programacao_CKD = os.path.join(caminho_pasta_programacoes, prog)

#------------------------------ Embalagens  -------------------------------------#
        if ('Cargolift SP - Embalagem FPT' in prog or 'Cargolift SP - Embalagem FPT' in prog) and prog.endswith(('.xlsm', '.xls', '.xlsx')):
            cargolift_sp_Embalagem_FPT = os.path.join(caminho_pasta_programacoes, prog)

        if ('Cargolift SP - Embalagem 2 Atualizada' in prog or 'Cargolift SP - Embalagem 2 Atualizada' in prog) and prog.endswith(('.xlsm', '.xls', '.xlsx')):
            cargolift_sp_Embalagem_2_Atualizada = os.path.join(caminho_pasta_programacoes, prog)

        if ('CGLFT SP - Embalagem - FIAPE' in prog or 'CGLFT SP - Embalagem - FIAPE' in prog) and prog.endswith(('.xlsm', '.xls', '.xlsx')):
            CGLFT_SP_Embalagem_FIAPE = os.path.join(caminho_pasta_programacoes, prog)


    # Opening main/Master file where all other filkes will be paste
    
    # Check if files were found before opening
    if not cargolift_sp_PFEP:
        q.put(("status", "ERRO: Arquivo 'Cargolift SP - PFEP' não encontrado na Matriz."))
        return
    if not cargolift_sp_Supplier:
        q.put(("status", "ERRO: Arquivo 'Cargolift SP - Suppliers' não encontrado na Matriz."))
        return
        
    q.put(("status", "Abrindo arquivos mestre..."))
    app_cargolift_sp_PFEP = xw.App(visible=True, add_book=False) 
    app_cargolift_sp_PFEP.display_alerts = False
    app_cargolift_sp_PFEP.api.AskToUpdateLinks = False
    wb_cargolift_sp_PFEP = app_cargolift_sp_PFEP.books.open(cargolift_sp_PFEP,update_links=False,read_only=False)
    sheet_wb_cargolift_sp_PFEP = wb_cargolift_sp_PFEP.sheets['PFEP']


    app_cargolift_sp_Supplier = xw.App(visible=True, add_book=False) 
    app_cargolift_sp_Supplier.display_alerts = False
    app_cargolift_sp_Supplier.api.AskToUpdateLinks = False
    wb_cargolift_sp_Supplier = app_cargolift_sp_Supplier.books.open(cargolift_sp_Supplier,update_links=False,read_only=False) # Use a app correta
    sheet_wb_cargolift_sp_Supplier = wb_cargolift_sp_Supplier.sheets['Cargolift SP - Suppliers DB Wk ']

    # Check if FPT file was found
    if cargolift_prog_FPT:
        q.put(("status", "Arquivo 'FPT BT' encontrado, iniciando cópia..."))
        Copiar_planejamentos_para_FPT_BT(
            q=q, cargolift_prog_FPT = cargolift_prog_FPT,sheet_wb_cargolift_sp_PFEP = sheet_wb_cargolift_sp_PFEP,
            sheet_wb_cargolift_sp_Supplier = sheet_wb_cargolift_sp_Supplier, Programacao_FPT_Sul = Programacao_FPT_Sul,
            programacao_fiasa_path = nome_prog_fiasa, Programacao_CKD_path = Programacao_CKD)
    else:
        q.put(("status", "AVISO: Arquivo 'FPT BT' não encontrado em Planilhas_Recebidos."))
        
    # You might want to save and close the master files here or in the calling function
    wb_cargolift_sp_PFEP.save()
    wb_cargolift_sp_PFEP.close()
    app_cargolift_sp_PFEP.quit()
    q.put(("status", "Processo FPT BT concluído."))





def Copiar_planejamentos_para_FPT_BT(q=None, cargolift_prog_FPT = None , sheet_wb_cargolift_sp_PFEP = None, sheet_wb_cargolift_sp_Supplier =  None, Programacao_FPT_Sul =  None, programacao_fiasa_path =  None ,Programacao_CKD_path = None):
    
    app_cargolift_prog_FPT = None
    wb_cargolift_prog_FPT = None
    
    try:
        q.put(("status", "Abrindo FPT BT..."))
        app_cargolift_prog_FPT = xw.App(visible=True, add_book=False) 
        app_cargolift_prog_FPT.display_alerts = False
        app_cargolift_prog_FPT.api.AskToUpdateLinks = False
        wb_cargolift_prog_FPT = app_cargolift_prog_FPT.books.open(cargolift_prog_FPT,update_links=False,read_only=False)

        sheet_cargolift_prog_FPT_PFEP = wb_cargolift_prog_FPT.sheets['PFEP']
        sheet_cargolift_prog_FPT__Suppliers_DB  = wb_cargolift_prog_FPT.sheets[' Suppliers DB ']

        filter_criteria = ['Milk Run SP', 'Line Haul SP']
        filter_criteria_sul = ['MR SUL']
        xlCellTypeVisible = 12 # VBA Constant for SpecialCells

        # --- Initialize lists for SUL data ---
        Dado_PFEP_sul_a_colar = []
        Dado_supplier_sul_a_colar = []

        # --- 1. Process Suppliers DB ---
        q.put(("status", "Processando FPT Suppliers DB..."))
        sheet_sup = sheet_cargolift_prog_FPT__Suppliers_DB
        sheet_sup.api.AutoFilterMode = False # Clear existing filters
        
        last_row_sup = sheet_sup.range('A' + str(sheet_sup.cells.last_cell.row)).end('up').row
        
        if last_row_sup > 2:
            filter_range_sup = sheet_sup.range(f'A2:X{last_row_sup}')
            data_range_sup = sheet_sup.range(f'A3:X{last_row_sup}')
            
            # --- Apply filter for SP ---
            filter_range_sup.api.AutoFilter(Field:=24, Criteria1:=filter_criteria, Operator:=xw.constants.AutoFilterOperator.xlFilterValues)
            
            try:
                visible_cells_sup = data_range_sup.api.SpecialCells(xlCellTypeVisible)
                data_sup = []
                for area in visible_cells_sup.Areas:
                    values = sheet_sup.range(area.Address).value
                    if isinstance(values, list) and isinstance(values[0], list): 
                        data_sup.extend(values)
                    else: 
                        data_sup.append(values)
                
                if not data_sup:
                    q.put(("status", "Nenhum dado SP visível em FPT Suppliers."))
                else:
                    next_row_sup = sheet_wb_cargolift_sp_Supplier.range('A' + str(sheet_wb_cargolift_sp_Supplier.cells.last_cell.row)).end('up').row + 1
                    sheet_wb_cargolift_sp_Supplier.range(f'A{next_row_sup}').value = data_sup
                    q.put(("status", f"{len(data_sup)} linhas SP coladas em 'Cargolift SP - Suppliers DB Wk '"))

            except Exception as e:
                q.put(("status", f"Nenhum dado SP encontrado no filtro FPT Suppliers: {e}"))
            
            # --- Apply filter for SUL ---
            q.put(("status", "Filtrando FPT Suppliers para MR SUL..."))
            filter_range_sup.api.AutoFilter(Field:=24, Criteria1:=filter_criteria_sul, Operator:=xw.constants.AutoFilterOperator.xlFilterValues)
            
            try:
                visible_cells_sup_sul = data_range_sup.api.SpecialCells(xlCellTypeVisible)
                for area in visible_cells_sup_sul.Areas:
                    values = sheet_sup.range(area.Address).value
                    if isinstance(values, list) and isinstance(values[0], list):
                        Dado_supplier_sul_a_colar.extend(values)
                    else:
                        Dado_supplier_sul_a_colar.append(values)
                
                if not Dado_supplier_sul_a_colar:
                    q.put(("status", "Nenhum dado SUL visível em FPT Suppliers."))
                else:
                    q.put(("status", f"{len(Dado_supplier_sul_a_colar)} linhas SUL coletadas de Suppliers."))

            except Exception as e:
                 q.put(("status", f"Nenhum dado SUL encontrado no filtro FPT Suppliers: {e}"))

            sheet_sup.api.AutoFilterMode = False # Clear filter
        else:
            q.put(("status", "Nenhum dado para processar em FPT Suppliers DB."))

        # --- 2. Process PFEP ---
        q.put(("status", "Processando FPT PFEP..."))
        sheet_pfep = sheet_cargolift_prog_FPT_PFEP
        sheet_pfep.api.AutoFilterMode = False # Clear existing filters
        
        last_row_pfep = sheet_pfep.range('AX' + str(sheet_pfep.cells.last_cell.row)).end('up').row
        
        if last_row_pfep > 3:
            filter_range_pfep = sheet_pfep.range(f'A3:AX{last_row_pfep}')
            data_range_pfep = sheet_pfep.range(f'M4:BI{last_row_pfep}')
            
            # --- Apply filter for SP ---
            filter_range_pfep.api.AutoFilter(Field:=50, Criteria1:=filter_criteria, Operator:=xw.constants.AutoFilterOperator.xlFilterValues)
            
            try:
                visible_cells_pfep = data_range_pfep.api.SpecialCells(xlCellTypeVisible)
                data_pfep = []
                for area in visible_cells_pfep.Areas:
                    values = sheet_pfep.range(area.Address).value
                    if isinstance(values, list) and isinstance(values[0], list): 
                        data_pfep.extend(values)
                    else: 
                        data_pfep.append(values)
                
                if not data_pfep:
                    q.put(("status", "Nenhum dado SP visível em FPT PFEP."))
                else:
                    next_row_pfep = sheet_wb_cargolift_sp_PFEP.range('A' + str(sheet_wb_cargolift_sp_PFEP.cells.last_cell.row)).end('up').row + 1
                    sheet_wb_cargolift_sp_PFEP.range(f'A{next_row_pfep}').value = data_pfep
                    q.put(("status", f"{len(data_pfep)} linhas SP coladas em 'PFEP'"))
                    
            except Exception as e:
                q.put(("status", f"Nenhum dado SP encontrado no filtro FPT PFEP: {e}"))
                
            # --- Apply filter for SUL ---
            q.put(("status", "Filtrando FPT PFEP para MR SUL..."))
            filter_range_pfep.api.AutoFilter(Field:=50, Criteria1:=filter_criteria_sul, Operator:=xw.constants.AutoFilterOperator.xlFilterValues)

            try:
                visible_cells_pfep_sul = data_range_pfep.api.SpecialCells(xlCellTypeVisible)
                for area in visible_cells_pfep_sul.Areas:
                    values = sheet_pfep.range(area.Address).value
                    if isinstance(values, list) and isinstance(values[0], list):
                        Dado_PFEP_sul_a_colar.extend(values)
                    else:
                        Dado_PFEP_sul_a_colar.append(values)

                if not Dado_PFEP_sul_a_colar:
                    q.put(("status", "Nenhum dado SUL visível em FPT PFEP."))
                else:
                    q.put(("status", f"{len(Dado_PFEP_sul_a_colar)} linhas SUL coletadas de PFEP."))
            
            except Exception as e:
                q.put(("status", f"Nenhum dado SUL encontrado no filtro FPT PFEP: {e}"))

            sheet_pfep.api.AutoFilterMode = False # Clear filter
        else:
            q.put(("status", "Nenhum dado para processar em FPT PFEP."))

       
    except Exception as e:
        q.put(("status", f"ERRO ao processar FPT_BT: {e}"))
        print(f"ERRO ao processar FPT_BT: {e}")
    
    finally:
        # Close the local FPT file
        if wb_cargolift_prog_FPT :
            wb_cargolift_prog_FPT.close()
        if app_cargolift_prog_FPT:
            app_cargolift_prog_FPT.quit()
        q.put(("status", "Arquivo FPT BT fechado."))

    # --- Call function to paste SUL data ---
    if Programacao_FPT_Sul and programacao_fiasa_path:
        if Dado_PFEP_sul_a_colar or Dado_supplier_sul_a_colar:
            Copiar_e_Colar_Programacao_Sul(Programacao_FPT_Sul_path = Programacao_FPT_Sul, q = q,
                                            Dado_PFEP_a_colar = Dado_PFEP_sul_a_colar, 
                                            Dado_supplier_a_colar = Dado_supplier_sul_a_colar ,
                                            programacao_fiasa_path = programacao_fiasa_path,
                                            Programacao_CKD_path = Programacao_CKD_path)
        else:
            q.put(("status", "Nenhum dado SUL encontrado. Pulando cópia para arquivo SUL."))
    else:
        q.put(("status", "Caminho para Programacao_FPT_Sul não fornecido. Pulando cópia SUL."))





def Copiar_e_Colar_Programacao_Sul(Programacao_FPT_Sul_path =  None, q = None , Dado_PFEP_a_colar =  None, Dado_supplier_a_colar =  None, programacao_fiasa_path =  None, Programacao_CKD_path =  None):
    
    # --- Step 1: Initialize variables to hold FIASA data ---
    data_fiasa_pfep = []
    data_fiasa_sup = []
    
    app_programacao_fiasa = None # Source workbook app
    wb_programacao_fiasa = None # Source workbook
    
    app_cargolift_prog_FPT_sul = None
    wb_cargolift_prog_FPT_sul = None # Destination workbook

    xlCellTypeVisible = 12 # VBA Constant for SpecialCells
    last_col_format_index = 81 # Column CC

    # --- Step 2: Open, Read, and CLOSE the FIASA (Source) file first ---
    if not programacao_fiasa_path:
        q.put(("status", "AVISO: Caminho para 'programacao_fiasa' não fornecido. Pulando esta etapa."))
    else:
        try:
            q.put(("status", f"Abrindo arquivo FIASA para ler dados: {programacao_fiasa_path}"))
            app_programacao_fiasa = xw.App(visible=True, add_book=False) 
            app_programacao_fiasa.display_alerts = False
            app_programacao_fiasa.api.AskToUpdateLinks = False
            # Open as read-only for safety and speed
            wb_programacao_fiasa = app_programacao_fiasa.books.open(programacao_fiasa_path, update_links=False, read_only=True) 

            sheets_wb_programacao_fiasa_PFEP = wb_programacao_fiasa.sheets['PFEP']
            sheets_wb_programacao_fiasa_SUPPLIER = wb_programacao_fiasa.sheets['Suppliers DB']
            filter_sul = 'CARGOLIFT SUL'

            # --- Read FIASA PFEP ---
            q.put(("status", "Lendo dados FIASA PFEP..."))
            sheet_fiasa_pfep = sheets_wb_programacao_fiasa_PFEP
            sheet_fiasa_pfep.api.AutoFilterMode = False
            last_row_f_pfep = sheet_fiasa_pfep.range('A' + str(sheet_fiasa_pfep.cells.last_cell.row)).end('up').row
            
            if last_row_f_pfep > 1:
                filter_range_f_pfep = sheet_fiasa_pfep.range(f'A1:AQ{last_row_f_pfep}') 
                data_range_f_pfep = sheet_fiasa_pfep.range(f'A2:AQ{last_row_f_pfep}') 
                filter_range_f_pfep.api.AutoFilter(Field:=43, Criteria1:=filter_sul)
                try:
                    visible_cells = data_range_f_pfep.api.SpecialCells(xlCellTypeVisible)
                    for area in visible_cells.Areas:
                        values = sheet_fiasa_pfep.range(area.Address).value
                        if isinstance(values, list) and isinstance(values[0], list):
                            data_fiasa_pfep.extend(values)
                        else:
                            data_fiasa_pfep.append(values)
                except Exception:
                    q.put(("status", "Nenhum dado 'CARGOLIFT SUL' em FIASA PFEP."))
                sheet_fiasa_pfep.api.AutoFilterMode = False
            
            # --- Read FIASA SUPPLIER ---
            q.put(("status", "Lendo dados FIASA Suppliers..."))
            sheet_fiasa_sup = sheets_wb_programacao_fiasa_SUPPLIER
            sheet_fiasa_sup.api.AutoFilterMode = False
            last_row_f_sup = sheet_fiasa_sup.range('A' + str(sheet_fiasa_sup.cells.last_cell.row)).end('up').row
            
            if last_row_f_sup > 1:
                filter_range_f_sup = sheet_fiasa_sup.range(f'A1:AB{last_row_f_sup}')
                data_range_f_sup = sheet_fiasa_sup.range(f'A2:AB{last_row_f_sup}')
                filter_range_f_sup.api.AutoFilter(Field:=28, Criteria1:=filter_sul)
                try:
                    visible_cells_sup = data_range_f_sup.api.SpecialCells(xlCellTypeVisible)
                    for area in visible_cells_sup.Areas:
                        values = sheet_fiasa_sup.range(area.Address).value
                        if isinstance(values, list) and isinstance(values[0], list):
                            data_fiasa_sup.extend(values)
                        else:
                            data_fiasa_sup.append(values)
                except Exception:
                    q.put(("status", "Nenhum dado 'CARGOLIFT SUL' em FIASA Suppliers."))
                sheet_fiasa_sup.api.AutoFilterMode = False

        except Exception as e:
            q.put(("status", f"ERRO ao ler arquivo FIASA: {e}"))
        finally:
            # --- IMPORTANT: Close FIASA file *before* opening SUL file ---
            if wb_programacao_fiasa:
                wb_programacao_fiasa.close()
            if app_programacao_fiasa:
                app_programacao_fiasa.quit()
            q.put(("status", "Arquivo Leitura (FIASA) fechado."))

    # --- Step 3: Now that FIASA is closed, open SUL (Destination) file ---
    try:
        q.put(("status", f"Abrindo arquivo SUL: {Programacao_FPT_Sul_path}"))
        app_cargolift_prog_FPT_sul = xw.App(visible=True, add_book=False) 
        app_cargolift_prog_FPT_sul.display_alerts = False
        app_cargolift_prog_FPT_sul.api.AskToUpdateLinks = False
        wb_cargolift_prog_FPT_sul = app_cargolift_prog_FPT_sul.books.open(Programacao_FPT_Sul_path, update_links=False, read_only=False)

        sheet_sup_sul = wb_cargolift_prog_FPT_sul.sheets['Suppliers DB']
        sheet_pfep_sul = wb_cargolift_prog_FPT_sul.sheets['PFEP']
        
        # --- Process Suppliers SUL (Paste FPT Data) ---
        try:
            q.put(("status", "Processando sheet Suppliers SUL (dados FPT)..."))
            sheet_sup_sul.api.Unprotect() 
            q.put(("status", "Limpando sheet Suppliers SUL..."))
            last_row_sup = sheet_sup_sul.cells.last_cell.row
            if last_row_sup > 1:
                last_col_sup = sheet_sup_sul.cells.last_cell.column
                sheet_sup_sul.range((2, 1), (last_row_sup, last_col_sup)).clear_contents()

            if Dado_supplier_a_colar:
                sheet_sup_sul.range('A2').value = Dado_supplier_a_colar
                q.put(("status", f"{len(Dado_supplier_a_colar)} linhas FPT coladas em Suppliers SUL."))
                
                q.put(("status", "Aplicando formatação FPT em Suppliers SUL..."))
                source_format_range_sup = sheet_sup_sul.range((2, 1), (2, last_col_format_index))
                dest_format_range_sup = sheet_sup_sul.range((3, 1), (len(Dado_supplier_a_colar) + 1, last_col_format_index))
                source_format_range_sup.copy()
                dest_format_range_sup.paste(paste='formats')
                app_cargolift_prog_FPT_sul.api.CutCopyMode = False 
            else:
                q.put(("status", "Sem dados FPT SUL para colar em Suppliers."))
        
        except Exception as e:
            q.put(("status", f"ERRO ao colar dados FPT Suppliers: {e}"))

        # --- Process PFEP SUL (Paste FPT Data) ---
        try:
            q.put(("status", "Processando sheet PFEP SUL (dados FPT)..."))
            sheet_pfep_sul.api.Unprotect()
            q.put(("status", "Limpando sheet PFEP SUL..."))
            last_row_pfep = sheet_pfep_sul.cells.last_cell.row
            if last_row_pfep > 1:
                last_col_pfep = sheet_pfep_sul.cells.last_cell.column
                sheet_pfep_sul.range((2, 1), (last_row_pfep, last_col_pfep)).clear_contents()

            if Dado_PFEP_a_colar:
                sheet_pfep_sul.range('A2').value = Dado_PFEP_a_colar
                q.put(("status", f"{len(Dado_PFEP_a_colar)} linhas FPT coladas em PFEP SUL."))

                q.put(("status", "Aplicando formatação FPT em PFEP SUL..."))
                source_format_range_pfep = sheet_pfep_sul.range((2, 1), (2, last_col_format_index))
                dest_format_range_pfep = sheet_pfep_sul.range((3, 1), (len(Dado_PFEP_a_colar) + 1, last_col_format_index))
                source_format_range_pfep.copy()
                dest_format_range_pfep.paste(paste='formats')
                app_cargolift_prog_FPT_sul.api.CutCopyMode = False 
            else:
                q.put(("status", "Sem dados FPT SUL para colar em PFEP."))
        
        except Exception as e:
            q.put(("status", f"ERRO ao colar dados FPT PFEP: {e}"))
            
        # --- Process Suppliers SUL (Paste FIASA Data) ---
        try:
            if data_fiasa_sup: # Use the list we read earlier
                next_row_sup_sul = sheet_sup_sul.range('A' + str(sheet_sup_sul.cells.last_cell.row)).end('up').row + 1
                sheet_sup_sul.range(f'A{next_row_sup_sul}').value = data_fiasa_sup
                q.put(("status", f"{len(data_fiasa_sup)} linhas FIASA coladas em Suppliers SUL."))
                
                q.put(("status", "Aplicando formatação FIASA em Suppliers SUL..."))
                source_format_range_sup = sheet_sup_sul.range((2, 1), (2, last_col_format_index))
                start_row = next_row_sup_sul
                end_row = next_row_sup_sul + len(data_fiasa_sup) - 1
                dest_format_range_sup = sheet_sup_sul.range((start_row, 1), (end_row, last_col_format_index))
                source_format_range_sup.copy()
                dest_format_range_sup.paste(paste='formats')
                app_cargolift_prog_FPT_sul.api.CutCopyMode = False
        except Exception as e:
            q.put(("status", f"ERRO ao colar dados FIASA Suppliers: {e}"))

        # --- Process PFEP SUL (Paste FIASA Data) ---
        try:
            if data_fiasa_pfep: # Use the list we read earlier
                next_row_pfep_sul = sheet_pfep_sul.range('A' + str(sheet_pfep_sul.cells.last_cell.row)).end('up').row + 1
                sheet_pfep_sul.range(f'A{next_row_pfep_sul}').value = data_fiasa_pfep
                q.put(("status", f"{len(data_fiasa_pfep)} linhas FIASA coladas em PFEP SUL."))
                
                q.put(("status", "Aplicando formatação FIASA em PFEP SUL..."))
                source_format_range_pfep = sheet_pfep_sul.range((2, 1), (2, last_col_format_index))
                start_row = next_row_pfep_sul
                end_row = next_row_pfep_sul + len(data_fiasa_pfep) - 1
                dest_format_range_pfep = sheet_pfep_sul.range((start_row, 1), (end_row, last_col_format_index))
                source_format_range_pfep.copy()
                dest_format_range_pfep.paste(paste='formats')
                app_cargolift_prog_FPT_sul.api.CutCopyMode = False
        except Exception as e:
            q.put(("status", f"ERRO ao colar dados FIASA PFEP: {e}"))

        # --- SAVE & CLOSE SUL FILE ---
        q.put(("status", "Salvando arquivo SUL..."))
        wb_cargolift_prog_FPT_sul.save()
        q.put(("status", "Arquivo SUL salvo com sucesso.")) # Your new status
        wb_cargolift_prog_FPT_sul.close()
        wb_cargolift_prog_FPT_sul = None 
        q.put(("status", "Arquivo SUL Fechado!")) # Your new status

    except Exception as e:
        q.put(("status", f"ERRO GERAL ao processar arquivo SUL: {e}"))
        print(f"ERRO GERAL na função SUL: {e}")
        if wb_cargolift_prog_FPT_sul:
            wb_cargolift_prog_FPT_sul.close()
            
    finally:
        # --- Final cleanup for SUL app ---
        if app_cargolift_prog_FPT_sul:
            app_cargolift_prog_FPT_sul.quit()
        q.put(("status", "Processo SUL concluído."))



