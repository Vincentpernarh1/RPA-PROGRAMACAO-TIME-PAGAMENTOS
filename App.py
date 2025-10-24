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



from Tasks import Processar_Demandas
from Tasks import download_Demanda


warnings.filterwarnings("ignore", category=UserWarning)
import sys

def get_playwright_browser_path():
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
        chromium_path = os.path.join(base_path, "ms-playwright", "chromium-1187", "chrome-win", "chrome.exe")
    else:
        base_path = r"C:\Users\perna\AppData\Local"

        # Join the rest of the Playwright folder path
        chromium_path = os.path.join(
            base_path,
            "ms-playwright",
            "chromium-1187",
            "chrome-win",
            "chrome.exe"
        )
   
    if chromium_path and not os.path.exists(chromium_path):
        raise FileNotFoundError(f"Chromium executable not found at {chromium_path}")

    return chromium_path


# --- GUI UPDATE FUNCTION ---
def update_gui(queue_instance, status_label, progress_bar, log_text):
    """Checks the queue for messages from the worker thread and updates the GUI."""
    try:
        while True:
            message_type, value = queue_instance.get_nowait()
            if message_type == "status":
                status_label.config(text=value)
                log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {value}\n")
                log_text.see(tk.END)
            elif message_type == "progress":
                progress_bar['value'] = value
            elif message_type == "done":
                status_label.config(text="Processo Concluído!")
                progress_bar['value'] = 100
                return # Stop checking
    except queue.Empty:
        pass
    status_label.after(100, lambda: update_gui(queue_instance, status_label, progress_bar, log_text))


def load_credentials():
    """Loads Credencial.json from the same directory as the running script or executable."""
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    cred_path = os.path.join(base_path, "credencial.json")

    if not os.path.exists(cred_path):
        raise FileNotFoundError(f"Credencial.json not found in: {cred_path}")

    with open(cred_path, "r", encoding="utf-8") as f:
        return json.load(f)
    

def load_modelos():
    """Loads Credencial.json from the same directory as the running script or executable."""
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    Model_path = os.path.join(base_path, "Modelos.json")

    if not os.path.exists(Model_path):
        raise FileNotFoundError(f"Credencial.json not found in: {Model_path}")

    with open(Model_path, "r", encoding="utf-8") as f:
        return json.load(f)










    
def run_automation(playwright: Playwright, q: queue.Queue):

    # Processar_Demandas()

    ecr_path, odm_path = None, None
    try:
        # 1. Load Credentials
        q.put(("status", "Carregando credenciais..."))
        q.put(("progress", 5))
        credentials = load_credentials()
        url_order, username, password = credentials['url_Elog'], credentials['user'], credentials['password']

        # 2. Launch Browser
        q.put(("status", "Iniciando navegador..."))

        chromium_path = get_playwright_browser_path()
        
        if chromium_path:
            # .exe → use bundled Chromium
            browser = playwright.chromium.launch(
                headless=True,
                executable_path=chromium_path,
                args=["--start-maximized"]
            )
        else:
            # .py → use default Playwright Chromium
            browser = playwright.chromium.launch(
                headless=True,
                args=["--start-maximized"]
            )
                    
        # context = browser.new_context(viewport={'width': 1920, 'height': 1080})
        context = browser.new_context(no_viewport=True)
        page = context.new_page()
         
         
       

        download_Demanda(page,url_order,q,username,password)
       




       

    except FileNotFoundError:
        q.put(("status", "Erro: 'Credencial.json' não encontrado."))
    except KeyError:
        q.put(("status", "Erro: JSON de credenciais inválido."))
    except TimeoutError:
        q.put(("status", "Erro de Timeout: Verifique os seletores ou a conexão."))
        page.screenshot(path="login_error.png")
    except Exception as e:
        q.put(("status", f"Ocorreu um erro inesperado: {e}"))
    finally:
        # 5. Clean Up and next step
        q.put(("status", "Fechando navegador..."))
        if 'context' in locals(): context.close()
        if 'browser' in locals(): browser.close()
        q.put(("done", True))


def main_process(q: queue.Queue):
    with sync_playwright() as playwright:
        run_automation(playwright, q)

# --- TKINTER APP SETUP ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Ferramenta de Automação e Processamento")
        self.root.geometry("600x400")

        self.queue = queue.Queue()

        # --- Widgets ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(main_frame, text="Pronto para iniciar. Clique em 'Processar'.", font=("Helvetica", 12))
        self.status_label.pack(pady=5, padx=5, fill=tk.X)

        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=10, padx=5, fill=tk.X)

        self.process_button = ttk.Button(main_frame, text="Processar", command=self.start_processing_thread)
        self.process_button.pack(pady=10)
        
        log_frame = ttk.LabelFrame(main_frame, text="Log de Atividades", padding="10")
        log_frame.pack(pady=10, padx=5, fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=70, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def start_processing_thread(self):
        self.process_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.log_text.delete('1.0', tk.END)
        self.status_label.config(text="Iniciando processo...")
        
        self.thread = threading.Thread(target=main_process, args=(self.queue,))
        self.thread.daemon = True
        self.thread.start()
        
        # Start checking the queue for updates
        update_gui(self.queue, self.status_label, self.progress_bar, self.log_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
