import sys
from pathlib import Path
import eel
import subprocess
import json
import os
import threading
from frontend import user_data
import tkinter as tk
from tkinter import filedialog
import ctypes
import psutil

active_processes = []

# Configurações de caminhos
BASE_DIR = Path(__file__).parent / "frontend"
REQUESTS_PATH = BASE_DIR / "requests.json"

ROOT_DIR = Path(__file__).parent
YSCLNRCL_PATH = ROOT_DIR / "backend/sap_manager/ysclnrcl_job.py"
COMPLETA_XL_PATH = ROOT_DIR / "backend/reports/completa_xl.py"
REDUZIDA_PATH = ROOT_DIR / "backend/reports/reduzida.py"

# -----------------------------
# Status global do job
# -----------------------------
job_status = {
    "running": False,
    "success": None,
    "message": ""
}

# -----------------------------
# Inicializa Eel
# -----------------------------
eel.init(str(BASE_DIR))

# -----------------------------
# Função para rodar o job Python
# -----------------------------
def run_job(switches: dict, paths: dict):
    global job_status
    job_status["running"] = True
    job_status["success"] = None
    job_status["message"] = "Executando automação..."
    
    results = []

    try:
        # --- Job SAP ---
        if switches.get("report_SAP"):
            completed = subprocess.Popen(
                [sys.executable, str(YSCLNRCL_PATH)],
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT
            )
            active_processes.append(completed)

            stdout_total = ""
            for line in completed.stdout:
                print(line, end="")  # imprime em tempo real
                stdout_total += line

            completed.wait()
            status_completa = "status_success" if "status_success" in stdout_total else "status_error"
            results.append("Job SAP executado com sucesso.")
        
        # --- Completa XL ---
        if switches.get("completa"):
            completed = subprocess.run(
                [sys.executable, str(COMPLETA_XL_PATH)],
                capture_output=True, text=True, check=True
            )
            status_completa_xl = "status_success" if "status_success" in completed.stdout else "status_error"
            results.append("Relatório COMPLETA_XL gerado com sucesso.")
        
        # --- Reduzida ---
        if switches.get("reduzida"):
            completed = subprocess.run(
                [sys.executable, str(REDUZIDA_PATH)],
                capture_output=True, text=True, check=True
            )
            status_reduzida = "status_success" if "status_success" in completed.stdout else "status_error"
            results.append("Relatório REDUZIDA gerado com sucesso.")

        # Atualiza requests.json
        if os.path.exists(REQUESTS_PATH):
            with open(REQUESTS_PATH, "r", encoding="utf-8") as f:
                existing_data = json.load(f)
        else:
            existing_data = {"paths": [], "requests": [], "status": [{}]}

        status_block = existing_data.get("status", [{}])[0]
        if switches.get("report_SAP"): status_block["ysclnrcl_job.py"] = status_completa
        if switches.get("completa"): status_block["completa_xl.py"] = status_completa_xl
        if switches.get("reduzida"): status_block["reduzida.py"] = status_reduzida
        existing_data["status"] = [status_block]

        with open(REQUESTS_PATH, "w", encoding="utf-8") as f:
            json.dump(existing_data, f, indent=4, ensure_ascii=False)

        job_status["success"] = True
        job_status["message"] = " | ".join(results)

    except subprocess.CalledProcessError as e:
        job_status["success"] = False
        job_status["message"] = f"Erro na execução do job:\n{e.stderr}"

    finally:
        job_status["running"] = False

# -----------------------------
# Funções expostas para JS
# -----------------------------
tk_root = None
@eel.expose
def selecionar_diretorio():
    """Abre um diálogo para selecionar pasta e retorna o caminho"""
    global tk_root
    try:
        if tk_root is None:
            tk_root = tk.Tk()
            tk_root.withdraw()  # cria apenas uma vez

        # --- Obtém handle (HWND) da janela principal Eel (janela Chromium) ---
        try:
            import win32gui
            hwnd_main = win32gui.GetForegroundWindow()  # pega a janela ativa
        except Exception:
            hwnd_main = None

        # --- Cria uma janela temporária associada ao HWND da aplicação principal ---
        if hwnd_main:
            try:
                tk_root.wm_attributes("-toolwindow", True)
                tk_root.wm_attributes("-topmost", True)
                tk_root.lift()
                tk_root.focus_force()

                # Vincula a janela Tk à janela principal (modal)
                ctypes.windll.user32.SetWindowLongW(
                    tk_root.winfo_id(),
                    -8,  # GWL_HWNDPARENT
                    hwnd_main
                )
            except Exception as e:
                print("Aviso: não foi possível vincular janela Tk ao app principal:", e)

        # --- Exibe o diálogo ---
        folder_selected = filedialog.askdirectory(
            parent=tk_root,
            title="Selecione um diretório de armazenamento"
        )
        print("selecionar_diretorio ->", folder_selected)
        return folder_selected if folder_selected else ""

    except Exception as e:
        print("Erro ao abrir diálogo de pasta:", e)
        return ""

@eel.expose
def selecionar_arquivo():
    """Abre um diálogo para selecionar arquivo e retorna o caminho completo"""
    global tk_root
    try:
        if tk_root is None:
            tk_root = tk.Tk()
            tk_root.withdraw()

        # --- Obtém handle (HWND) da janela principal Eel (janela Chromium) ---
        try:
            import win32gui
            hwnd_main = win32gui.GetForegroundWindow()  # pega a janela ativa
        except Exception:
            hwnd_main = None

        # --- Cria uma janela temporária associada ao HWND da aplicação principal ---
        if hwnd_main:
            try:
                tk_root.wm_attributes("-toolwindow", True)
                tk_root.wm_attributes("-topmost", True)
                tk_root.lift()
                tk_root.focus_force()

                # Vincula a janela Tk à janela principal (modal)
                ctypes.windll.user32.SetWindowLongW(
                    tk_root.winfo_id(),
                    -8,  # GWL_HWNDPARENT
                    hwnd_main
                )
            except Exception as e:
                print("Aviso: não foi possível vincular janela Tk ao app principal:", e)

        # --- Exibe o diálogo ---
        arquivo_selecionado = filedialog.askopenfilename(
            parent=tk_root,
            title="Selecione um diretório de armazenamento"
        )
        print("selecionar_diretorio ->", arquivo_selecionado)
        return arquivo_selecionado if arquivo_selecionado else ""

    except Exception as e:
        print("Erro ao abrir diálogo de pasta:", e)
        return ""

@eel.expose
def start_job(switches, paths):
    if job_status["running"]:
        return {"status": "already_running"}

    # Se switch2 (completa) estiver ativo, pede arquivo antes de rodar
    if switches.get("completa"):
        file_path = selecionar_arquivo()
        if not file_path:
            return {"status": "cancelled", "message": "Execução cancelada pelo usuário."}

        # Atualiza requests.json com o campo "file_completa"
        if os.path.exists(REQUESTS_PATH):
            with open(REQUESTS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        else:
            data = {"paths": [], "requests": [], "switches": switches, "status": [{}]}

        data["file_completa"] = file_path

        with open(REQUESTS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    
    # Se switch3 (reduzida) estiver ativo
    if switches.get("reduzida"):
        file_path = selecionar_arquivo()
        if not file_path:
            return {"status": "cancelled", "message": "Execução cancelada pelo usuário."}
        if os.path.exists(REQUESTS_PATH):
            with open(REQUESTS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        else:
            data = {"paths": [], "requests": [], "switches": switches, "status": [{}]}
        data["file_reduzida"] = file_path
        with open(REQUESTS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

    # Inicia o job em thread
    threading.Thread(target=run_job, args=(switches, paths), daemon=True).start()
    return {"status": "started"}

@eel.expose
def get_job_status():
    return job_status

@eel.expose
def cancel_job():
    """Cancela o job em execução e força recarregamento da interface"""
    global job_status, active_processes

    try:
        if active_processes:
            for proc in active_processes:
                if proc and proc.poll() is None:
                    proc.terminate()
            active_processes = []  # limpa a lista

        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and proc.info['name'].lower() in ['saplogon.exe', 'sapgui.exe']:
                    print(f"Encerrando processo SAP: {proc.info['name']} (PID {proc.info['pid']})")
                    proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        job_status["running"] = False
        job_status["success"] = None
        job_status["message"] = "Job cancelado manualmente."

        print("Job cancelado manualmente (SAP encerrado, se aberto).")

    except Exception as e:
        print("Aviso: não foi possível encerrar subprocessos:", e)

@eel.expose
def save_requests(data):
    """Grava requests.json via front-end"""
    with open(REQUESTS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    return {"status": "saved"}

@eel.expose
def read_requests_json():
    """Lê requests.json e retorna como dict"""
    if os.path.exists(REQUESTS_PATH):
        with open(REQUESTS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"paths": [], "requests": [], "status": [{}]}

@eel.expose
def write_requests_json(data):
    """Grava requests.json diretamente"""
    with open(REQUESTS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    return {"status": "saved"}

@eel.expose
def get_welcome_name():
    return user_data.full_name

# -----------------------------
# Inicializa app desktop
# -----------------------------
if __name__ == "__main__":
    eel.start(
        "index.html",
        port=8000,
        size=(1200, 800),
        cmdline_args=['--start-maximized']  # abre maximizado
    )