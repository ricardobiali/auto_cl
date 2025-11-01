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
import win32gui

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
    destinos_dict = None
    status_completa = None
    status_completa_xl = None
    status_reduzida = None

    try:
        # --- Job SAP ---
        if switches.get("report_SAP"):
            completed = subprocess.Popen(
                [sys.executable, "-u", str(YSCLNRCL_PATH)],
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                cwd=str(ROOT_DIR)
            )
            active_processes.append(completed)

            stdout_total = ""
            for line in completed.stdout:
                print(line, end="")
                stdout_total += line
                if line.startswith("DESTINOS_DICT_JSON:"):
                    json_str = line.replace("DESTINOS_DICT_JSON:", "").strip()
                    try:
                        destinos_dict = json.loads(json_str)
                    except json.JSONDecodeError:
                        destinos_dict = None

            completed.wait()
            status_completa = "status_success" if "status_success" in stdout_total else "status_error"
            results.append("Job SAP executado com sucesso.")
        
        # --- Completa XL ---
        if switches.get("completa"):
            completed = subprocess.run(
                [sys.executable, "-u", str(COMPLETA_XL_PATH)],
                capture_output=True, text=True, check=True
            )
            status_completa_xl = "status_success" if "status_success" in completed.stdout else "status_error"
            results.append("Relatório COMPLETA_XL gerado com sucesso.")
        
        # --- Reduzida ---
        if switches.get("reduzida"):
            completed = subprocess.run(
                [sys.executable, "-u", str(REDUZIDA_PATH)],
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

        if destinos_dict:
            existing_data["destino"] = destinos_dict.get("destino", [])

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
    global tk_root
    try:
        if tk_root is None:
            tk_root = tk.Tk()
            tk_root.withdraw()

        try:
            hwnd_main = win32gui.GetForegroundWindow()
        except Exception:
            hwnd_main = None

        if hwnd_main:
            try:
                tk_root.wm_attributes("-toolwindow", True)
                tk_root.wm_attributes("-topmost", True)
                tk_root.lift()
                tk_root.focus_force()
                ctypes.windll.user32.SetWindowLongW(
                    tk_root.winfo_id(),
                    -8,
                    hwnd_main
                )
            except Exception as e:
                print("Aviso: não foi possível vincular janela Tk ao app principal:", e)

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
    global tk_root
    try:
        if tk_root is None:
            tk_root = tk.Tk()
            tk_root.withdraw()

        try:
            hwnd_main = win32gui.GetForegroundWindow()
        except Exception:
            hwnd_main = None

        if hwnd_main:
            try:
                tk_root.wm_attributes("-toolwindow", True)
                tk_root.wm_attributes("-topmost", True)
                tk_root.lift()
                tk_root.focus_force()
                ctypes.windll.user32.SetWindowLongW(
                    tk_root.winfo_id(),
                    -8,
                    hwnd_main
                )
            except Exception as e:
                print("Aviso: não foi possível vincular janela Tk ao app principal:", e)

        arquivo_selecionado = filedialog.askopenfilename(
            parent=tk_root,
            title="Selecione um diretório de armazenamento"
        )
        print("selecionar_arquivo ->", arquivo_selecionado)
        return arquivo_selecionado if arquivo_selecionado else ""

    except Exception as e:
        print("Erro ao abrir diálogo de pasta:", e)
        return ""

@eel.expose
def start_job(switches, paths):
    if job_status["running"]:
        return {"status": "already_running"}
    
    if switches.get("completa") and not switches.get("report_SAP"):
        if not paths.get("file_completa"):
            file_completa = selecionar_arquivo()
            if not file_completa:
                return {"status": "cancelled"}
            paths["file_completa"] = file_completa
    
    if switches.get("reduzida") and not switches.get("report_SAP") and not switches.get("completa"):
        if not paths.get("file_reduzida"):
            file_reduzida = selecionar_arquivo()
            if not file_reduzida:
                return {"status": "cancelled"}
            paths["file_reduzida"] = file_reduzida

    threading.Thread(target=_run_sequenced_job, args=(switches, paths), daemon=True).start()
    return {"status": "started"}

def _run_sequenced_job(switches, paths):
    global job_status

    job_status["running"] = True
    job_status["success"] = None
    job_status["message"] = "Executando automação..."

    destino_final = None
    file_completa = None
    file_reduzida = None

    # 1️⃣ Switch1 - SAP
    if switches.get("report_SAP"):
        run_job({"report_SAP": True}, paths)
        if os.path.exists(REQUESTS_PATH):
            with open(REQUESTS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                destinos = data.get("destino", [])
                if destinos:
                    destino_final = destinos[0] if isinstance(destinos, list) else destinos

    # 2️⃣ Switch2 - Completa
    if switches.get("completa"):
        if destino_final:
            file_completa = destino_final
        elif paths.get("file_completa"):
            file_completa = paths.get("file_completa")
        else:
            job_status.update({
                "running": False,
                "success": False,
                "message": "Execução abortada: destino_final não retornado pelo SAP."
            })
            return

        run_job({"completa": True}, paths)

        # Recarrega requests.json e atualiza destinos_dict
        if os.path.exists(REQUESTS_PATH):
            with open(REQUESTS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            destinos = data.get("destino", [])
            if destinos:
                destino_final = destinos[0] if isinstance(destinos, list) else destinos

    # 3️⃣ Switch3 - Reduzida
    if switches.get("reduzida"):
        if switches.get("report_SAP") and destino_final:
            file_reduzida = destino_final
        elif switches.get("completa") and file_completa:
            file_reduzida = file_completa
        else:
            if destino_final:
                file_reduzida = destino_final
            else:
                if paths.get("file_reduzida"):
                    file_reduzida = paths.get("file_reduzida")
                else:
                    job_status.update({
                        "running": False,
                        "success": False,
                        "message": "Execução cancelada: nenhum arquivo selecionado para Reduzida."
                    })
                    return

        run_job({"reduzida": True}, paths)

        # Atualiza destinos_dict no requests.json para Reduzida
        if os.path.exists(REQUESTS_PATH):
            with open(REQUESTS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            destinos = data.get("destino", [])
            if destinos:
                destino_final = destinos[0] if isinstance(destinos, list) else destinos

    job_status.update({"running": False, "success": True, "message": "Jobs concluídos em sequência."})

@eel.expose
def get_job_status():
    return job_status

@eel.expose
def cancel_job():
    global job_status, active_processes
    try:
        if active_processes:
            for proc in active_processes:
                if proc and proc.poll() is None:
                    proc.terminate()
            active_processes = []

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
    with open(REQUESTS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    return {"status": "saved"}

@eel.expose
def read_requests_json():
    if os.path.exists(REQUESTS_PATH):
        with open(REQUESTS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"paths": [], "requests": [], "status": [{}]}

@eel.expose
def write_requests_json(data):
    with open(REQUESTS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    return {"status": "saved"}

@eel.expose
def get_welcome_name():
    return user_data.full_name


# -----------------------------
# Inicializa app desktop ou depuração local
# -----------------------------
if __name__ == "__main__":
    # ⚙️ Modo de depuração local no VS Code
    if "--debug" in sys.argv:
        print("\n=== MODO DEBUG ATIVADO ===")
        switches = {"report_SAP": True, "completa": True, "reduzida": False}
        paths = {"file_completa": "", "file_reduzida": ""}
        _run_sequenced_job(switches, paths)
        print("\n=== EXECUÇÃO DEBUG CONCLUÍDA ===\n")
    else:
        # ⚙️ Modo normal (frontend Eel)
        eel.start(
            "index.html",
            port=8000,
            size=(1200, 800),
            cmdline_args=['--start-maximized']
        )