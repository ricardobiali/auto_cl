import win32com.client
import subprocess
import psutil
import time

# Caminho para o executável do SAP Logon
SAP_PATH = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"

sapgui = None
App = None
connection = None


# ============================================================
# 🔹 Funções utilitárias de processo SAP
# ============================================================
def is_sap_running():
    """Verifica se o processo saplogon.exe está ativo"""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'saplogon.exe' in proc.info['name'].lower():
            return True
    return False


def open_sap_process_and_wait(timeout=30):
    """Abre o SAP Logon e aguarda até estar ativo"""
    print("Abrindo SAP Logon...")
    subprocess.Popen([SAP_PATH], shell=False)
    for _ in range(timeout):
        if is_sap_running():
            print("SAP Logon iniciado com sucesso.")
            return True
        time.sleep(1)
    raise TimeoutError("SAP Logon não iniciou dentro do tempo esperado.")


def force_close_sap_process():
    """Força o encerramento de todos os processos SAP Logon"""
    print("Encerrando SAP Logon...")
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'saplogon.exe' in proc.info['name'].lower():
            proc.kill()
            print(f"Processo SAP (PID {proc.info['pid']}) encerrado.")


# ============================================================
# 🔹 Funções principais de automação SAP GUI
# ============================================================
def start_sap_manager():
    """Abre o SAP Logon se necessário"""
    if not is_sap_running():
        open_sap_process_and_wait()
        print("SAP Manager iniciado (SAP foi aberto).")
        return False  # indica que o script abriu o SAP
    else:
        print("SAP já estava em execução.")
        return True  # SAP já estava aberto


def close_sap_manager(started_by_script: bool):
    """Fecha SAP apenas se foi iniciado pelo script"""
    if not started_by_script:
        try:
            sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
            sap.Children(0).CloseConnection()
        except Exception:
            pass
        force_close_sap_process()
    print("SAP Manager encerrado.")


def start_connection():
    """Inicia a conexão SAP existente ou abre uma nova"""
    global sapgui, App, connection

    print("Tentando conectar ao SAP GUI Scripting Engine...")

    # Espera até o SAP GUI estar pronto para scripting
    for i in range(30):  # tenta por até ~30 segundos
        try:
            sapgui = win32com.client.GetObject("SAPGUI")
            break
        except Exception:
            time.sleep(1)
    else:
        raise RuntimeError("Não foi possível acessar o SAPGUI via COM (Scripting Engine não disponível).")

    App = sapgui.GetScriptingEngine

    if App.Connections.Count > 0:
        connection = App.Children(0)
        print("Conexão SAP existente detectada.")
        return True
    else:
        print("Abrindo nova conexão SAP...")
        conn_str = (
            '   /SAP_CODEPAGE=1100  /FULLMENU  SNC_PARTNERNAME="p:CN=SAPKERB@PETROBRAS.BIZ" '
            'SNC_QOP=9 /M/pbsap.petrobras.com.br/S/3620/G/SAPSCRIPT /UPDOWNLOAD_CP=2'
        )
        connection = App.OpenConnectionByConnectionString(conn_str, True)
        print("Nova conexão SAP criada.")
        return False

def get_sap_free_session():
    """Obtém uma sessão SAP livre (ou cria uma nova)"""
    print("Procurando sessão livre SAP...")

    global connection

    if connection is None:
        raise RuntimeError("Conexão SAP não inicializada. Chame start_connection() antes.")

    while connection.Sessions.Count >= 6:
        print("Limite de sessões atingido. Aguardando...")
        time.sleep(10)

    main = connection.Children(0)

    while main.Busy:
        time.sleep(1)

    if connection.Sessions.Count == 0:
        main.CreateSession()
        print("Primeira sessão SAP criada.")
    else:
        print("Usando sessão SAP existente.")

    # Pega a sessão mais recente
    ss = connection.Children(connection.Sessions.Count - 1)
    ss.findById("wnd[0]").maximize()
    print(f"Sessão livre obtida: {ss.Id}")
    return ss

def get_sap_session_by_id(session_id: str):
    """Retorna a sessão SAP pelo ID"""
    try:
        session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.findById(session_id)
        print(f"Sessão localizada: {session_id}")
        return session
    except Exception:
        print(f"Erro ao localizar sessão {session_id}.")
        return None


def close_sap_opened_session(session_id: str):
    """Fecha uma sessão SAP específica"""
    try:
        sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        sap.Children(0).CloseSession(session_id)
        print(f"Sessão {session_id} encerrada.")
    except Exception as e:
        print(f"Erro ao fechar sessão {session_id}: {e}")

if __name__ == "__main__":
    started_by_script = start_sap_manager()
    start_connection()
    session = get_sap_free_session()
    print(f"Sessão ativa: {session.Id}")