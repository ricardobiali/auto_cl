import os
import sys
import shutil
import glob
import time
from datetime import datetime

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
from backend.sap_manager.sap_connect import get_sap_free_session, start_sap_manager, start_connection

# Inicializa SAP
started_by_script = start_sap_manager()
start_connection()
session = get_sap_free_session()

# Nome usuário Windows
username = os.getlogin()

# Caminhos de origem e destino
origem = fr"C:\Users\{username}\PETROBRAS\GPP-E&P RXC GDI - Conteúdo Local\RGIT"
destino = fr"C:\Users\{username}\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 0 - Arquivos de Texto do SAP"

# Padrão do arquivo a localizar
padrao = "RGT_RCL.CSV_U33V_JV3A5118530_D__20240101_2024_1T*.txt"

# Intervalo entre verificações (em segundos)
intervalo_busca = 120

# --- Abre SM37 e marca PRELIM ---
session.findById("wnd[0]/tbar[0]/okcd").text = "/nsm37"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/chkBTCH2170-PRELIM").selected = True
session.findById("wnd[0]/tbar[1]/btn[8]").press()

print(f"🔍 Iniciando monitoramento da pasta:\n   {origem}")
print(f"Aguardando arquivo com padrão: {padrao}\n")

while True:
    # Verifica se o arquivo já chegou
    arquivos = glob.glob(os.path.join(origem, padrao))
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    if arquivos:
        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)
            destino_final = os.path.join(destino, nome_arquivo)
            try:
                shutil.move(arquivo, destino_final)
                print(f"\n✅ [{datetime.now().strftime('%H:%M:%S')}] Arquivo encontrado e movido com sucesso:")
                print(f"   ➜ {nome_arquivo}")
                print(f"   ➜ De: {origem}")
                print(f"   ➜ Para: {destino_final}")
                print("\nEncerrando monitoramento.")
                exit(0)
            except Exception as e:
                print(f"⚠️ Erro ao mover {nome_arquivo}: {e}")
                time.sleep(intervalo_busca)
    else:
        # Mantém sessão ativa na SM37 (opcional: você pode repetir algum refresh se quiser)
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Arquivo ainda não encontrado... tentando novamente em {intervalo_busca} segundos.")
        time.sleep(intervalo_busca)
