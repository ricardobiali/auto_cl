import os
import sys
import shutil
import glob
import time
import json
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

# Caminho do requests.json
requests_path = os.path.join(
    fr"C:\Users\{username}\OneDrive - PETROBRAS\Desktop\python\auto_cl_prototype\frontend\framework",
    "requests.json"
)

# Lê o arquivo requests.json e extrai dados do primeiro item
if os.path.exists(requests_path):
    with open(requests_path, "r", encoding="utf-8") as f:
        data = json.load(f)
        if isinstance(data, list) and len(data) > 0:
            first = data[0]
            defprojeto = first.get("defprojeto", "").strip()
            fase = first.get("fase", "").strip()
            status = first.get("status", "").strip()
            datainicio = first.get("datainicio", "").strip()
            exercicio = first.get("exercicio", "").strip()
            trimestre = first.get("trimestre", "").strip()

            # 🗓️ Converte ddmmaaaa → aaaammdd
            if len(datainicio) == 8 and datainicio.isdigit():
                datainicio = datainicio[4:] + datainicio[2:4] + datainicio[:2]
            else:
                print(f"⚠️ Formato inesperado de datainicio: {datainicio}")

        else:
            print("⚠️ Nenhum registro encontrado no requests.json, usando valores padrão.")
            defprojeto = fase = status = datainicio = exercicio = trimestre = "DEFAULT"
else:
    print(f"⚠️ Arquivo requests.json não encontrado em {requests_path}, usando valores padrão.")
    defprojeto = fase = status = datainicio = exercicio = trimestre = "DEFAULT"

# 📅 Data corrente no formato aaaammdd
datacorrente = datetime.now().strftime("%Y%m%d")

# Padrão dinâmico de arquivo (inclui datacorrente)
padrao = f"RGT_RCL.CSV_{username}_{defprojeto}_{fase}_{status}_{datainicio}_{exercicio}_{trimestre}T_{datacorrente}_*.txt"

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
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Arquivo ainda não encontrado... tentando novamente em {intervalo_busca} segundos.")
        time.sleep(intervalo_busca)
