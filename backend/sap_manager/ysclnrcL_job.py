import glob
from pathlib import Path
import shutil
import sys
import time
from sap_connect import get_sap_free_session, start_sap_manager, start_connection, close_sap_manager
from datetime import datetime, timedelta
import os
import json

# Caminho atual do script
current_dir = Path(__file__).resolve()
username = os.getlogin()

# Sobe até encontrar a pasta 'auto_cl_prototype'
root_dir = current_dir
while root_dir.name != "auto_cl_prototype":
    if root_dir.parent == root_dir:
        raise FileNotFoundError("Pasta 'auto_cl_prototype' não encontrada.")
    root_dir = root_dir.parent

json_path = fr"{root_dir}\frontend\requests.json"

def create_YSCLBLRIT_requests(session, init_date=None, init_time=None, interval=None, requests_data=None):
    """
    Executa a transação YSCLNRCL e cria requisições de forma automatizada.
    """

    for i, req in enumerate(requests_data, start=1):
        print(f"Processando requisição {i}...")

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nYSCLNRCL"
        session.findById("wnd[0]").sendVKey(0)

        # --- Preenche campos principais ---
        session.findById("wnd[0]/usr/ctxtP_BUK_N").text = req.get("empresa", "")
        session.findById("wnd[0]/usr/txtPC_ANO").text = req.get("exercicio", "")
        session.findById("wnd[0]/usr/cmbTRI").key = req.get("trimestre", "1")
        session.findById("wnd[0]/usr/ctxtPC_CODCB").text = req.get("campo", "")
        session.findById("wnd[0]/usr/ctxtPC_FASE").text = req.get("fase", "")
        session.findById("wnd[0]/usr/ctxtPC_STAT").text = req.get("status", "")
        session.findById("wnd[0]/usr/ctxtP_VERSAO").text = req.get("versao", "")
        session.findById("wnd[0]/usr/ctxtP_SECAO").text = req.get("secao", "")

        # --- Abre filtro avançado ---
        session.findById("wnd[0]/tbar[1]/btn[19]").press()
        session.findById("wnd[0]/usr/ctxtSC_PSPID-LOW").text = req.get("defprojeto", "")
        session.findById("wnd[0]/usr/ctxtSD_DTINI-LOW").text = req.get("datainicio", "")
        session.findById("wnd[0]/usr/ctxtPC_BID").text = req.get("bidround", "")

        session.findById("wnd[0]/usr/chkP_PART").selected = True

        # --- Executa e abre tela de agendamento ---
        session.findById("wnd[0]/mbar/menu[0]/menu[2]").select()
        session.findById("wnd[1]/tbar[0]/btn[13]").press()
        session.findById("wnd[1]/usr/btnDATE_PUSH").press()

        # --- Calcula data/hora do agendamento ---
        job_init_datetime = datetime.now() + timedelta(minutes=1)
        str_date_plan = job_init_datetime.strftime("%d.%m.%Y")
        str_time_plan = job_init_datetime.strftime("%H:%M:%S")

        # --- Preenche no SAP ---
        session.findById("wnd[1]/usr/ctxtBTCH1010-SDLSTRTDT").text = str_date_plan
        session.findById("wnd[1]/usr/ctxtBTCH1010-SDLSTRTTM").text = str_time_plan
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        print(f"Requisição {i} agendada para {str_date_plan} às {str_time_plan}")

# Execução principal
if __name__ == "__main__":
    try:
        started_by_script = start_sap_manager()
        start_connection()
        session = get_sap_free_session()

        # --- Lê requests.json ---
        requests_data = []
        path_data = []

        if os.path.exists(json_path):
            with open(json_path, "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                    if isinstance(data, dict):
                        requests_data = data.get("requests", [])
                        path_data = data.get("paths", [])
                except json.JSONDecodeError as e:
                    print("Erro ao ler requests.json:", e)

        # --- Cria requisições ---
        create_YSCLBLRIT_requests(
            session=session,
            init_date="01.01.2011",
            init_time="08:00",
            interval=15,
            requests_data=requests_data
        )

        try:
            # # Caminho do requests.json
            requests_data = os.path.join(
                fr"{root_dir}\frontend",
                "requests.json"
            )

            # Valores padrão
            defprojeto = fase = status = datainicio = exercicio = trimestre = path1 = "DEFAULT"

            # Lê o arquivo requests.json e extrai dados do primeiro item
            if os.path.exists(requests_data):
                with open(requests_data, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    requests_list = data.get("requests", [])
                    paths_list = data.get("paths", [])

                    if requests_list and isinstance(requests_list, list):
                        first = requests_list[0]
                        defprojeto = first.get("defprojeto", "").strip()
                        fase = first.get("fase", "").strip()
                        status = first.get("status", "").strip()
                        datainicio = first.get("datainicio", "").strip()
                        exercicio = first.get("exercicio", "").strip()
                        trimestre = first.get("trimestre", "").strip()

                        # Converte ddmmaaaa → aaaammdd
                        if len(datainicio) == 8 and datainicio.isdigit():
                            datainicio = datainicio[4:] + datainicio[2:4] + datainicio[:2]
                        else:
                            print(f"Formato inesperado de datainicio: {datainicio}")

                    else:
                        print("Nenhum registro em 'requests', usando valores padrão.")

                    # Lê path1 do primeiro item em 'paths', se existir
                    if paths_list and isinstance(paths_list, list):
                        path1 = paths_list[0].get("path1", "").strip()
                        if not path1:
                            print("'path1' vazio no requests.json, usando padrão.")
                    else:
                        print("Nenhum registro em 'paths', usando padrão.")
            else:
                print(f"Arquivo requests.json não encontrado em {requests_data}, usando valores padrão.")

            # Caminhos de origem e destino
            origem = fr"C:\Users\{username}\PETROBRAS\GPP-E&P RXC GDI - Conteúdo Local\RGIT"
            destino = path1  # Agora usa path1 em vez de caminho fixo

            # Data corrente no formato aaaammdd
            datacorrente = datetime.now().strftime("%Y%m%d")

            # --- Geração dos padrões de arquivo para todas as requisições ---
            padroes = []
            datacorrente = datetime.now().strftime("%Y%m%d")

            if requests_list and isinstance(requests_list, list):
                for req in requests_list:
                    defprojeto = req.get("defprojeto", "").strip()
                    fase = req.get("fase", "").strip()
                    status = req.get("status", "").strip()
                    datainicio = req.get("datainicio", "").strip()
                    exercicio = req.get("exercicio", "").strip()
                    trimestre = req.get("trimestre", "").strip()

                    # Converte ddmmaaaa → aaaammdd
                    if len(datainicio) == 8 and datainicio.isdigit():
                        datainicio = datainicio[4:] + datainicio[2:4] + datainicio[:2]

                    padrao = f"RGT_RCL.CSV_{username}_{defprojeto}_{fase}_{status}_{datainicio}_{exercicio}_{trimestre}T_{datacorrente}_*.txt"
                    padroes.append(padrao)
            else:
                print("Nenhum request válido encontrado para montar padrões, usando padrão único.")
                padroes = [f"RGT_RCL.CSV_{username}_DEFAULT_DEFAULT_DEFAULT_DEFAULT_DEFAULT_DEFAULTT_{datacorrente}_*.txt"]

            intervalo_busca = 120

            # --- Abre SM37 e marca PRELIM ---
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nsm37"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/chkBTCH2170-PRELIM").selected = True
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            print(f"Iniciando monitoramento da pasta:\n   {origem}")
            print("Aguardando arquivos:")
            for p in padroes:
                print(f"   - {p}")
            print()

            encontrados = set()

            while True:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                for padrao in padroes:
                    arquivos = glob.glob(os.path.join(origem, padrao))
                    if arquivos:
                        for arquivo in arquivos:
                            nome_arquivo = os.path.basename(arquivo)
                            destino_final = os.path.join(destino, nome_arquivo)
                            try:
                                shutil.move(arquivo, destino_final)
                                print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Arquivo encontrado e movido com sucesso:")
                                print(f"   De: {origem}")
                                print(f"DESTINO_FINAL: {destino_final}")
                                encontrados.add(padrao)
                            except Exception as e:
                                import traceback
                                print(f"Erro ao mover {nome_arquivo}: {e}", flush=True)
                                print(traceback.format_exc(), flush=True)
                                status_done = "status_error"
                                os._exit(0)

                # --- Verifica se todos os padrões já foram encontrados ---
                if len(encontrados) == len(padroes):
                    print("\nTodos os arquivos foram encontrados e movidos com sucesso.")
                    print("Encerrando monitoramento.")
                    status_done = "status_success"
                    os._exit(0)

                print(f"[{datetime.now().strftime('%H:%M:%S')}] Ainda aguardando {len(padroes) - len(encontrados)} arquivo(s)...")
                time.sleep(intervalo_busca)


        except Exception as e:
            print(f"Erro geral durante a execução: {e}")
            status_done = "status_error"

    except Exception as e:
        print("Ocorreu um erro na execução do ysclnrcl_job.py:", e)
        status_done = "status_error"