import glob
from pathlib import Path
import shutil
import sys
import time
from datetime import datetime, timedelta
import os
import json

username = os.getlogin()

try:
    # repo_root = .../auto_cl
    repo_root = Path(__file__).resolve().parents[2]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
except Exception:
    pass

from backend.sap_manager.sap_connect import (
    get_sap_free_session,
    start_sap_manager,
    start_connection,
    close_sap_manager,
)

from app.paths import Paths
from app.services.file_io import load_json

P = Paths.build()
requests_path = P.requests_json  


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
        visao_corp = bool(req.get("visao_corp", False))
        session.findById("wnd[0]/usr/chkP_CORP").selected = visao_corp

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

        # ✅ aborta com mensagem se requests.json não existir
        if not requests_path.exists():
            print(f"ERRO: requests.json não encontrado em: {requests_path}")
            print("Dica: rode a interface e clique em Executar (ela cria o requests.json antes do job).")
            sys.exit(1)

        data = load_json(requests_path)
        if not isinstance(data, dict):
            print(f"ERRO: requests.json inválido (não é um objeto JSON). Caminho: {requests_path}")
            sys.exit(1)

        requests_list = data.get("requests", [])
        paths_list = data.get("paths", [])

        if not requests_list or not isinstance(requests_list, list):
            print(f"ERRO: requests.json não contém 'requests' válido (lista). Caminho: {requests_path}")
            sys.exit(1)

        if not paths_list or not isinstance(paths_list, list):
            print(f"ERRO: requests.json não contém 'paths' válido (lista). Caminho: {requests_path}")
            sys.exit(1)

        # --- Cria requisições ---
        create_YSCLBLRIT_requests(
            session=session,
            init_date="01.01.2011",
            init_time="08:00",
            interval=15,
            requests_data=requests_list
        )

        try:
            # Valores padrão
            defprojeto = fase = status = datainicio = exercicio = trimestre = path1 = "DEFAULT"

            # Lê dados do primeiro item
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

            # path1
            path1 = paths_list[0].get("path1", "").strip()
            if not path1:
                print("ERRO: 'path1' vazio no requests.json.")
                sys.exit(1)

            destino = path1
            datacorrente = datetime.now().strftime("%Y%m%d")

            # --- Geração dos padrões de arquivo para todas as requisições ---
            padroes = []
            origens_por_padrao = []

            for req in requests_list:
                defprojeto = req.get("defprojeto", "").strip()
                fase = req.get("fase", "").strip()
                status = req.get("status", "").strip()
                datainicio = req.get("datainicio", "").strip()
                exercicio = req.get("exercicio", "").strip()
                trimestre = req.get("trimestre", "").strip()
                rit_flag = req.get("rit", False)

                if len(datainicio) == 8 and datainicio.isdigit():
                    datainicio = datainicio[4:] + datainicio[2:4] + datainicio[:2]

                if rit_flag:
                    origem = fr"C:\Users\{username}\PETROBRAS\GPP-E&P RXC GDI - Conteúdo Local\RIT"
                    padrao = f"_RCL.CSV_{username}_{defprojeto}_{fase}_{status}_{datainicio}_{exercicio}_{trimestre}T_{datacorrente}_*.txt"
                else:
                    origem = fr"C:\Users\{username}\PETROBRAS\GPP-E&P RXC GDI - Conteúdo Local\RGIT"
                    padrao = f"RGT_RCL.CSV_{username}_{defprojeto}_{fase}_{status}_{datainicio}_{exercicio}_{trimestre}T_{datacorrente}_*.txt"

                padroes.append(padrao)
                origens_por_padrao.append(origem)

            intervalo_busca = 120

            # --- Abre SM37 e marca PRELIM ---
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nsm37"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/chkBTCH2170-PRELIM").selected = True
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            print("Aguardando arquivos:")
            for p, o in zip(padroes, origens_por_padrao):
                print(f"   - {p} (Origem: {o})")
            print()

            encontrados = set()
            dest_counter = 0
            arquivo_counter = 1
            destinos_dict = {"destino": []}

            while True:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
                arquivos_encontrados_dict = {}

                for padrao, origem in zip(padroes, origens_por_padrao):
                    arquivos = glob.glob(os.path.join(origem, padrao))
                    if arquivos:
                        for arquivo in arquivos:
                            nome_arquivo = os.path.basename(arquivo)
                            destino_final = os.path.join(destino, nome_arquivo)
                            try:
                                shutil.move(arquivo, destino_final)
                                print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Arquivo encontrado e movido com sucesso:")
                                print(f"   De: {origem}")
                                print(f"DESTINO_FINAL_{dest_counter}: {destino_final}")
                                encontrados.add(padrao)

                                key_name = f"file_completa{arquivo_counter}"
                                arquivos_encontrados_dict[key_name] = destino_final
                                arquivo_counter += 1

                            except Exception as e:
                                import traceback
                                print(f"Erro ao mover {nome_arquivo}: {e}", flush=True)
                                print(traceback.format_exc(), flush=True)
                                status_done = "status_error"
                                os._exit(0)

                if arquivos_encontrados_dict:
                    destinos_dict["destino"].append(arquivos_encontrados_dict)
                    dest_counter += 1

                if len(encontrados) == len(padroes):
                    print("\nTodos os arquivos foram encontrados e movidos com sucesso.")
                    print("Encerrando monitoramento.")
                    print("Lista de arquivos movidos:", destinos_dict)
                    print("DESTINOS_DICT_JSON:", json.dumps(destinos_dict, ensure_ascii=False))
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
