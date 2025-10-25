from pathlib import Path
from sap_connect import get_sap_free_session, start_sap_manager, start_connection, close_sap_manager
from datetime import datetime, timedelta
import subprocess
import os
import json

# Caminho atual do script
current_dir = Path(__file__).resolve()

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

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nYSCLNRCL"
    session.findById("wnd[0]").sendVKey(0)

    for i, req in enumerate(requests_data, start=1):
        print(f"Processando requisição {i}...")

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

# ============================================================
# Execução principal
# ============================================================
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

        # --- Executa relatório adicional completa.py ---
        reports_path = os.path.join(os.path.dirname(__file__), "..", "reports", "completa.py")
        if os.path.exists(reports_path):
            subprocess.run(["python", reports_path], check=True)
        else:
            print(f"Script de reports não encontrado: {reports_path}")

        # ✅ Se chegou aqui, tudo ocorreu sem erros
        status_done = "status_success"
        print("===STATUS_DONE===" + status_done)


    except Exception as e:
        print("Ocorreu um erro na execução do ysclnrcl_job.py:", e)
        status_done = "status_error"
        print("===STATUS_DONE===" + status_done)

