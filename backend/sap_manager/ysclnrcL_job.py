from sap_connect import get_sap_free_session, start_sap_manager, start_connection, close_sap_manager
from datetime import datetime, timedelta
import subprocess
import os
import json

json_path = os.path.join(os.path.dirname(__file__), "requests.json")

def create_YSCLBLRIT_requests(session, init_date=None, init_time=None, interval=None, requests_data=None):
    """
    Executa a transação YSCLNRCL e cria requisições de forma automatizada.
    """

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nYSCLNRCL"
    session.findById("wnd[0]").sendVKey(0)

    for i, req in enumerate(requests_data, start=1):
        print(f"🔹 Processando requisição {i}...")

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

        print(f"✅ Requisição {i} agendada para {str_date_plan} às {str_time_plan}")

# ============================================================
# Execução principal
# ============================================================
if __name__ == "__main__":
    started_by_script = start_sap_manager()  # abre SAP se necessário
    start_connection()
    session = get_sap_free_session()

    # Simula dados que seriam lidos de planilha ou banco
    with open(json_path, "r", encoding="utf-8") as f:
        requests_data = json.load(f)

    create_YSCLBLRIT_requests(
        session=session,
        init_date="01.01.2011",
        init_time="08:00",
        interval=15,
        requests_data=requests_data
    )

# Caminho absoluto ou relativo
reports_path = os.path.join(os.path.dirname(__file__), "..", "reports", "completa.py")
subprocess.run(["python", reports_path], check=True)