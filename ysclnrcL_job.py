from sap_manager import get_sap_free_session, start_sap_manager, start_connection, close_sap_manager
from datetime import datetime, timedelta
import time

def create_YSCLBLRIT_requests(session, init_date=None, init_time=None, interval=None, requests_data=None):
    """
    Executa a transação YSCLNRCL e cria requisições de forma automatizada.
    """

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nYSCLNRCL"
    session.findById("wnd[0]").sendVKey(0)

    new_interval = 0

    for i, req in enumerate(requests_data, start=1):
        print(f"🔹 Processando requisição {i}...")

        # --- Preenche campos principais ---
        session.findById("wnd[0]/usr/ctxtP_BUK_N").text = "1000"
        session.findById("wnd[0]/usr/txtPC_ANO").text = "2011"
        session.findById("wnd[0]/usr/cmbTRI").key = "1"
        session.findById("wnd[0]/usr/ctxtPC_CODCB").text = ""
        session.findById("wnd[0]/usr/ctxtPC_FASE").text = "D"
        session.findById("wnd[0]/usr/ctxtPC_STAT").text = "3"
        session.findById("wnd[0]/usr/ctxtP_VERSAO").text = "2"
        session.findById("wnd[0]/usr/ctxtP_SECAO").text = "ANP_0901"

        # --- Abre filtro avançado ---
        session.findById("wnd[0]/tbar[1]/btn[19]").press()
        session.findById("wnd[0]/usr/ctxtSC_PSPID-LOW").text = req.get("SC_PSPID_LOW", "JV3A03108410")
        session.findById("wnd[0]/usr/ctxtSD_DTINI-LOW").text = req.get("SD_DTINI_LOW", "01.01.2011")
        session.findById("wnd[0]/usr/ctxtPC_BID").text = "002"

        session.findById("wnd[0]/usr/chkP_PART").selected = True

        # --- Executa e abre tela de agendamento ---
        session.findById("wnd[0]/mbar/menu[0]/menu[2]").select()
        session.findById("wnd[1]/tbar[0]/btn[13]").press()
        session.findById("wnd[1]/usr/btnDATE_PUSH").press()

        # --- Calcula data/hora do agendamento ---
        # if init_date:
        #     job_init_datetime = datetime.strptime(init_date, "%d.%m.%Y")
        # else:
        #     job_init_datetime = datetime.now()

        # if init_time:
        #     t = datetime.strptime(init_time, "%H:%M").time()
        #     job_init_datetime = datetime.combine(job_init_datetime.date(), t)
        # else:
        #     job_init_datetime = datetime.now()

        # new_interval += int(interval) if interval else 1
        job_init_datetime = datetime.now() + timedelta(minutes=1)

        str_date_plan = job_init_datetime.strftime("%d.%m.%Y")
        str_time_plan = job_init_datetime.strftime("%H:%M:%S")

        # --- Preenche no SAP ---
        session.findById("wnd[1]/usr/ctxtBTCH1010-SDLSTRTDT").text = str_date_plan
        session.findById("wnd[1]/usr/ctxtBTCH1010-SDLSTRTTM").text = str_time_plan
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        print(f"✅ Requisição {i} agendada para {str_date_plan} às {str_time_plan}")

        # --- Limpa filtros ---
        session.findById("wnd[0]/tbar[1]/btn[19]").press()
        session.findById("wnd[0]/usr/ctxtSC_PSPID-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtSD_DTINI-LOW").text = ""
        session.findById("wnd[0]/tbar[1]/btn[20]").press()

    # Fecha a transação
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    print("🔚 Transação YSCLNRCL concluída com sucesso.")


# ============================================================
# Execução principal
# ============================================================
if __name__ == "__main__":
    started_by_script = start_sap_manager()  # abre SAP se necessário
    start_connection()
    session = get_sap_free_session()

    # Simula dados que seriam lidos de planilha ou banco
    requests_data = [
        {"PC_CODCB": "ABC123", "SC_PSPID_LOW": "JV3A03108410", "SD_DTINI_LOW": "01.01.2011"},
        {"PC_CODCB": "XYZ789", "SC_PSPID_LOW": "JV3A03108420", "SD_DTINI_LOW": "01.01.2011"},
    ]

    create_YSCLBLRIT_requests(
        session=session,
        init_date="01.01.2011",
        init_time="08:00",
        interval=15,
        requests_data=requests_data
    )

    # close_sap_manager(started_by_script)  # opcional, se quiser fechar ao fim
