from sap_manager import get_sap_free_session, start_sap_manager, start_connection, close_sap_manager
from ysanpcbc_job import process_YSANPCBC
from datetime import datetime, timedelta
import time

# ============================================================
# 🔹 Funções utilitárias
# ============================================================
def wait_for_element(session, element_id, timeout=15):
    """
    Espera até que o elemento esteja disponível no SAP.
    Retorna o elemento encontrado ou levanta TimeoutError.
    """
    for _ in range(timeout):
        try:
            elem = session.findById(element_id)
            if elem:
                return elem
        except Exception:
            time.sleep(1)
    raise TimeoutError(f"Elemento '{element_id}' não encontrado após {timeout}s.")


# ============================================================
# 🔹 Função principal para YSCLNRCL
# ============================================================
def create_YSCLBLRIT_requests(session, interval=None, requests_data=None):
    """
    Executa a transação YSCLNRCL e cria requisições de forma automatizada.
    """

    print("🔹 Iniciando transação YSCLNRCL...")
    wait_for_element(session, "wnd[0]/tbar[0]/okcd").text = "/nYSCLNRCL"
    wait_for_element(session, "wnd[0]").sendVKey(0)

    new_interval = 0

    for i, req in enumerate(requests_data, start=1):
        print(f"\n➡️  Processando requisição {i}...")

        # --- Preenche campos principais ---
        wait_for_element(session, "wnd[0]/usr/ctxtP_BUK_N").text = "1000"
        wait_for_element(session, "wnd[0]/usr/txtPC_ANO").text = "2011"
        wait_for_element(session, "wnd[0]/usr/cmbTRI").key = "1"
        wait_for_element(session, "wnd[0]/usr/ctxtPC_CODCB").text = ""
        wait_for_element(session, "wnd[0]/usr/ctxtPC_FASE").text = "D"
        wait_for_element(session, "wnd[0]/usr/ctxtPC_STAT").text = "3"
        wait_for_element(session, "wnd[0]/usr/ctxtP_VERSAO").text = "2"
        wait_for_element(session, "wnd[0]/usr/ctxtP_SECAO").text = "ANP_0901"

        # --- Abre filtro avançado ---
        wait_for_element(session, "wnd[0]/tbar[1]/btn[19]").press()
        wait_for_element(session, "wnd[0]/usr/ctxtSC_PSPID-LOW").text = req.get("SC_PSPID_LOW", "JV3A03108410")
        wait_for_element(session, "wnd[0]/usr/ctxtSD_DTINI-LOW").text = req.get("SD_DTINI_LOW", "01.01.2011")
        wait_for_element(session, "wnd[0]/usr/ctxtPC_BID").text = "002"
        wait_for_element(session, "wnd[0]/usr/chkP_PART").selected = True

        # --- Executa e abre tela de agendamento ---
        wait_for_element(session, "wnd[0]/mbar/menu[0]/menu[2]").select()
        wait_for_element(session, "wnd[1]/tbar[0]/btn[13]").press()
        wait_for_element(session, "wnd[1]/usr/btnDATE_PUSH").press()

        # --- Calcula data/hora do agendamento ---
        new_interval += int(interval) if interval else 1
        job_init_datetime = datetime.now() + timedelta(minutes=2 + new_interval)
        str_date_plan = job_init_datetime.strftime("%d.%m.%Y")
        str_time_plan = job_init_datetime.strftime("%H:%M:%S")
        print(f"🕒 Job agendado para {str_time_plan} de {str_date_plan}")

        # --- Preenche no SAP ---
        wait_for_element(session, "wnd[1]/usr/ctxtBTCH1010-SDLSTRTDT").text = str_date_plan
        wait_for_element(session, "wnd[1]/usr/ctxtBTCH1010-SDLSTRTTM").text = str_time_plan
        wait_for_element(session, "wnd[1]/tbar[0]/btn[0]").press()
        wait_for_element(session, "wnd[1]/tbar[0]/btn[11]").press()

        # --- Limpa filtros ---
        wait_for_element(session, "wnd[0]/tbar[1]/btn[19]").press()
        wait_for_element(session, "wnd[0]/usr/ctxtSC_PSPID-LOW").text = ""
        wait_for_element(session, "wnd[0]/usr/ctxtSD_DTINI-LOW").text = ""
        wait_for_element(session, "wnd[0]/tbar[1]/btn[20]").press()

        time.sleep(1)

    # Fecha a transação
    wait_for_element(session, "wnd[0]/tbar[0]/btn[15]").press()
    print("\n✅ Transação YSCLNRCL concluída com sucesso.")


# ============================================================
# 🔹 Execução principal
# ============================================================
if __name__ == "__main__":
    started_by_script = start_sap_manager()  # abre SAP se necessário
    start_connection()
    session = get_sap_free_session()

    # Exemplo de dados de entrada (pode vir de CSV futuramente)
    requests_data = [
        {"PC_CODCB": "ABC123", "SC_PSPID_LOW": "JV3A03108410", "SD_DTINI_LOW": "01.01.2011"},
        {"PC_CODCB": "XYZ789", "SC_PSPID_LOW": "JV3A03108420", "SD_DTINI_LOW": "01.01.2011"},
    ]

    # --- Executa YSCLNRCL ---
    create_YSCLBLRIT_requests(session=session, interval=15, requests_data=requests_data)

    # --- Executa YSANPCBC na mesma sessão ---
    codcb_list = [req["PC_CODCB"] for req in requests_data]
    process_YSANPCBC(session, codcb_list)

    # Opcional: fecha SAP se foi aberto pelo script
    # close_sap_manager(started_by_script)
