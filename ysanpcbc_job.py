import time

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


def process_YSANPCBC(session, codcb_list):
    """
    Executa a transação YSANPCBC para cada CODCB na lista.
    session: sessão SAP já aberta
    codcb_list: lista de códigos a serem processados
    """

    print("🔹 Iniciando transação YSANPCBC...")
    wait_for_element(session, "wnd[0]/tbar[0]/okcd").text = "YSANPCBC"
    wait_for_element(session, "wnd[0]").sendVKey(0)

    # Configuração inicial de colunas
    wait_for_element(session, "wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").pressToolbarButton("&COL0")
    wait_for_element(
        session,
        "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_CONFIGURATION:SAPLSALV_CUL_COLUMN_SELECTION:0620/"
        "cntlCONTAINER2_LAYO/shellcont/shell"
    ).selectColumn("SELTEXT")
    wait_for_element(
        session,
        "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_CONFIGURATION:SAPLSALV_CUL_COLUMN_SELECTION:0620/"
        "btnAPP_FL_SING"
    ).press()

    # Fecha janela de aviso se existir
    try:
        info_window = session.findById("wnd[2]")
        if info_window:
            info_window.findById("tbar[0]/btn[0]").press()
    except Exception:
        pass

    # Aplica a configuração
    wait_for_element(session, "wnd[1]/tbar[0]/btn[0]").press()

    # Ordena pela coluna DATUM decrescente
    wait_for_element(session, "wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectColumn("DATUM")
    wait_for_element(session, "wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").pressToolbarButton("&SORT_DSC")

    # Loop para cada CODCB
    for codcb in codcb_list:
        if codcb:
            wait_for_element(session, "wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectColumn("CODCB")
            wait_for_element(session, "wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").pressToolbarButton("&FIND")

            # Busca exata
            lpadded_value = str(codcb).zfill(8)
            wait_for_element(session, "wnd[1]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
            wait_for_element(session, "wnd[1]/usr/txtGS_SEARCH-VALUE").text = lpadded_value
            wait_for_element(session, "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
            wait_for_element(session, "wnd[1]/tbar[0]/btn[0]").press()

            # Verifica se encontrou resultados
            try:
                search_info = session.findById("wnd[1]/usr/txtGS_SEARCH-SEARCH_INFO").text
            except Exception:
                search_info = ""
            if "Nenhuma" not in search_info:
                wait_for_element(session, "wnd[1]/tbar[0]/btn[12]").press()

                # Exemplo: pegar valor da coluna DESANP
                shell = wait_for_element(session, "wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell")
                desanp_value = shell.GetCellValue(shell.currentCellRow, "DESANP")
                print(f"CODCB {lpadded_value} -> DESANP: {desanp_value}")
            else:
                wait_for_element(session, "wnd[1]/tbar[0]/btn[12]").press()

            time.sleep(0.5)

    # Fecha a transação
    wait_for_element(session, "wnd[0]/tbar[0]/btn[15]").press()
    print("\n✅ Transação YSANPCBC concluída com sucesso.")
