import time

def executar_ks13(session, objetos_e):
    """Executa KS13 e retorna dict {Exxxxx: GERÊNCIA}"""
    if not objetos_e:
        return {}

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKS13"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    gerencias = {}

    try:
        session.findById("wnd[0]").sendVKey(6)
        session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "ACPB"
        session.findById("wnd[1]").sendVKey(0)
    except:
        pass

    session.findById("wnd[0]/usr/subKOSTL_SELECTION:SAPLKMS1:0100/ctxtKMAS_D-KOSTL").setFocus()
    session.findById("wnd[0]").sendVKey(4)
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001").select()
    session.findById(
        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/"
        "ssubSUBSCR_PRESEL:SAPLSDH4:0220/"
        "sub:SAPLSDH4:0220/btnG_SELFLD_TAB-MORE[0,56]"
    ).press()

    # Preenche lista de objetos
    for obj in objetos_e:
        try:
            session.findById(
                "wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/"
                "ssubSCREEN_HEADER:SAPLALDB:3010/"
                "tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]"
            ).text = obj
            session.findById("wnd[2]/tbar[0]/btn[13]").press()
        except:
            pass

    session.findById("wnd[2]/tbar[0]/btn[8]").press()
    time.sleep(1)
    session.findById(
        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/"
        "ssubSUBSCR_PRESEL:SAPLSDH4:0220/chkG_SELPOP_STATE-BUTTON"
    ).selected = True
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
        container = session.findById("wnd[1]/usr")
        max_scroll = container.VerticalScrollbar.Maximum
        scroll = 0
        while True:
            container.VerticalScrollbar.Position = scroll
            elements = container.Children
            for i in range(19, len(elements), 10):
                cost_center = elements.ElementAt(i - 9).Text
                responsible = elements.ElementAt(i - 5).Text
                until_date = elements.ElementAt(i).Text
                if ".9999" in until_date:
                    gerencias[cost_center] = responsible
            if scroll >= max_scroll:
                break
            scroll += container.VerticalScrollbar.Range
            if scroll > max_scroll:
                scroll = max_scroll
        session.findById("wnd[1]/tbar[0]/btn[12]").press()
    except Exception as e:
        print(f"⚠️ Erro durante leitura KS13: {e}")

    return gerencias
