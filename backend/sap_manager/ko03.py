import time

def executar_ko03(session, ordens_or):
    """Executa KO03 e retorna dict {ORxxxxx: Exxxxx}"""
    if not ordens_or:
        return {}

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKO03"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    or_para_e = {}

    for ordem in ordens_or:
        try:
            ordem_num = ordem.replace("OR", "")
            session.findById("wnd[0]/usr/ctxtCOAS-AUFNR").text = ordem_num
            session.findById("wnd[0]/tbar[1]/btn[42]").press()  # Executar
            time.sleep(0.5)

            status_message = session.findById("wnd[0]/sbar").text.strip()
            if not status_message:
                centro = session.findById(
                    "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT1/"
                    "ssubAREA_FOR_601:SAPMKAUF:0601/"
                    "subAREA1:SAPMKAUF:0315/ctxtCOAS-KOSTV"
                ).text.strip()
                if centro:
                    or_para_e[ordem] = centro
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                time.sleep(0.5)
        except Exception as e:
            print(f"⚠️ Erro ao buscar {ordem}: {e}")
            continue

    return or_para_e
