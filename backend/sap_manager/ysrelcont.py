import time

def executar_ysrelcont(session, contratos_unicos):
    """Executa YSRELCONT no SAP e retorna dict {contrato: gerente}"""
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nYSRELCONT"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    gerentes = {}
    contr_types = ["ZCVR", "ZCVM"]

    for contr_type in contr_types:
        session.findById("wnd[0]/usr/ctxtSC_BSART-LOW").text = contr_type
        session.findById("wnd[0]/usr/ctxtSC_EKORG-LOW").text = "0001"
        session.findById("wnd[0]/usr/ctxtSC_EKORG-HIGH").text = "9999"

        # Abre seleção de contratos
        session.findById("wnd[0]/usr/btn%_SC_EBELN_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").select()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()

        for contrato in contratos_unicos:
            campo = (
                "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                "ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
                "ctxtRSCSEL_255-SLOW_I[1,0]"
            )
            session.findById(campo).text = contrato
            session.findById("wnd[1]/tbar[0]/btn[13]").press()

        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(1)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(2)

        # Verifica popup “sem resultados”
        try:
            info_popup = session.findById("wnd[1]", False)
            if info_popup:
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                continue
        except:
            pass

        table = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        for r in range(table.rowCount):
            contrato = table.getCellValue(r, "EBELN").strip()
            gerente = table.getCellValue(r, "GERENTE").strip()
            if contrato not in gerentes:
                gerentes[contrato] = gerente

        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(1)

    return gerentes
