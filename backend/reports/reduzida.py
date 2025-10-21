import sys
import os

# Caminho absoluto até o diretório raiz do projeto
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from backend.sap_manager.sap_connect import get_sap_free_session, start_sap_manager, start_connection

import time
import pandas as pd
import win32com.client

# --- Caminhos ---
arquivo_origem = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 0 - Arquivos de Texto do SAP\RGT_RCL.CSV_U33V_JV3A5118530_D__20240101_2024_1T_20251019_194620.txt"
pasta_destino = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 2 - Arquivos de Excel Reduzidos"
os.makedirs(pasta_destino, exist_ok=True)

# --- Colunas desejadas ---
colunas_desejadas = [
    2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 16, 17, 18, 19, 24, 26, 27, 31, 
    33, 34, 35, 36, 37, 38, 39, 41, 42, 43, 44, 45, 46, 47, 49, 51, 52, 53, 
    13, 56, 58, 59, 60, 61, 62, 65, 66, 70, 72, 75, 76, 82, 88, 91, 92, 97, 100, 
    102, 103, 106, 114, 115, 116, 117, 118, 119, 120, 122, 123, 124, 128, 129, 
    134, 136, 137, 139, 144, 148, 157, 158, 159, 162, 163, 178, 180, 184, 188, 
    189, 192, 194, 195, 196, 213, 214, 215, 217, 218, 219, 220, 221, 222, 223
]
colunas_zero_base = [i - 1 for i in colunas_desejadas]

# --- Ler arquivo ---
try:
    df = pd.read_csv(arquivo_origem, sep=';', encoding='utf-8', low_memory=False)
except UnicodeDecodeError:
    df = pd.read_csv(arquivo_origem, sep=';', encoding='latin1', low_memory=False)

if df.shape[1] < max(colunas_zero_base) + 1:
    raise ValueError(f"O arquivo possui apenas {df.shape[1]} colunas — esperado no mínimo {max(colunas_zero_base) + 1}")

df_reduzido = df.iloc[:, colunas_zero_base]

# --- Remover linhas com "X" em Doc custo Expurgado ---
if "Doc custo Expurgado" in df_reduzido.columns:
    linhas_antes = len(df_reduzido)
    df_reduzido = df_reduzido[df_reduzido["Doc custo Expurgado"].astype(str).str.strip().str.upper() != "X"]
    print(f"🧹 {linhas_antes - len(df_reduzido)} linhas removidas (Doc custo Expurgado = 'X').")

# --- Converter colunas numéricas ---
colunas_numericas = ["Valor/Moeda obj", "Valor total em reais", "Val suj cont loc R$", "Valor cont local R$", "Valor/moeda ACC"]

for col in colunas_numericas:
    if col in df_reduzido.columns:
        df_reduzido[col] = pd.to_numeric(df_reduzido[col].astype(str).str.replace(",", "").str.strip(), errors='coerce').fillna(0)

# --- Criar coluna Estrangeiro $ ---
if "Valor cont local R$" in df_reduzido.columns and "Valor/moeda ACC" in df_reduzido.columns:
    idx = df_reduzido.columns.get_loc("Valor/moeda ACC") + 1
    df_reduzido.insert(idx, "Estrangeiro $", df_reduzido["Valor cont local R$"] - df_reduzido["Valor/moeda ACC"])

# --- Formatar números em padrão brasileiro ---
def formata_brasileiro(x):
    if pd.notnull(x):
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return x

for col in colunas_numericas + ["Estrangeiro $"]:
    if col in df_reduzido.columns:
        df_reduzido[col] = df_reduzido[col].apply(formata_brasileiro)

# --- Adicionar colunas vazias ---
novas_colunas = ["Tipo de Gasto", "Bem/Serviço", "Gestor do Contrato", "Gerência responsável pelo objeto parceiro", "Disciplina"]
for col in novas_colunas:
    df_reduzido[col] = ""

# --- Preencher Tipo de Gasto ---
if "Tipo de Gasto" in df_reduzido.columns:
    # Garante que as colunas necessárias existam
    for col in ["Protocolo", "Objeto parceiro", "Doc.material"]:
        if col not in df_reduzido.columns:
            df_reduzido[col] = None

    # Converte tipos
    protocolo_num = pd.to_numeric(df_reduzido["Protocolo"], errors='coerce').fillna(0)
    doc_material_num = pd.to_numeric(df_reduzido["Doc.material"], errors='coerce').fillna(0)

    # Limpa e normaliza Objeto parceiro
    objeto_parceiro = (
        df_reduzido["Objeto parceiro"]
        .astype(str)
        .str.strip()
        .fillna("")  # substitui NaN por string vazia
    )
    # remove casos em que virou "nan" como texto
    objeto_parceiro = objeto_parceiro.replace("nan", "")

    # Começa tudo como 'Outros'
    df_reduzido["Tipo de Gasto"] = "Outros"

    mask_direto = protocolo_num > 0
    df_reduzido.loc[mask_direto, "Tipo de Gasto"] = "Direto"

    mask_indireto = (df_reduzido["Tipo de Gasto"] == "Outros") & (objeto_parceiro != "")
    df_reduzido.loc[mask_indireto, "Tipo de Gasto"] = "Indireto"

    mask_estoque = (
        (df_reduzido["Tipo de Gasto"] == "Outros")
        & (doc_material_num > 4899999999)
        & (doc_material_num < 5000000000)
    )
    df_reduzido.loc[mask_estoque, "Tipo de Gasto"] = "Estoque"

    print("✅ Coluna 'Tipo de Gasto' preenchida conforme regras de prioridade (Direto → Indireto → Estoque → Outros).")

# --- Preencher Bem/Serviço ---
if "Material" in df_reduzido.columns and "Bem/Serviço" in df_reduzido.columns:
    material_str = df_reduzido["Material"].astype(str).str.strip()

    df_reduzido.loc[
        material_str.str.match(r"^(50|70|80)"), "Bem/Serviço"
    ] = "Serviço"

    df_reduzido.loc[
        material_str.str.match(r"^(10|11|12)"), "Bem/Serviço"
    ] = "Material"

    print("✅ Coluna 'Bem/Serviço' preenchida conforme prefixos de 'Material'.")
else:
    print("⚠️ Coluna 'Material' ou 'Bem/Serviço' não encontrada — nenhuma regra aplicada.")

# Gerar lista única de contratos válidos (sem alterar o DataFrame)
contratos_unicos = (
    df["Contrato"]
    .dropna()  # remove células vazias
    .astype(str)  # garante que tudo é string
    .str.strip()  # remove espaços em branco
)
contratos_unicos = [
    c for c in contratos_unicos.unique() if c and c != "*"  # remove "*" e strings vazias
]

# --- Inicialização SAP ---
print("🚀 Iniciando SAP GUI...")
start_sap_manager()
start_connection()
session = get_sap_free_session()
time.sleep(2)

# --- Executa transação SAP - Contratos/Gerentes ---
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

        # Verifica se há popup de “sem resultados”
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

print("🔍 Executando consulta YSRELCONT...")
gerentes_por_contrato = executar_ysrelcont(session, contratos_unicos)
print(f"✅ Consulta SAP concluída. {len(gerentes_por_contrato)} contratos encontrados.")

# --- Preenche coluna Contrato ---
df_reduzido['Contrato'] = df_reduzido['Contrato'].astype(str).str.strip()
df_reduzido['Gestor do Contrato'] = df_reduzido['Contrato'].map(gerentes_por_contrato).fillna('')

# Gerar lista única de objetos válidos (sem alterar o DataFrame)
objetos_unicos = (
    df["Objeto parceiro"]
    .dropna()  # remove células vazias
    .astype(str)  # garante que tudo é string
    .str.strip()  # remove espaços em branco
)
objetos_unicos = [
    c for c in objetos_unicos.unique() if c and c != "*"  # remove "*" e strings vazias
]

# --- Separa por tipo ---
objetos_e = [o for o in objetos_unicos if o.startswith("E")]
objetos_or = [o for o in objetos_unicos if o.startswith("OR")]

# --- KO03 - Encontrar centro de custo correspondente a cada OR ---
def executar_ko03(session, ordens_or):
    """Executa KO03 e retorna dict {ORxxxxx: Exxxxx}"""
    if not ordens_or:
        return {}

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKO03"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)
    # session.findById("wnd[0]").maximize()

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
                session.findById("wnd[0]/tbar[0]/btn[3]").press()  # Voltar
                time.sleep(0.5)
        except Exception as e:
            print(f"⚠️ Erro ao buscar {ordem}: {e}")
            continue

    return or_para_e


# --- KS13 - Buscar Gerência Responsável ---
def executar_ks13(session, objetos_e):
    """Executa KS13 e retorna dict {Exxxxx: GERÊNCIA}"""
    if not objetos_e:
        return {}

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nKS13"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    gerencias = {}

    # Seleciona área "ACPB"
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

    # Captura da tabela (varre todas as linhas)
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


# --- Execução KO03 + KS13 ---
print("🚀 Executando KO03 (ordens OR → centros E)...")
or_para_e = executar_ko03(session, objetos_or)
print(f"✅ {len(or_para_e)} ordens convertidas para centros de custo.")

# Monta lista definitiva de objetos E
objetos_definitivos = objetos_e + list(or_para_e.values())
objetos_definitivos = list(set(objetos_definitivos))  # remove duplicatas

print("🚀 Executando KS13 (centros E → gerências responsáveis)...")
gerencias_por_objeto = executar_ks13(session, objetos_definitivos)
print(f"✅ {len(gerencias_por_objeto)} gerências encontradas.")

# --- Preenche coluna no DataFrame ---
def mapear_gerencia(obj):
    """Mapeia OR via seu E correspondente ou direto"""
    if pd.isna(obj):
        return ""
    obj = str(obj).strip()
    if obj.startswith("E"):
        return gerencias_por_objeto.get(obj, "")
    if obj.startswith("OR"):
        centro = or_para_e.get(obj, "")
        return gerencias_por_objeto.get(centro, "") if centro else ""
    return ""

df_reduzido["Gerência responsável pelo objeto parceiro"] = df_reduzido["Objeto parceiro"].apply(mapear_gerencia)
print("✅ Coluna 'Gerência responsável pelo objeto parceiro' preenchida com sucesso.")
# --- Fim do uso do SAP ---

# --- Salvar arquivo final ---
nome_base = os.path.basename(arquivo_origem)
nome_reduzido = nome_base.replace(".txt", "_Reduzida.txt")
caminho_saida = os.path.join(pasta_destino, nome_reduzido)

df_reduzido.to_csv(caminho_saida, sep=';', index=False, encoding='utf-8')
print(f"✅ Arquivo reduzido criado com sucesso!\nDe: {arquivo_origem}\nPara: {caminho_saida}")
