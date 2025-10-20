import os
import pandas as pd

# --- Caminhos ---
arquivo_origem = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 0 - Arquivos de Texto do SAP\RGT_RCL.CSV_U33V_JV3A5118530_D__20240101_2024_1T_20251019_194620.txt"
pasta_destino = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 2 - Arquivos de Excel Reduzidos"
os.makedirs(pasta_destino, exist_ok=True)

# --- Colunas desejadas (1-base) ---
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
    # Garantir colunas de referência
    for col in ["Protocolo", "Objeto parceiro", "Doc.material"]:
        if col not in df_reduzido.columns:
            df_reduzido[col] = None

    protocolo_num = pd.to_numeric(df_reduzido["Protocolo"], errors='coerce').fillna(0)
    doc_material_num = pd.to_numeric(df_reduzido["Doc.material"], errors='coerce').fillna(0)
    objeto_parceiro = df_reduzido["Objeto parceiro"].astype(str).str.strip()

    df_reduzido["Tipo de Gasto"] = "Outros"
    df_reduzido.loc[protocolo_num > 0, "Tipo de Gasto"] = "Direto"
    df_reduzido.loc[(df_reduzido["Tipo de Gasto"] == "Outros") & (objeto_parceiro == ""), "Tipo de Gasto"] = "Indireto"
    df_reduzido.loc[(df_reduzido["Tipo de Gasto"] == "Outros") & 
                    (doc_material_num > 4899999999) & (doc_material_num < 5000000000), "Tipo de Gasto"] = "Estoque"

    print("✅ Coluna 'Tipo de Gasto' preenchida com regras otimizadas.")

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

# --- Salvar arquivo final ---
nome_base = os.path.basename(arquivo_origem)
nome_reduzido = nome_base.replace(".txt", "_Reduzida.txt")
caminho_saida = os.path.join(pasta_destino, nome_reduzido)

df_reduzido.to_csv(caminho_saida, sep=';', index=False, encoding='utf-8')
print(f"✅ Arquivo reduzido criado com sucesso!\nDe: {arquivo_origem}\nPara: {caminho_saida}")
