import os
import pandas as pd

# Caminho do arquivo original
arquivo_origem = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 0 - Arquivos de Texto do SAP\RGT_RCL.CSV_U33V_JV3A5118530_D__20240101_2024_1T_20251019_194620.txt"

# Caminho da pasta de destino
pasta_destino = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 2 - Arquivos de Excel Reduzidos"

# Criar a pasta de destino se não existir
os.makedirs(pasta_destino, exist_ok=True)

# Lista de colunas desejadas (índices 1-baseados conforme descrição)
colunas_desejadas = [
    2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 16, 17, 18, 19, 24, 26, 27, 31, 
    33, 34, 35, 36, 37, 38, 39, 41, 42, 43, 44, 45, 46, 47, 49, 51, 52, 53, 
    13, 56, 58, 59, 60, 61, 62, 65, 66, 70, 72, 75, 76, 82, 88, 91, 92, 97, 100, 
    102, 103, 106, 114, 115, 116, 117, 118, 119, 120, 122, 123, 124, 128, 129, 
    134, 136, 137, 139, 144, 148, 157, 158, 159, 162, 163, 178, 180, 184, 188, 
    189, 192, 194, 195, 196, 213, 214, 215, 217, 218, 219, 220, 221, 222, 223
]

# Converter para índice 0-baseado (pandas usa índice a partir de 0)
colunas_zero_base = [i - 1 for i in colunas_desejadas]

# Ler o arquivo TXT (assumindo separador ponto e vírgula)
try:
    df = pd.read_csv(arquivo_origem, sep=';', encoding='utf-8', low_memory=False)
except UnicodeDecodeError:
    df = pd.read_csv(arquivo_origem, sep=';', encoding='latin1', low_memory=False)

# Validar se o número de colunas é suficiente
if df.shape[1] < max(colunas_zero_base) + 1:
    raise ValueError(f"O arquivo possui apenas {df.shape[1]} colunas — esperado no mínimo {max(colunas_zero_base) + 1}")

# Selecionar apenas as colunas desejadas
df_reduzido = df.iloc[:, colunas_zero_base]

# --- 🧹 Remover linhas com "X" na coluna "Doc custo Expurgado" ---
if "Doc custo Expurgado" in df_reduzido.columns:
    linhas_antes = len(df_reduzido)
    df_reduzido = df_reduzido[df_reduzido["Doc custo Expurgado"].astype(str).str.strip().str.upper() != "X"]
    linhas_removidas = linhas_antes - len(df_reduzido)
    print(f"🧹 {linhas_removidas} linhas removidas (Doc custo Expurgado = 'X').")
else:
    print("⚠️ Atenção: coluna 'Doc custo Expurgado' não encontrada — nenhuma linha foi removida.")

# --- 💰 Converter colunas numéricas diretamente para float ---
colunas_numericas = [
    "Valor/Moeda obj", 
    "Valor total em reais", 
    "Val suj cont loc R$", 
    "Valor cont local R$", 
    "Valor/moeda ACC"
]

def to_float(valor):
    try:
        # remover milhar e ajustar separador decimal, se necessário
        valor = str(valor).replace(",", "")
        return float(valor)
    except:
        return 0.0

for col in colunas_numericas:
    if col in df_reduzido.columns:
        df_reduzido[col] = df_reduzido[col].apply(to_float)

# --- ➕ Criar coluna "Estrangeiro $" ---
if "Valor cont local R$" in df_reduzido.columns and "Valor/moeda ACC" in df_reduzido.columns:
    df_reduzido["Estrangeiro $"] = df_reduzido["Valor cont local R$"] - df_reduzido["Valor/moeda ACC"]

    # Inserir logo após a coluna "Valor/moeda ACC"
    idx = df_reduzido.columns.get_loc("Valor/moeda ACC") + 1
    cols = list(df_reduzido.columns)
    cols.insert(idx, cols.pop(cols.index("Estrangeiro $")))
    df_reduzido = df_reduzido[cols]
else:
    print("⚠️ Não foi possível criar a coluna 'Estrangeiro $' — colunas base não encontradas.")

# --- 🇧🇷 Formatar números em padrão brasileiro no final ---
def formata_brasileiro(x):
    if isinstance(x, (int, float)):
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return x

for col in colunas_numericas + ["Estrangeiro $"]:
    if col in df_reduzido.columns:
        df_reduzido[col] = df_reduzido[col].apply(formata_brasileiro)

# --- 🆕 Adicionar colunas vazias ao final ---
novas_colunas = [
    "Tipo de Gasto",
    "Bem/Serviço",
    "Gestor do Contrato",
    "Gerência responsável pelo objeto parceiro",
    "Disciplina"
]

for col in novas_colunas:
    df_reduzido[col] = ""

# Gerar nome do novo arquivo reduzido
nome_base = os.path.basename(arquivo_origem)
nome_reduzido = nome_base.replace(".txt", "_Reduzida.txt")
caminho_saida = os.path.join(pasta_destino, nome_reduzido)

# Salvar o novo arquivo
df_reduzido.to_csv(caminho_saida, sep=';', index=False, encoding='utf-8')

print(f"✅ Arquivo reduzido criado com sucesso!\nDe: {arquivo_origem}\nPara: {caminho_saida}")