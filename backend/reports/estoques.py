import os
import pandas as pd

# --- Caminhos ---
arquivo_origem = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 2 - Arquivos de Excel Reduzidos\RGT_RCL.CSV_U33V_JV3A5118530_D__20240101_2024_1T_20251019_194620_Reduzida.txt"
pasta_destino = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 3 - Arquivos de Excel Finais\Estoques"
os.makedirs(pasta_destino, exist_ok=True)

# --- Lê o arquivo fonte ---
try:
    df = pd.read_csv(arquivo_origem, sep=';', encoding='utf-8', low_memory=False)
except UnicodeDecodeError:
    df = pd.read_csv(arquivo_origem, sep=';', encoding='latin1', low_memory=False)

# --- FILTRO DAS LINHAS COM 'Direto' ---
df_diretos = df[df["Tipo de Gasto"].str.strip().str.lower() == "estoque"]

# --- Salvar arquivo final ---
nome_base = os.path.basename(arquivo_origem)
nome_gastosDiretos = nome_base.replace("_Reduzida.txt", "_estoques.txt")
caminho_saida = os.path.join(pasta_destino, nome_gastosDiretos)

df_diretos.to_csv(caminho_saida, sep=';', index=False, encoding='utf-8')
print(f"✅ Arquivo reduzido criado com sucesso!\nDe: {arquivo_origem}\nPara: {caminho_saida}")