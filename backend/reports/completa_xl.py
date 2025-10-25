import pandas as pd
from pathlib import Path

# Caminhos
arquivo_txt = Path(r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 0 - Arquivos de Texto do SAP\RGT_RCL.CSV_U33V_JV3A5118530_D_3_20240101_2024_1T_20251024_2047011000.txt")
pasta_excel = Path(r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 1 - Arquivos de Excel Completos")

# Garante que a pasta de destino existe
pasta_excel.mkdir(parents=True, exist_ok=True)

# Nome do arquivo Excel
nome_excel = arquivo_txt.stem + ".xlsx"  # mantém o mesmo nome do txt

arquivo_excel = pasta_excel / nome_excel

# Lê o CSV
df = pd.read_csv(arquivo_txt, sep=";", encoding="utf-8")  # ou sep="," dependendo do arquivo

# Salva em Excel
df.to_excel(arquivo_excel, index=False)

print(f"Arquivo salvo em: {arquivo_excel}")
