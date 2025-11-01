import pandas as pd
from pathlib import Path
import json
import os

# Caminho atual do script
current_dir = Path(__file__).resolve()
username = os.getlogin()

# Sobe até encontrar a pasta 'auto_cl_prototype'
root_dir = current_dir
while root_dir.name != "auto_cl_prototype":
    if root_dir.parent == root_dir:
        raise FileNotFoundError("Pasta 'auto_cl_prototype' não encontrada.")
    root_dir = root_dir.parent

# Caminho do requests.json
requests_path = os.path.join(
    fr"{root_dir}\frontend",
    "requests.json"
)

# Lê o arquivo JSON
with open(requests_path, "r", encoding="utf-8") as f:
    data = json.load(f)

# Extrai o path2 do bloco "paths"
path2_value = ""
if "paths" in data and len(data["paths"]) > 0:
    path2_value = data["paths"][0].get("path2", "")

# Caminho de destino
pasta_excel = Path(path2_value)
pasta_excel.mkdir(parents=True, exist_ok=True)

# 🔹 Coleta todos os arquivos file_completaN do bloco "destino"
files_completa = []
destino_list = data.get("destino", [])
for destino_dict in destino_list:
    # Ordena as chaves file_completa1, 2, 3, ... para processar na sequência correta
    for key in sorted(destino_dict.keys(), key=lambda x: int(x.replace("file_completa", "")) if x != "file_completa" else 0):
        file_path = destino_dict[key]
        if file_path and os.path.exists(file_path):
            files_completa.append(file_path)
        else:
            print(f"Aviso: arquivo não encontrado -> {file_path}")

if not files_completa:
    print("Nenhum arquivo 'file_completa' válido encontrado no JSON.")
    status_done = "status_error"
else:
    # Processa cada arquivo em sequência
    for path_txtOrigin in files_completa:
        arquivo_txt = Path(path_txtOrigin)
        nome_excel = arquivo_txt.stem + ".xlsx"
        arquivo_excel = pasta_excel / nome_excel

        try:
            # Lê o CSV e salva em Excel
            df = pd.read_csv(arquivo_txt, sep=";", encoding="utf-8")
            df.to_excel(arquivo_excel, index=False)
            print(f"[OK] Convertido: {arquivo_txt.name} → {arquivo_excel.name}")

        except Exception as e:
            print(f"[ERRO] Falha ao converter {arquivo_txt}: {e}")

    status_done = "status_success"
    print(status_done)