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
path_txtOrigin = ""
path_txtOrigin = data["file_completa"]

path2_value = ""
if "paths" in data and len(data["paths"]) > 0:
    path2_value = data["paths"][0].get("path2", "")

# Caminhos baseados no path2
arquivo_txt = Path(path_txtOrigin)
pasta_excel = Path(path2_value)

# Garante que a pasta de destino existe
pasta_excel.mkdir(parents=True, exist_ok=True)

# Nome do arquivo Excel
nome_excel = arquivo_txt.stem + ".xlsx"
arquivo_excel = pasta_excel / nome_excel

try:
    # Lê o CSV e salva em Excel
    df = pd.read_csv(arquivo_txt, sep=";", encoding="utf-8")
    df.to_excel(arquivo_excel, index=False)

    # Marca status de sucesso
    status_done = "status_success"
    print(status_done)

except Exception as e:
    # Caso dê erro, marca status de erro
    status_done = "status_error"
    print(f"Ocorreu um erro: {e}")
    print(status_done)
