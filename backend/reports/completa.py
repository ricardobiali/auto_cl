import pandas as pd
from pathlib import Path
import json
import os
import sys

APP_NAME = "AUTO_CL"

def _requests_path_appdata() -> Path:
    """
    requests.json persistente em AppData:
    %LOCALAPPDATA%\AUTO_CL\requests.json
    """
    base = os.environ.get("LOCALAPPDATA")
    if base:
        appdata_dir = Path(base) / APP_NAME
    else:
        appdata_dir = Path.home() / f".{APP_NAME.lower()}"
    appdata_dir.mkdir(parents=True, exist_ok=True)
    return appdata_dir / "requests.json"


# Caminho do requests.json (centralizado em AppData)
requests_path = _requests_path_appdata()

if not requests_path.exists():
    print(f"[ERRO] Arquivo requests.json não encontrado em: {requests_path}")
    sys.exit(1)

# Lê o arquivo JSON
try:
    with open(requests_path, "r", encoding="utf-8") as f:
        data = json.load(f)
except Exception as e:
    print(f"[ERRO] Falha ao ler requests.json ({requests_path}): {e}")
    sys.exit(1)

# Extrai o path2 do bloco "paths"
path2_value = ""
if "paths" in data and isinstance(data["paths"], list) and len(data["paths"]) > 0:
    path2_value = (data["paths"][0].get("path2", "") or "").strip()

# Caminho de destino
if not path2_value:
    print("[ERRO] 'path2' vazio no requests.json (paths[0].path2).")
    sys.exit(1)

pasta_excel = Path(path2_value)
pasta_excel.mkdir(parents=True, exist_ok=True)

# 🔹 Coleta todos os arquivos file_completaN do bloco "destino"
files_completa = []
destino_list = data.get("destino", [])

if not isinstance(destino_list, list):
    destino_list = [destino_list]

for destino_dict in destino_list:
    if not isinstance(destino_dict, dict):
        continue

    # Ordena as chaves file_completa1, 2, 3, ...
    def _key_sort(x: str) -> int:
        try:
            return int(x.replace("file_completa", ""))
        except Exception:
            return 0

    for key in sorted(destino_dict.keys(), key=_key_sort):
        file_path = destino_dict.get(key)
        if file_path and os.path.exists(file_path):
            files_completa.append(file_path)
        else:
            print(f"Aviso: arquivo não encontrado - {file_path}")

if not files_completa:
    print("[ERRO] Nenhum arquivo 'file_completaN' válido encontrado no requests.json.")
    print("status_error")
    sys.exit(1)

# Processa cada arquivo em sequência
for path_txtOrigin in files_completa:
    arquivo_txt = Path(path_txtOrigin)
    nome_excel = arquivo_txt.stem + ".xlsx"
    arquivo_excel = pasta_excel / nome_excel

    try:
        df = pd.read_csv(arquivo_txt, sep=";", encoding="utf-8")
        df.to_excel(arquivo_excel, index=False)
        print(f"[OK] Convertido: {arquivo_txt.name} -> {arquivo_excel.name}")
    except Exception as e:
        print(f"[ERRO] Falha ao converter {arquivo_txt}: {e}")

print("status_success")
