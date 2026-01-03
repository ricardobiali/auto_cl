# app/services/user_data.py
import os
import csv

username = os.getlogin().strip().lower()

base_dir = os.path.dirname(__file__)
csv_path = os.path.join(base_dir, "user_data.csv")

full_name = None
gender = None  # "m" ou "f"

try:
    with open(csv_path, newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile, delimiter=";")
        for row in reader:
            # normaliza chaves e valores para comparação
            row_clean = {k.strip().lower(): (v.strip().lower() if v else "") for k, v in row.items()}

            if row_clean.get("chave") == username:
                # pega o nome "original" (sem lowercase) se existir
                full_name = (row.get("nome") or username).strip()

                g = (row.get("gender") or row.get("genero") or "").strip().lower()
                gender = g if g in ("m", "f") else None
                break

except FileNotFoundError:
    print(f"Aviso: CSV de usuários não encontrado em {csv_path}")
except Exception as e:
    print(f"Aviso: Erro ao ler CSV de usuários: {e}")

if not full_name:
    full_name = username

# default seguro (se não tiver gender no CSV)
if gender is None:
    gender = "m"
