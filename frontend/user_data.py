import os
import csv

# Nome do usuário Windows (em lowercase e sem espaços)
username = os.getlogin().strip().lower()

# Caminho do CSV relativo ao arquivo user_data.py
base_dir = os.path.dirname(__file__)
csv_path = os.path.join(base_dir, "user_data.csv")

full_name = None

try:
    with open(csv_path, newline='', encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile, delimiter=';')
        for row in reader:
            # Limpa espaços e converte para lowercase
            row_clean = {k.strip().lower(): (v.strip().lower() if v else '') for k, v in row.items()}
            if 'chave' in row_clean and row_clean['chave'] == username:
                full_name = row.get('nome', username).strip()
                break
except FileNotFoundError:
    print(f"Aviso: CSV de usuários não encontrado em {csv_path}")
except Exception as e:
    print(f"Aviso: Erro ao ler CSV de usuários: {e}")

if not full_name:
        full_name = username