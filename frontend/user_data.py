import os
import csv

# Nome do usuário Windows (em lowercase e sem espaços)
username = os.getlogin().strip().lower()

# Caminho do CSV relativo ao arquivo user_data.py
base_dir = os.path.dirname(__file__)
csv_path = os.path.join(base_dir, "user_data.csv")

full_name = None

# try:
with open(csv_path, newline='', encoding='utf-8-sig') as csvfile:
    # 'utf-8-sig' ignora o BOM se existir
    reader = csv.DictReader(csvfile, delimiter=';')
    for row in reader:
        # Limpa espaços e converte para lowercase todas as colunas
        row_clean = {k.strip().lower(): (v.strip().lower() if v else '') for k, v in row.items()}
        
        if 'chave' in row_clean and row_clean['chave'] == username:
            full_name = row.get('nome', '').strip()
            break

if not full_name:
        full_name = username

# except FileNotFoundError:
#     print(f"Arquivo CSV não encontrado: {csv_path}")
# except Exception as e:
#     print(f"Erro ao ler o CSV: {e}")

# if full_name:
#     print(f"Nome do usuário: {full_name}")
# else:
#     print(f"Usuário '{username}' não encontrado no CSV")
