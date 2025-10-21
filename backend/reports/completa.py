import os
import shutil
import glob
import time
from datetime import datetime

# Nome usuário Windows
username = os.getlogin()

# Caminhos de origem e destino
origem = r"C:\Users\U33V\PETROBRAS\GPP-E&P RXC GDI - Conteúdo Local\RGIT"
destino = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 0 - Arquivos de Texto do SAP"

# Padrão do arquivo a localizar
padrao = "RGT_RCL.CSV_U33V_JV3A5118530_D__20240101_2024_1T*.txt"

# Intervalo entre verificações (em segundos)
intervalo_busca = 120  # verifica a cada 60 segundos

print(f"🔍 Iniciando monitoramento da pasta:\n   {origem}")
print(f"Aguardando arquivo com padrão: {padrao}\n")

while True:
    arquivos = glob.glob(os.path.join(origem, padrao))

    if arquivos:
        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)
            destino_final = os.path.join(destino, nome_arquivo)
            try:
                shutil.move(arquivo, destino_final)
                print(f"\n✅ [{datetime.now().strftime('%H:%M:%S')}] Arquivo encontrado e movido com sucesso:")
                print(f"   ➜ {nome_arquivo}")
                print(f"   ➜ De: {origem}")
                print(f"   ➜ Para: {destino_final}")
                print("\nEncerrando monitoramento.")
                exit(0)
            except Exception as e:
                print(f"⚠️ Erro ao mover {nome_arquivo}: {e}")
                time.sleep(intervalo_busca)
    else:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Arquivo ainda não encontrado... tentando novamente em {intervalo_busca} segundos.")
        time.sleep(intervalo_busca)