import os
import shutil
import glob

# Nome usuário Windows
username = os.getlogin()

# Caminhos de origem e destino
origem = r"C:\Users\U33V\PETROBRAS\GPP-E&P RXC GDI - Conteúdo Local\RGIT"
destino = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\Auto_CL\Fase 0 - Arquivos de Texto do SAP"

# Padrão do arquivo a localizar
padrao = "RGT_RCL.CSV_U33V_JV3A5118530_D__20240101_2024_1T*.txt"

# Monta o caminho completo de busca
arquivos = glob.glob(os.path.join(origem, padrao))

if not arquivos:
    print("❌ Nenhum arquivo encontrado com o padrão especificado.")
else:
    for arquivo in arquivos:
        nome_arquivo = os.path.basename(arquivo)
        destino_final = os.path.join(destino, nome_arquivo)

        try:
            shutil.move(arquivo, destino_final)
            print(f"✅ Arquivo movido com sucesso: {nome_arquivo}")
            print(f"   ➜ De: {origem}")
            print(f"   ➜ Para: {destino_final}")
        except Exception as e:
            print(f"⚠️ Erro ao mover {nome_arquivo}: {e}")
