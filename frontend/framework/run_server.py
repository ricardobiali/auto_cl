from flask import Flask, send_from_directory, jsonify, request
import os
import json
import subprocess
import user_data
from flask_cors import CORS

# Caminho base do frontend (onde estão o index.html e outros arquivos)
BASE_DIR = os.path.dirname(__file__)

# Caminho completo do requests.json (ajuste conforme sua rede)
REQUESTS_PATH = os.path.join(BASE_DIR, "requests.json")

# ✅ Caminho completo do script ysclnrcl_job.py
YSCLNRCL_PATH = r"C:\Users\U33V\OneDrive - PETROBRAS\Desktop\python\auto_cl_prototype\backend\sap_manager\ysclnrcl_job.py"

app = Flask(__name__, static_folder=BASE_DIR, static_url_path='')
CORS(app)

# -----------------------------
# Rota principal (index.html)
# -----------------------------
@app.route('/')
@app.route('/index.html')
def index():
    index_path = os.path.join(BASE_DIR, "index.html")

    with open(index_path, "r", encoding="utf-8") as f:
        html_template = f.read()

    # Insere o nome do usuário dinamicamente
    html_content = html_template.replace(
        '<p class="mb-0" style="text-align: left; font-weight: bold;">Seja bem-vindo</p>',
        f'<p class="mb-0" style="text-align: left; font-weight: bold;">Seja bem-vindo, {user_data.full_name}!</p>'
    )

    return html_content


# -----------------------------
# Endpoint para salvar o JSON
# -----------------------------
@app.route('/save_requests', methods=['POST'])
def save_requests():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Nenhum dado recebido"}), 400

        # Salva o arquivo diretamente no local desejado
        with open(REQUESTS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        
        # ✅ Executa o script ysclnrcl_job.py logo após salvar o JSON
        subprocess.Popen(["python", YSCLNRCL_PATH], shell=True)

        return jsonify({
            "status": "ok",
            "message": f"Arquivo salvo em {REQUESTS_PATH}"
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# -----------------------------
# Servir arquivos estáticos (CSS, JS, etc.)
# -----------------------------
@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory(BASE_DIR, filename)


# -----------------------------
# Inicialização do servidor
# -----------------------------
if __name__ == '__main__':
    port = 8000
    print(f"🚀 Servidor rodando em http://localhost:{port}")
    print(f"📂 Salvando requests.json em: {REQUESTS_PATH}")
    print(f"⚙️ Executando job em: {YSCLNRCL_PATH}")
    app.run(host='127.0.0.1', port=port, debug=True)
