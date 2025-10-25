from pathlib import Path
from flask import Flask, send_from_directory, jsonify, request
import os
import json
from flask_cors import CORS
from threading import Thread
import subprocess
import user_data

# Configurações de caminhos
current_dir = Path(__file__).resolve()

# Sobe até encontrar a pasta 'auto_cl_prototype'
root_dir = current_dir
while root_dir.name != "auto_cl_prototype":
    if root_dir.parent == root_dir:
        raise FileNotFoundError("Pasta 'auto_cl_prototype' não encontrada.")
    root_dir = root_dir.parent

BASE_DIR = os.path.dirname(__file__)
REQUESTS_PATH = os.path.join(BASE_DIR, "requests.json")
YSCLNRCL_PATH = fr"{root_dir}\backend\sap_manager\ysclnrcl_job.py"
COMPLETA_XL_PATH = fr"{root_dir}\backend\reports\completa_xl.py"


# Flask app
app = Flask(__name__, static_folder=BASE_DIR, static_url_path='')
CORS(app)

# Status do job (global simples)
job_status = {
    "running": False,
    "success": None,
    "message": ""
}

# -----------------------------
# Endpoint principal (index.html)
# -----------------------------
@app.route('/')
@app.route('/index.html')
def index():
    index_path = os.path.join(BASE_DIR, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        html_template = f.read()

    html_content = html_template.replace(
        '<p class="mb-0" style="text-align: left; font-weight: bold;">Seja bem-vindo</p>',
        f'<p class="mb-0" style="text-align: left; font-weight: bold;">Seja bem-vindo, {user_data.full_name}!</p>'
    )

    return html_content

# -----------------------------
# Função que executa o job
# -----------------------------
def run_job():
    global job_status
    job_status["running"] = True
    job_status["success"] = None
    job_status["message"] = "Executando automação..."

    try:
        completed_process = subprocess.run(
            ["python", YSCLNRCL_PATH],
            capture_output=True,
            text=True,
            check=True
        )
        print(completed_process.stdout)
        job_status["success"] = True
        job_status["message"] = "Arquivo criado com sucesso!"

    except subprocess.CalledProcessError as e:
        print(e.stderr)
        job_status["success"] = False
        job_status["message"] = f"Ocorreu um erro na execução do job:\n{e.stderr}"

    finally:
        job_status["running"] = False

# -----------------------------
# Endpoint para iniciar a automação
# -----------------------------
@app.route('/save_requests', methods=['POST'])
def save_requests():
    data = request.json
    if not data:
        return jsonify({"status": "error", "message": "Nenhum dado recebido"}), 400

    final_data = {
        "paths": data.get("paths", []),
        "requests": data.get("requests", [])
    }

    # Salva o JSON de entrada
    with open(REQUESTS_PATH, "w", encoding="utf-8") as f:
        json.dump(final_data, f, indent=4, ensure_ascii=False)

    # --- Identifica switches ativos ---
    switches = data.get("switches", {})
    switch1_active = switches.get("report_SAP", False)
    switch2_active = switches.get("completa", False)

    results = []

    # --- Executa scripts conforme switches ---
    try:
        if switch1_active:
            completed_process_1 = subprocess.run(
                ["python", YSCLNRCL_PATH],
                capture_output=True,
                text=True,
                check=True
            )
            print(completed_process_1.stdout)
            results.append("Job SAP executado com sucesso.")

        if switch2_active:
            completed_process_2 = subprocess.run(
                ["python", COMPLETA_XL_PATH],
                capture_output=True,
                text=True,
                check=True
            )
            print(completed_process_2.stdout)
            results.append("Relatório completo gerado com sucesso.")

        if not (switch1_active or switch2_active):
            return jsonify({"status": "error", "message": "Nenhum modo de execução selecionado."}), 400

        return jsonify({
            "status": "success",
            "message": " | ".join(results)
        })

    except subprocess.CalledProcessError as e:
        print(e.stderr)
        return jsonify({
            "status": "error",
            "message": f"Ocorreu um erro na execução:\n{e.stderr}"
        }), 500

# -----------------------------
# Endpoint para consultar status do job
# -----------------------------
@app.route('/job_status', methods=['GET'])
def check_status():
    return jsonify(job_status)

# -----------------------------
# Servir arquivos estáticos
# -----------------------------
@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory(BASE_DIR, filename)

# -----------------------------
# Inicialização
# -----------------------------
if __name__ == '__main__':
    port = 8000
    print(f"Servidor rodando em http://localhost:{port}")
    print(f"Salvando requests.json em: {REQUESTS_PATH}")
    print(f"Executando job em: {YSCLNRCL_PATH}")
    app.run(host='127.0.0.1', port=port, debug=True)
