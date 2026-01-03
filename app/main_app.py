# app/main_app.py
from __future__ import annotations

import os
import sys
import subprocess
import runpy
import traceback
import eel

from app.hook_logging import setup_logging
from app.paths import Paths
from app.state import JobState
from app.services import user_data
from app.services.job_runner import JobRunner

from app.eel_api import register_eel_api
from app.services.dialogs import selecionar_diretorio, selecionar_arquivo


# ----------------------------
# ✅ Runner mode (PyInstaller)
# ----------------------------
def _dispatch_run_mode() -> None:
    """
    Quando chamado como:
        AUTO_CL.exe --run <script_name> [args...]
    executa o módulo correspondente e sai.
    """
    try:
        idx = sys.argv.index("--run")
    except ValueError:
        return  # normal mode

    # pega script_name e args extras
    if idx + 1 >= len(sys.argv):
        print("[ERRO] --run sem nome do script.", flush=True)
        raise SystemExit(2)

    script_name = sys.argv[idx + 1].strip()
    extra_args = sys.argv[idx + 2 :]  # se quiser repassar argumentos

    module_map = {
        "ysclnrcl_job": "backend.sap_manager.ysclnrcl_job",
        "completa_xl": "backend.reports.completa_xl",
        "reduzida": "backend.reports.reduzida",
    }

    mod = module_map.get(script_name)
    if not mod:
        print(f"[ERRO] Script desconhecido em --run: {script_name}", flush=True)
        print(f"Disponíveis: {', '.join(module_map.keys())}", flush=True)
        raise SystemExit(2)

    # ✅ unbuffer/encoding o mais cedo possível (antes de qualquer coisa)
    os.environ["PYTHONUNBUFFERED"] = "1"
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

    # ✅ força line buffering / write-through
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True, write_through=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True, write_through=True)
    except Exception:
        pass

    # ✅ logging também no worker (vai para o mesmo arquivo do app)
    # (se o seu Paths.build() aponta para a raiz do app, o log fica na pasta do exe)
    setup_logging()

    # ✅ muito importante: ajustar sys.argv para simular "python -m módulo args..."
    sys.argv = [mod, *extra_args]

    try:
        runpy.run_module(mod, run_name="__main__")
        raise SystemExit(0)
    except SystemExit:
        # respeita sys.exit() do script
        raise
    except Exception as e:
        # imprime traceback no stdout (vai para o streaming) e no log
        print("[ERRO] Falha no worker --run:", flush=True)
        traceback.print_exc()
        raise SystemExit(1)


def main() -> None:
    _dispatch_run_mode()

    setup_logging()

    P = Paths.build()
    base_dir = P.frontend
    root_dir = P.root
    requests_path = P.requests_json

    sap_script = root_dir / "backend/sap_manager/ysclnrcl_job.py"
    completa_script = root_dir / "backend/reports/completa_xl.py"
    reduzida_script = root_dir / "backend/reports/reduzida.py"

    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

    state = JobState()
    runner = JobRunner(
        state=state,
        requests_path=requests_path,
        sap_script=sap_script,
        completa_script=completa_script,
        reduzida_script=reduzida_script,
        creationflags=creationflags,
    )

    eel.init(str(base_dir))

    register_eel_api(
        eel,
        state=state,
        runner=runner,
        requests_path=requests_path,
        user_data_module=user_data,
        selecionar_diretorio_cb=selecionar_diretorio,
        selecionar_arquivo_cb=selecionar_arquivo,
    )

    try:
        eel.start("index.html", port=0, size=(1200, 800), cmdline_args=["--start-maximized"])
    except OSError:
        eel.start("index.html", port=8080, size=(1200, 800), cmdline_args=["--start-maximized"])


if __name__ == "__main__":
    main()
