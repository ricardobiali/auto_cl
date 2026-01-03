# main_app.py
from __future__ import annotations

import os
import subprocess
import eel

from app.hook_logging import setup_logging
from app.paths import Paths
from app.state import JobState
from app.services import user_data
from app.services.job_runner import JobRunner

from app.eel_api import register_eel_api
from app.services.dialogs import selecionar_diretorio, selecionar_arquivo


def main() -> None:
    # logging primeiro
    setup_logging()

    # paths
    P = Paths.build()
    base_dir = P.frontend
    root_dir = P.root
    requests_path = P.requests_json

    # scripts
    sap_script = root_dir / "backend/sap_manager/ysclnrcl_job.py"
    completa_script = root_dir / "backend/reports/completa_xl.py"
    reduzida_script = root_dir / "backend/reports/reduzida.py"

    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

    # estado + runner
    state = JobState()
    runner = JobRunner(
        state=state,
        requests_path=requests_path,
        sap_script=sap_script,
        completa_script=completa_script,
        reduzida_script=reduzida_script,
        creationflags=creationflags,
    )

    # eel
    eel.init(str(base_dir))

    # registra API exposta ao JS
    register_eel_api(
        eel,
        state=state,
        runner=runner,
        requests_path=requests_path,
        user_data_module=user_data,
        selecionar_diretorio_cb=selecionar_diretorio,
        selecionar_arquivo_cb=selecionar_arquivo,
    )

    # start
    try:
        eel.start(
            "index.html",
            port=0,
            size=(1200, 800),
            cmdline_args=["--start-maximized"],
        )
    except OSError:
        eel.start(
            "index.html",
            port=8080,
            size=(1200, 800),
            cmdline_args=["--start-maximized"],
        )

if __name__ == "__main__":
    main()
