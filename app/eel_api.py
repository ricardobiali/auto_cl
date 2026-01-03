# app/eel_api.py
from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

import psutil

from app.state import JobState
from app.services.file_io import load_json, save_json_atomic
from app.services.job_runner import JobRunner


def register_eel_api(
    eel: Any,
    *,
    state: JobState,
    runner: JobRunner,
    requests_path: Path,
    user_data_module: Any,
    selecionar_diretorio_cb: Callable[[], str],
    selecionar_arquivo_cb: Callable[[], list[str]],
) -> None:
    """
    Registra (eel.expose) todas as funções que o frontend (JS) chama.
    Mantém o main_app.py enxuto.
    """

    # -----------------------------
    # Dialogs (Tk) expostos ao JS
    # -----------------------------
    @eel.expose
    def selecionar_diretorio() -> str:
        return selecionar_diretorio_cb()

    @eel.expose
    def selecionar_arquivo() -> list[str]:
        return selecionar_arquivo_cb()

    # -----------------------------
    # Jobs (start/status/cancel)
    # -----------------------------
    @eel.expose
    def start_job(switches: dict, paths: dict) -> dict:
        # evita concorrência
        if state.is_running():
            return {"status": "already_running"}

        # validações mínimas (mantém compatibilidade do fluxo atual)
        if switches.get("completa") and not switches.get("report_SAP"):
            if not paths.get("file_completa"):
                files = selecionar_arquivo_cb()
                if not files:
                    return {"status": "cancelled"}
                paths["file_completa"] = files

        if switches.get("reduzida") and not switches.get("report_SAP") and not switches.get("completa"):
            if not paths.get("file_reduzida"):
                files = selecionar_arquivo_cb()
                if not files:
                    return {"status": "cancelled"}
                paths["file_reduzida"] = files

        # marca como rodando e dispara em thread
        state.set_running("Executando automação...")
        import threading

        threading.Thread(
            target=_run_job_thread,
            args=(switches, paths),
            daemon=True,
        ).start()

        return {"status": "started"}

    def _run_job_thread(switches: dict, paths: dict) -> None:
        runner.run_sequence(
            switches=switches,
            paths=paths,
            selecionar_arquivo_cb=selecionar_arquivo_cb,
        )

    @eel.expose
    def get_job_status() -> dict:
        return state.snapshot()

    @eel.expose
    def cancel_job() -> dict:
        """
        Cancelamento real:
        1) sinaliza cancelamento no state
        2) termina subprocessos filhos iniciados pelo app
        3) (opcional) mata SAPGUI/SAPLOGON (agressivo, mantém seu comportamento)
        """
        try:
            state.request_cancel()
            state.terminate_children()

            # mantém comportamento antigo: mata SAPGUI/SAPLOGON
            for proc in psutil.process_iter(["pid", "name"]):
                try:
                    name = (proc.info.get("name") or "").lower()
                    if name in ["saplogon.exe", "sapgui.exe"]:
                        proc.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            state.set_done(False, "Job cancelado manualmente.")
            return {"status": "cancelled"}

        except Exception:
            state.set_done(False, "Cancelamento solicitado (com avisos).")
            return {"status": "cancel_requested_with_warnings"}

    # -----------------------------
    # requests.json helpers (mantidos)
    # -----------------------------
    @eel.expose
    def save_requests(data: dict) -> dict:
        save_json_atomic(requests_path, data)
        return {"status": "saved"}

    @eel.expose
    def read_requests_json() -> dict:
        return load_json(requests_path)

    @eel.expose
    def write_requests_json(data: dict) -> dict:
        save_json_atomic(requests_path, data)
        return {"status": "saved"}

    # -----------------------------
    # Welcome (nome/gênero)
    # -----------------------------
    @eel.expose
    def get_welcome_user() -> dict:
        """
        Retorna { name, gender } para o frontend decidir 'bem-vindo/bem-vinda'.
        """
        name = getattr(user_data_module, "full_name", None) or ""
        gender = getattr(user_data_module, "gender", None) or ""
        return {"name": name, "gender": gender}

    # Backward-compatible (se seu frontend ainda chamar isso)
    @eel.expose
    def get_welcome_name() -> str:
        return getattr(user_data_module, "full_name", None) or ""
