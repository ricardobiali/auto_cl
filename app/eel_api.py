# app/eel_api.py
from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from app.services.file_io import load_json, save_json_atomic
from app.services.importer import import_requests_from_excel, ImportErrorExcel
from app.state import JobState
from app.services.job_runner import JobRunner


def register_eel_api(
    eel,
    *,
    state: JobState,
    runner: JobRunner,
    requests_path: Path,
    user_data_module,
    selecionar_diretorio_cb,
    selecionar_arquivo_cb,
    selecionar_planilha_cb,
) -> None:
    """
    Registra todas as funções expostas ao JS.
    """

    # -----------------------------
    # Helpers internos
    # -----------------------------
    def _is_import_mode_enabled(data: Dict[str, Any]) -> bool:
        imp = data.get("imported") or {}
        return bool(isinstance(imp, dict) and imp.get("enabled") is True)

    def _set_paths_only(data: Dict[str, Any], paths: Dict[str, Any]) -> Dict[str, Any]:
        # mantém requests (principal)
        data["paths"] = [paths]
        return data

    # -----------------------------
    # API: welcome
    # -----------------------------
    @eel.expose
    def get_welcome_user():
        # ideal: seu user_data_module já retorna {name, gender}
        # se você tiver só full_name, adapte aqui
        try:
            return {
                "name": getattr(user_data_module, "full_name", "") or "",
                "gender": getattr(user_data_module, "gender", "m") or "m",
            }
        except Exception:
            return {"name": "", "gender": "m"}

    # -----------------------------
    # API: dialogs
    # -----------------------------
    @eel.expose
    def selecionar_diretorio():
        return selecionar_diretorio_cb()

    @eel.expose
    def selecionar_arquivo():
        return selecionar_arquivo_cb()

    # ✅ NOVO: selecionar planilha
    @eel.expose
    def selecionar_planilha():
        return selecionar_planilha_cb()

    # -----------------------------
    # API: status/cancel
    # -----------------------------
    @eel.expose
    def get_job_status():
        return state.snapshot()

    @eel.expose
    def cancel_job():
        state.request_cancel()
        try:
            state.terminate_children()
        except Exception:
            pass
        return {"ok": True}

    @eel.expose
    def import_planilha(switches: dict, paths: dict):
        if state.is_running():
            return {"status": "already_running"}

        if not switches.get("report_SAP"):
            return {"status": "invalid", "error": "Importação disponível apenas com Switch 1 (SAP) ativado."}

        file_path = selecionar_planilha_cb()
        if not file_path:
            return {"status": "cancelled"}

        try:
            result = import_requests_from_excel(file_path)
        except Exception as e:
            return {"status": "error", "error": str(e)}

        data = load_json(requests_path)
        data["requests"] = result.requests
        data["imported"] = {
            "enabled": True,
            "source_file": result.source_file,
            "imported_at": datetime.now().isoformat(timespec="seconds"),
            "rows_used": result.total_rows_used,
        }
        data["paths"] = [paths]
        data["switches"] = switches

        save_json_atomic(requests_path, data)

        # ✅ NÃO executa nada aqui
        return {"status": "imported", "imported_rows": result.total_rows_used}

    # -----------------------------
    # API: save/read/write requests.json
    # ✅ Blindagem: se imported.enabled=true, não deixa tabela sobrescrever requests
    # -----------------------------
    @eel.expose
    def save_requests(payload: Dict[str, Any]):
        existing = load_json(requests_path)

        # Se estiver em modo planilha, mantém requests importados
        if _is_import_mode_enabled(existing):
            # só atualiza paths/switches, ignora payload.requests
            paths_list = payload.get("paths") or []
            if isinstance(paths_list, list) and paths_list and isinstance(paths_list[0], dict):
                existing["paths"] = paths_list
            sw = payload.get("switches")
            if isinstance(sw, dict):
                existing["switches"] = sw

            save_json_atomic(requests_path, existing)
            return {"status": "saved", "note": "import_mode_active_requests_preserved"}

        # modo normal: salva tudo
        save_json_atomic(requests_path, payload)
        return {"status": "saved"}

    @eel.expose
    def read_requests_json():
        return load_json(requests_path)

    @eel.expose
    def write_requests_json(data: Dict[str, Any]):
        save_json_atomic(requests_path, data)
        return {"status": "saved"}

    # -----------------------------
    # API: start_job (modo normal, tabela)
    # ✅ Se existir import_mode ativo, ignora tabela e roda o que já está no JSON
    # -----------------------------
    @eel.expose
    def start_job(switches: Dict[str, Any], paths: Dict[str, Any]):
        if state.is_running():
            return {"status": "already_running"}

        existing = load_json(requests_path)
        if _is_import_mode_enabled(existing):
            # modo planilha ativo: apenas atualiza paths e roda
            existing = _set_paths_only(existing, paths)
            existing["switches"] = switches
            save_json_atomic(requests_path, existing)

            state.set_running("Executando automação (modo planilha)...")
            threading.Thread(
                target=runner.run_sequence,
                kwargs={
                    "switches": switches,
                    "paths": paths,
                    "selecionar_arquivo_cb": selecionar_arquivo_cb,
                },
                daemon=True,
            ).start()
            return {"status": "started", "mode": "import"}

        # modo normal: segue fluxo atual
        state.set_running("Executando automação...")
        threading.Thread(
            target=runner.run_sequence,
            kwargs={
                "switches": switches,
                "paths": paths,
                "selecionar_arquivo_cb": selecionar_arquivo_cb,
            },
            daemon=True,
        ).start()
        return {"status": "started", "mode": "table"}
