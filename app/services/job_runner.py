# app/services/job_runner.py
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from app.services.file_io import load_json, save_json_atomic, ensure_destino_list, set_status
from app.services.subprocess_runner import build_python_cmd, run_capture, spawn_stream
from app.state import JobState


class JobRunner:
    """
    Orquestra: SAP -> COMPLETA -> REDUZIDA

    - Atualiza STATE (mensagens, done)
    - Usa subprocess_runner (anti-deadlock)
    - Escreve requests.json com IO atômico
    """

    def __init__(
        self,
        state: JobState,
        requests_path: Path,
        sap_script: Path,
        completa_script: Path,
        reduzida_script: Path,
        creationflags: int,
        logger: Optional[logging.Logger] = None,
    ) -> None:
        self.state = state
        self.requests_path = requests_path
        self.sap_script = sap_script
        self.completa_script = completa_script
        self.reduzida_script = reduzida_script
        self.creationflags = creationflags
        self.log = logger or logging.getLogger(__name__)

    # --------------------
    # Helpers
    # --------------------
    def _cancel_point(self, msg: str = "Cancelado pelo usuário.") -> None:
        if self.state.cancel_requested():
            raise RuntimeError(msg)

    def _read_destino_first(self) -> Optional[Any]:
        data = load_json(self.requests_path)
        destinos = data.get("destino", [])
        if destinos:
            return destinos[0] if isinstance(destinos, list) else destinos
        return None

    def _write_file_completa1(self, value: Any) -> None:
        data = load_json(self.requests_path)
        ensure_destino_list(data)
        data["destino"][0]["file_completa1"] = value
        save_json_atomic(self.requests_path, data)

    def _status_update(self, script_name: str, status: str) -> None:
        data = load_json(self.requests_path)
        set_status(data, script_name, status)
        save_json_atomic(self.requests_path, data)

    def _normalize_files_iter(self, file_reduzida: Any) -> List[str]:
        if isinstance(file_reduzida, (str, Path)):
            return [str(file_reduzida)]
        if isinstance(file_reduzida, list):
            return [str(x) for x in file_reduzida]
        return [str(file_reduzida)]

    # ✅ decide se uma linha deve aparecer no "console do usuário"
    def _should_surface_sap_line(self, line: str) -> bool:
        s = (line or "").strip()
        if not s:
            return False
        keywords = (
            "Aguardando arquivos",
            "Ainda aguardando",
            "Arquivo encontrado",
            "Todos os arquivos foram encontrados",
            "Encerrando monitoramento",
            "DESTINO_FINAL_",
            "[",  # timestamps tipo [HH:MM:SS]
        )
        return any(k in s for k in keywords)

    # --------------------
    # Steps
    # --------------------
    def run_sap(self) -> Tuple[bool, Optional[dict], str]:
        """
        Roda SAP via spawn_stream e captura destinos_dict.
        Retorna: (ok, destinos_dict, stdout_total)
        """
        self._cancel_point()
        destinos_dict: Optional[dict] = None

        # UX: limpa logs visíveis e começa com uma linha inicial
        self.state.clear_logs()
        self.state.append_log("Iniciando SAP...")

        def on_line(line: str):
            nonlocal destinos_dict
            line = (line or "").rstrip("\r\n")

            # sempre loga pro arquivo
            self.log.info(line)

            # envia pro frontend (apenas linhas relevantes)
            if self._should_surface_sap_line(line):
                self.state.append_log(line)

            # captura destinos_dict como antes
            if line.startswith("DESTINOS_DICT_JSON:"):
                try:
                    destinos_dict = json.loads(line.replace("DESTINOS_DICT_JSON:", "").strip())
                except json.JSONDecodeError:
                    destinos_dict = None

        cmd = build_python_cmd(self.sap_script)

        rc, stdout_total = spawn_stream(
            cmd,
            on_line=on_line,
            creationflags=self.creationflags,
            cancel_check=self.state.cancel_requested,
            register_proc=self.state.register_proc,
        )

        ok = rc == 0
        status = "status_success" if "status_success" in stdout_total else "status_error"
        self._status_update("ysclnrcl_job.py", status)

        # se veio destinos_dict, grava no requests.json (mantém seu contrato atual)
        if destinos_dict:
            data = load_json(self.requests_path)
            data["destino"] = destinos_dict.get("destino", [])
            if not isinstance(data["destino"], list):
                data["destino"] = [data["destino"]]
            save_json_atomic(self.requests_path, data)

        # UX: linha final visível ao usuário
        self.state.append_log("SAP finalizado." if ok else "SAP finalizado com erro.")

        return ok, destinos_dict, stdout_total

    def run_completa(self) -> Tuple[bool, str]:
        self._cancel_point()
        cmd = build_python_cmd(self.completa_script)
        r = run_capture(cmd, creationflags=self.creationflags)

        if r.stdout:
            self.log.info(r.stdout)
        if r.stderr:
            self.log.error(r.stderr)

        status = "status_success" if "status_success" in r.stdout else "status_error"
        self._status_update("completa_xl.py", status)
        return (r.returncode == 0), r.stdout

    def run_reduzida(self) -> Tuple[bool, str]:
        self._cancel_point()
        cmd = build_python_cmd(self.reduzida_script)
        r = run_capture(cmd, creationflags=self.creationflags)

        if r.stdout:
            self.log.info(r.stdout)
        if r.stderr:
            self.log.error(r.stderr)

        status = "status_success" if "status_success" in r.stdout else "status_error"
        self._status_update("reduzida.py", status)
        return (r.returncode == 0), r.stdout

    # --------------------
    # Public orchestration
    # --------------------
    def run_sequence(
        self,
        switches: Dict[str, Any],
        paths: Dict[str, Any],
        selecionar_arquivo_cb: Callable[[], List[str]],
    ) -> None:
        """
        Executa a sequência completa com atualizações de STATE.
        selecionar_arquivo_cb: callback para abrir dialog e retornar lista de arquivos .txt
        """
        try:
            self.state.set_running("Executando automação...")

            destino_final = None
            file_completa = None
            file_reduzida = None

            # 1) SAP
            if switches.get("report_SAP"):
                self.state.set_message("Etapa 1/3: SAP...")
                ok, _destinos, _out = self.run_sap()

                if not ok and not switches.get("completa") and not switches.get("reduzida"):
                    self.state.set_done(False, "Falha no Job SAP.")
                    return

                destino_final = self._read_destino_first()

            self._cancel_point()

            # 2) COMPLETA
            if switches.get("completa"):
                self.state.set_message("Etapa 2/3: COMPLETA...")

                if destino_final:
                    file_completa = destino_final
                elif paths.get("file_completa"):
                    file_completa = paths.get("file_completa")
                else:
                    files = selecionar_arquivo_cb()
                    if not files:
                        self.state.set_done(False, "Execução cancelada: nenhum arquivo selecionado para Completa.")
                        return
                    file_completa = files
                    paths["file_completa"] = files

                # compat: escreve destino[0].file_completa1
                if isinstance(file_completa, dict):
                    file_completa_value = file_completa.get("file_completa1") or next(iter(file_completa.values()))
                else:
                    file_completa_value = file_completa

                self._write_file_completa1(file_completa_value)

                ok, _out = self.run_completa()
                if not ok:
                    self.state.set_done(False, "Falha no job COMPLETA.")
                    return

                destino_final = self._read_destino_first()

            self._cancel_point()

            # 3) REDUZIDA
            if switches.get("reduzida"):
                # ✅ NOVO: feedback inicial
                self.state.set_message("Etapa 3/3: REDUZIDA...")

                if switches.get("report_SAP") and destino_final:
                    file_reduzida = destino_final
                elif switches.get("completa") and file_completa:
                    file_reduzida = file_completa
                    paths["file_reduzida"] = file_completa
                else:
                    file_reduzida = paths.get("file_reduzida", [])
                    if not file_reduzida:
                        files = selecionar_arquivo_cb()
                        if not files:
                            self.state.set_done(False, "Execução cancelada: nenhum arquivo selecionado para Reduzida.")
                            return
                        file_reduzida = files
                        paths["file_reduzida"] = files

                files_iter = self._normalize_files_iter(file_reduzida)
                total = len(files_iter)

                for idx, file_txt in enumerate(files_iter, start=1):
                    self._cancel_point()

                    nome = Path(file_txt).name
                    # ✅ NOVO: progresso "executando 1/4"
                    self.state.set_message(f"Etapa 3/3: REDUZIDA — executando {idx}/{total} ({nome})")

                    self._write_file_completa1(file_txt)
                    ok, _out = self.run_reduzida()
                    if not ok:
                        self.state.set_done(False, f"Falha no job REDUZIDA ({idx}/{total}).")
                        return

            self.state.set_done(True, "Jobs concluídos em sequência.")

        except Exception as e:
            self.state.set_done(False, f"Erro: {e}")
