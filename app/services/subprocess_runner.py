# app/services/subprocess_runner.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional, List, Union, Tuple

import os
import subprocess
import threading
import queue
import sys
import time


LineCallback = Callable[[str], None]


@dataclass
class Completed:
    stdout: str
    stderr: str
    returncode: int


def _get_python_exec() -> str:
    """
    Resolve qual python executar em DEV vs PyInstaller.
    """
    if getattr(sys, "frozen", False):
        python_exec = str(Path(sys._MEIPASS).parent / "python.exe")  # type: ignore[attr-defined]
        if not os.path.exists(python_exec):
            return "python"
        return python_exec
    return sys.executable


def build_python_cmd(script_path: Union[str, Path], args: Optional[List[str]] = None) -> List[str]:
    """
    Monta comando [python, -u, script, ...args]
    """
    py = _get_python_exec()
    cmd = [py, "-u", str(script_path)]
    if args:
        cmd.extend(args)
    return cmd


def run_capture(
    cmd: List[str],
    creationflags: int = 0,
    timeout: Optional[float] = None,
) -> Completed:
    """
    Executa e captura stdout/stderr em memória (bom para jobs curtos).
    """
    r = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        creationflags=creationflags,
        timeout=timeout,
    )
    return Completed(stdout=r.stdout or "", stderr=r.stderr or "", returncode=r.returncode)


def spawn_stream(
    cmd: List[str],
    on_line: Optional[LineCallback] = None,
    creationflags: int = 0,
    cancel_check: Optional[Callable[[], bool]] = None,
    register_proc: Optional[Callable[[subprocess.Popen], None]] = None,
    poll_interval: float = 0.05,
) -> Tuple[int, str]:
    """
    Executa um processo e faz streaming de stdout linha-a-linha de forma segura.
    - Lê stdout em UMA única thread e envia as linhas para a fila.
    - O thread principal consome a fila e chama on_line (se houver).
    - Se cancel_check() ficar True: terminate().
    Retorna (returncode, stdout_total).
    """
    proc = subprocess.Popen(
        cmd,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        creationflags=creationflags,
        bufsize=1,
        universal_newlines=True,
    )

    if register_proc:
        try:
            register_proc(proc)
        except Exception:
            pass

    q: "queue.Queue[Optional[str]]" = queue.Queue()
    stdout_chunks: List[str] = []

    def _reader():
        try:
            if proc.stdout is None:
                q.put(None)
                return
            for line in proc.stdout:
                q.put(line)
        finally:
            q.put(None)

    t = threading.Thread(target=_reader, daemon=True)
    t.start()

    # Consome linhas e permite cancelamento
    done = False
    while not done:
        if cancel_check and cancel_check():
            try:
                proc.terminate()
            except Exception:
                pass
            # drena um pouco e sai
            # (a leitura pode ainda colocar linhas na fila)
        try:
            item = q.get(timeout=poll_interval)
        except queue.Empty:
            # verifica se já terminou
            if proc.poll() is not None:
                # ainda pode ter itens na fila, tenta drenar
                continue
            continue

        if item is None:
            done = True
            break

        stdout_chunks.append(item)
        if on_line:
            try:
                on_line(item.rstrip("\n"))
            except Exception:
                # callback não pode derrubar o runner
                pass

    # garante que terminou
    try:
        proc.wait(timeout=5)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass

    returncode = proc.returncode if proc.returncode is not None else -1
    stdout_total = "".join(stdout_chunks)
    return returncode, stdout_total
