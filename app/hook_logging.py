# app/hook_logging.py
from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
import sys
import threading

from app.paths import Paths


def setup_logging(level: int = logging.INFO) -> None:
    """
    Logging robusto:
    - arquivo em AppData\\logs\\auto_cl.log
    - rotação (2MB, 5 backups)
    - thread excepthook
    - mantém loggers existentes
    """
    P = Paths.build()
    log_file = P.logs / "auto_cl.log"

    # raiz
    root = logging.getLogger()
    root.setLevel(level)

    # evita duplicar handlers se chamar duas vezes
    for h in list(root.handlers):
        root.removeHandler(h)

    fmt = logging.Formatter(
        fmt="%(asctime)s %(levelname)s %(name)s [%(process)d %(threadName)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = RotatingFileHandler(
        filename=str(log_file),
        mode="a",
        maxBytes=2 * 1024 * 1024,  # 2MB
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setLevel(level)
    file_handler.setFormatter(fmt)
    root.addHandler(file_handler)

    # console em DEV (opcional)
    if not getattr(sys, "frozen", False):
        console = logging.StreamHandler()
        console.setLevel(level)
        console.setFormatter(fmt)
        root.addHandler(console)

    # Exceções não tratadas em threads (Python 3.8+)
    def _thread_excepthook(args):
        logging.getLogger("thread").exception(
            "Exceção não tratada em thread",
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
        )

    try:
        threading.excepthook = _thread_excepthook  # type: ignore[attr-defined]
    except Exception:
        pass

    logging.getLogger(__name__).info("Logging inicializado: %s", log_file)
