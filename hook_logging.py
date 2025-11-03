import logging
import logging.config
from pathlib import Path
import sys

try:
    # Determina o caminho da config mesmo dentro do bundle
    if getattr(sys, 'frozen', False):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(__file__).parent

    log_config = base_path / "logging.ini"
    if log_config.exists():
        logging.config.fileConfig(log_config, disable_existing_loggers=False)
    else:
        logging.basicConfig(filename="app_log.txt", level=logging.DEBUG,
                            format="%(asctime)s - %(levelname)s - %(message)s")

    logging.info("Logging inicializado com sucesso.")
except Exception as e:
    print("Falha ao iniciar logging:", e)
