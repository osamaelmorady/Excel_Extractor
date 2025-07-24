import logging
import yaml
from pathlib import Path

class LoggerHandler:
    def __init__(self, config_path: str, name: str = "AutomationLogger"):
        self.config_path = config_path
        self.logger = self._setup_logger(name)

    def _get_log_file_path(self) -> Path:
        with open(self.config_path) as f:
            config_yaml = yaml.safe_load(f)
        log_path = Path(config_yaml['log_file'])
        log_path.parent.mkdir(parents=True, exist_ok=True)
        return log_path

    def _setup_logger(self, name: str) -> logging.Logger:
        logger = logging.getLogger(name)
        if logger.handlers:
            return logger  # Already configured

        log_path = self._get_log_file_path()

        logger.setLevel(logging.INFO)
        fmt = logging.Formatter("%(asctime)s %(levelname)s: %(message)s")

        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)

        ch = logging.StreamHandler()
        ch.setFormatter(fmt)
        logger.addHandler(ch)

        return logger