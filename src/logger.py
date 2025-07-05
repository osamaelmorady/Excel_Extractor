# src/logger.py
import logging
import yaml

def setup_logger():
    with open("config/settings.yaml") as f:
        config = yaml.safe_load(f)
    
    logging.basicConfig(
        filename=config["log_file"],
        level=logging.INFO,
        format='%(asctime)s %(levelname)s:%(message)s'
    )
    return logging.getLogger("AutomationLogger")