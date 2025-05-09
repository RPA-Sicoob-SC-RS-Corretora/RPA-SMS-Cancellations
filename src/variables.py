import os
import configparser
from dotenv import load_dotenv

load_dotenv()

_config = configparser.ConfigParser()
_config.read("config.ini", encoding="utf-8")

REQUIRED_KEYS = ["message_crediauc", "message_credisc", "file_path"]
for key in REQUIRED_KEYS:
    if not _config["DEFAULT"].get(key):
        raise ValueError(f"{key} is empty or missing, please enter a value in the config file")

if not os.getenv("URL_API"):
    raise ValueError("URL API is empty, please set the URL_API environment variable")

VARIABLES = {
    "AUTH_KEY": os.getenv("AUTH_TOKEN"),
    "URL": os.getenv("URL_API"),
    "MESSAGE_CREDIAUC": _config["DEFAULT"]["message_crediauc"],
    "MESSAGE_CREDISC": _config["DEFAULT"]["message_credisc"],
    "FILE_PATH": _config["DEFAULT"]["file_path"],
    "EN2": ["3067 SICOOB CREDIAUC", "3258 SICOOB CREDISC"],
    "REPORT_PATH": "relatorio_cancelamento.xlsx",
    "EMAIL_RECIPIENTS_PATH": os.path.join(_config["DEFAULT"]["file_path"], "Emails", "recipients.xlsx")
}
