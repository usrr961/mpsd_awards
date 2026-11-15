"""Settings that should be reviewed before running the CSA application."""

import os
from pathlib import Path

BASE_DIRECTORY = Path(__file__).resolve().parent
DATABASE_URL = f"sqlite:///{BASE_DIRECTORY / 'csa_database.db'}"
JWT_SECRET_KEY = "ChrToqMftSAEdWuKeGo6ee4JLW8wUjOeK2NLTr2swRJ"
COUNSELOR_GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLScxrnpZk4X0ttQQwpiubsPHrApCACXcL-oivvafYOvdz8mwow/viewform?usp=pp_url"
COUNSELOR_FORM_STUDENT_ID_ENTRY = "2053230864"
COUNSELOR_FORM_STUDENT_NAME_ENTRY = "1912836676"
COUNSELOR_FORM_JWT_ENTRY = "1416663386"
COUNSELOR_FORM_CONFIRM_ENTRY = "882323090"
COUNSELOR_CONFIRMATION = "Yes, I nominate this student for an Community Service Awards Owl pin."
AWARD_MESSAGE_FILE = BASE_DIRECTORY / "data" / "award_message.txt"
WEB_RESOURCES_FILE = BASE_DIRECTORY / "data" / "web_resources.json"
OWL_TOP_PERCENTAGE = 20.0

# Award thresholds are cumulative hours across grades 9-12.
AWARD_THRESHOLDS = {
    "First Year": 10.0,
    "Bronze": 20.0,
    "Silver": 30.0,
    "Gold": 40.0,
}

# When true, each award requires the immediately preceding award in pin history.
# First Year is the entry award and can be earned without a prior pin.
AWARD_REQUIRES_PREVIOUS = {
    "First Year": False,
    "Bronze": True,
    "Silver": True,
    "Gold": True,
}

SMTP_SERVER = os.getenv("CSA_SMTP_SERVER", "smtp.ethereal.email")
SMTP_PORT = int(os.getenv("CSA_SMTP_PORT", "587"))
SMTP_USERNAME = os.getenv("CSA_SMTP_USERNAME", "")
SMTP_PASSWORD = os.getenv("CSA_SMTP_PASSWORD", "")
SENDER_EMAIL = os.getenv("CSA_SENDER_EMAIL", SMTP_USERNAME)
SMTP_USE_TLS = os.getenv("CSA_SMTP_USE_TLS", "true").lower() in {"1", "true", "yes", "on"}
SMTP_USE_SSL = os.getenv("CSA_SMTP_USE_SSL", "false").lower() in {"1", "true", "yes", "on"}
SMTP_TIMEOUT = int(os.getenv("CSA_SMTP_TIMEOUT", "30"))

EXPORT_DIRECTORY = Path("exports")
FILE_PRODUCER = os.getenv("CSA_FILE_PRODUCER", "Community Service Awards Manager")
FILE_PUBLISHER = os.getenv("CSA_FILE_PUBLISHER", "Montclair Public Schools")
DASHBOARD_DEFAULT_WIDTH = int(os.getenv("CSA_DASHBOARD_WIDTH", "1000"))
DASHBOARD_DEFAULT_HEIGHT = int(os.getenv("CSA_DASHBOARD_HEIGHT", "900"))
