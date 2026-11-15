"""SMTP email helpers for counselor requests and award notifications."""

import smtplib
from email.message import EmailMessage
from urllib.parse import quote

import jwt

from config import (COUNSELOR_CONFIRMATION, COUNSELOR_FORM_CONFIRM_ENTRY, COUNSELOR_FORM_JWT_ENTRY,
                    COUNSELOR_FORM_STUDENT_ID_ENTRY, COUNSELOR_FORM_STUDENT_NAME_ENTRY,
                    COUNSELOR_GOOGLE_FORM_URL, JWT_SECRET_KEY, SENDER_EMAIL, SMTP_PASSWORD,
                    SMTP_PORT, SMTP_SERVER, SMTP_TIMEOUT, SMTP_USE_SSL, SMTP_USE_TLS, SMTP_USERNAME)


def generate_counselor_token(student_id: int, counselor_email: str) -> str:
    """Create the token embedded in the counselor Google Form export."""
    return jwt.encode({"student_id": student_id, "counselor": counselor_email}, JWT_SECRET_KEY, algorithm="HS256")


def send_email(recipient: str, subject: str, body: str, *, settings=None) -> None:
    settings = settings or {
        "server": SMTP_SERVER,
        "port": SMTP_PORT,
        "username": SMTP_USERNAME,
        "password": SMTP_PASSWORD,
        "from_email": SENDER_EMAIL,
        "use_tls": SMTP_USE_TLS,
        "use_ssl": SMTP_USE_SSL,
        "timeout": SMTP_TIMEOUT,
    }
    message = EmailMessage()
    message["From"] = settings["from_email"]
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)
    smtp_class = smtplib.SMTP_SSL if settings.get("use_ssl") else smtplib.SMTP
    with smtp_class(settings["server"], int(settings["port"]), timeout=int(settings.get("timeout", 30))) as server:
        if settings.get("use_tls") and not settings.get("use_ssl"):
            server.starttls()
        username = settings.get("username")
        if username:
            server.login(username, settings.get("password", ""))
        server.send_message(message)


def test_smtp_connection(settings) -> None:
    """Connect and authenticate without sending a message."""
    smtp_class = smtplib.SMTP_SSL if settings.get("use_ssl") else smtplib.SMTP
    with smtp_class(settings["server"], int(settings["port"]), timeout=int(settings.get("timeout", 30))) as server:
        if settings.get("use_tls") and not settings.get("use_ssl"):
            server.starttls()
        if settings.get("username"):
            server.login(settings["username"], settings.get("password", ""))


def send_counselor_owl_request(student, dry_run: bool = False) -> str:
    if not student.counselor_email:
        raise ValueError(f"No counselor email for {student.full_name or student.email}")
    token = generate_counselor_token(student.id, student.counselor_email)
    link = build_counselor_google_form_url(student, token)
    body = (f"Dear Counselor,\n\n{student.full_name} has earned an Owl Pin. "
            f"Please submit a one-paragraph write-up here:\n{link}\n\nThank you,\nCSA Team")
    if not dry_run:
        send_email(student.counselor_email, f"CSA Owl Pin write-up for {student.full_name}", body)
    return link


def build_counselor_google_form_url(student, token: str) -> str:
    """Create the Google Form prefilled URL used in counselor emails."""
    return (f"{COUNSELOR_GOOGLE_FORM_URL}&entry.{COUNSELOR_FORM_STUDENT_ID_ENTRY}={quote(student.student_id or '')}"
            f"&entry.{COUNSELOR_FORM_STUDENT_NAME_ENTRY}={quote(student.full_name or '')}"
            f"&entry.{COUNSELOR_FORM_JWT_ENTRY}={quote(token)}"
            f"&entry.{COUNSELOR_FORM_CONFIRM_ENTRY}={quote(COUNSELOR_CONFIRMATION)}")


def send_award_result(recipient: str, student_name: str, award: str, dry_run: bool = False) -> None:
    if not dry_run:
        send_email(recipient, f"CSA award result for {student_name}",
                   f"{student_name} has qualified for the {award} Community Service Award.")
