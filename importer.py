"""Import roster and Google Form exports into the local database."""

import csv
import re
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Optional

import jwt

from config import JWT_SECRET_KEY
from database import PinHistory, SessionLocal, Student, Submission


def active_graduation_years(as_of: Optional[date] = None) -> list[int]:
    """Return the four active graduating classes using a July 1 rollover."""
    today = as_of or date.today()
    senior_year = today.year + 1 if today.month >= 7 else today.year
    return [senior_year + offset for offset in range(4)]


def _rows_from_file(path: Any) -> Iterable[dict[str, Any]]:
    path = Path(path)
    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as file:
            yield from csv.DictReader(file)
        return
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        from openpyxl import load_workbook

        workbook = load_workbook(path, read_only=True, data_only=True)
        sheet = workbook.active
        rows = sheet.iter_rows(values_only=True)
        headers = [str(value or "").strip() for value in next(rows)]
        for values in rows:
            yield dict(zip(headers, values))
        return
    raise ValueError("Supported files are CSV, XLSX, and XLSM")


def _value(row: dict[str, Any], *names: str) -> str:
    normalized = {re.sub(r"[^a-z0-9]", "", str(key).lower()): value for key, value in row.items()}
    for name in names:
        value = normalized.get(re.sub(r"[^a-z0-9]", "", name.lower()))
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""


def import_student_directory(path: Any, session=None) -> int:
    """Upsert roster records and graduate active students missing from the roster."""
    owns_session = session is None
    session = session or SessionLocal()
    seen_emails: set[str] = set()
    try:
        for row in _rows_from_file(path):
            email = _value(row, "email", "email address").lower()
            if not email:
                continue
            graduation_year = _value(row, "graduation year", "class year", "expected graduation")
            is_active = True
            if graduation_year:
                try:
                    is_active = int(float(graduation_year)) in active_graduation_years()
                except ValueError:
                    pass
            seen_emails.add(email)
            student = session.query(Student).filter_by(email=email).one_or_none()
            if student is None:
                student = Student(email=email)
                session.add(student)
            student.student_id = _value(row, "student id", "student number", "id") or student.student_id
            student.full_name = _value(row, "full name", "name", "student name") or student.full_name
            student.grade_level = _value(row, "grade", "grade level") or student.grade_level
            student.counselor_email = _value(row, "counselor email", "counselor") or student.counselor_email
            student.parent_email = _value(row, "parent email", "parent email address", "guardian email") or student.parent_email
            rubric_score = _value(row, "owl rubric score", "owl score", "rubric score")
            if rubric_score:
                try:
                    student.owl_rubric_score = int(float(rubric_score))
                except ValueError:
                    pass
            student.status = "ACTIVE" if is_active else "GRADUATED"
        for student in session.query(Student).filter_by(status="ACTIVE"):
            if student.email not in seen_emails:
                student.status = "GRADUATED"
        if owns_session:
            session.commit()
        return len(seen_emails)
    except Exception:
        if owns_session:
            session.rollback()
        raise
    finally:
        if owns_session:
            session.close()


def rollover_students(path: Any, session=None) -> dict[str, int]:
    """Advance active students one grade and append students from the new roster.

    The roster is used for new students and updated directory details. Existing
    active students are advanced from their stored grade; grade 12 students
    become graduated after rollover.
    """
    owns_session = session is None
    session = session or SessionLocal()
    added = 0
    updated = 0
    graduated = 0
    try:
        rows = list(_rows_from_file(path))
        roster = {}
        for row in rows:
            email = _value(row, "email", "email address").lower()
            if email:
                roster[email] = row

        for student in session.query(Student).filter_by(status="ACTIVE").all():
            try:
                current_grade = int(str(student.grade_level).strip())
            except (TypeError, ValueError):
                current_grade = None
            if current_grade is None:
                continue
            if current_grade >= 12:
                student.status = "GRADUATED"
                graduated += 1
            else:
                student.grade_level = str(current_grade + 1)
                updated += 1
            row = roster.get(student.email)
            if row:
                student.full_name = _value(row, "full name", "name", "student name") or student.full_name
                student.counselor_email = _value(row, "counselor email", "counselor") or student.counselor_email
                student.parent_email = _value(row, "parent email", "parent email address", "guardian email") or student.parent_email

        existing_emails = {student.email for student in session.query(Student)}
        for email, row in roster.items():
            if email in existing_emails:
                continue
            student = Student(email=email, student_id=_value(row, "student id", "student number", "id"),
                              full_name=_value(row, "full name", "name", "student name"),
                              grade_level=_value(row, "grade", "grade level"),
                              counselor_email=_value(row, "counselor email", "counselor"), status="ACTIVE")
            student.parent_email = _value(row, "parent email", "parent email address", "guardian email")
            rubric_score = _value(row, "owl rubric score", "owl score", "rubric score")
            if rubric_score:
                try:
                    student.owl_rubric_score = int(float(rubric_score))
                except ValueError:
                    pass
            session.add(student)
            added += 1
        if owns_session:
            session.commit()
        return {"added": added, "updated": updated, "graduated": graduated}
    except Exception:
        if owns_session:
            session.rollback()
        raise
    finally:
        if owns_session:
            session.close()


def import_form_responses(path: Any, session=None) -> int:
    """Import one submission per response row from a CSV/XLSX export."""
    owns_session = session is None
    session = session or SessionLocal()
    imported = 0
    try:
        for row in _rows_from_file(path):
            email = _value(row, "email", "email address").lower()
            school_year = _value(row, "school year", "academic year")
            hours_text = _value(row, "hours", "hours submitted", "service hours")
            if not email or not school_year or not hours_text:
                continue
            try:
                hours = float(hours_text)
            except ValueError:
                continue
            student = session.query(Student).filter_by(email=email).one_or_none()
            if student is None:
                student = Student(email=email, status="ACTIVE")
                session.add(student)
                session.flush()
            session.add(Submission(student_id=student.id, school_year=school_year, hours_submitted=hours,
                                   file_upload_url=_value(row, "file upload url", "file upload")))
            imported += 1
        if owns_session:
            session.commit()
        return imported
    except Exception:
        if owns_session:
            session.rollback()
        raise
    finally:
        if owns_session:
            session.close()


def import_counselor_form_responses(path: Any, session=None) -> int:
    """Import Google Form counselor responses using the exported column names."""
    owns_session = session is None
    session = session or SessionLocal()
    imported = 0
    try:
        for row in _rows_from_file(path):
            student_id = _value(row, "student id")
            student_name = _value(row, "student name")
            writeup = _value(row, "owl pin blurb")
            token = _value(row, "jwt (do not edit)", "jwt")
            confirmation = _value(row, "confirm").lower()
            if not student_id or not writeup or "nominate" not in confirmation:
                continue
            student = session.query(Student).filter_by(student_id=student_id).one_or_none()
            if student is None and student_name:
                student = session.query(Student).filter_by(full_name=student_name).one_or_none()
            if student is None:
                continue
            if token:
                try:
                    payload = jwt.decode(token, JWT_SECRET_KEY, algorithms=["HS256"])
                    if int(payload.get("student_id")) != student.id:
                        continue
                except (jwt.PyJWTError, TypeError, ValueError):
                    continue
            record = (session.query(PinHistory).filter_by(student_id=student.id, pin_awarded="Owl")
                      .order_by(PinHistory.id.desc()).first())
            if record is None:
                record = PinHistory(student_id=student.id, academic_year="", pin_awarded="Owl")
                session.add(record)
            record.owl_writeup = writeup
            record.owl_requested = True
            imported += 1
        if owns_session:
            session.commit()
        return imported
    except Exception:
        if owns_session:
            session.rollback()
        raise
    finally:
        if owns_session:
            session.close()
