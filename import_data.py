"""Command-line fallback for importing roster and Google Form exports."""

import sys

from award_engine import calculate_awards
from database import SessionLocal
from importer import import_form_responses, import_student_directory


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: .venv/bin/python import_data.py ROSTER_FILE FORM_RESPONSES_FILE")
        print("Example: .venv/bin/python import_data.py roster.xlsx form_responses.csv")
        return 2

    roster_file, form_file = sys.argv[1:]
    roster_count = import_student_directory(roster_file)
    submission_count = import_form_responses(form_file)
    session = SessionLocal()
    try:
        summaries = calculate_awards(session)
        session.commit()
    finally:
        session.close()
    print(f"Imported {roster_count} roster students.")
    print(f"Imported {submission_count} service submissions.")
    print(f"Calculated awards for {len(summaries)} students.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
