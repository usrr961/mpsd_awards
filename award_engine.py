"""Calculate cumulative service awards and maintain pin history."""

from collections import defaultdict
from typing import Optional

from config import AWARD_REQUIRES_PREVIOUS, AWARD_THRESHOLDS
from database import AwardOverride, PinHistory, SessionLocal, Student
from sqlalchemy.orm import selectinload


AWARD_ORDER = tuple(AWARD_THRESHOLDS)


def award_for_hours(total_hours: float, earned_awards=None) -> Optional[str]:
    """Return the highest newly eligible sequential award for total hours."""
    earned_awards = set(earned_awards or ())
    for award in reversed(AWARD_ORDER):
        if total_hours < AWARD_THRESHOLDS[award]:
            continue
        if AWARD_REQUIRES_PREVIOUS.get(award, False):
            index = AWARD_ORDER.index(award)
            if index and AWARD_ORDER[index - 1] not in earned_awards:
                continue
        return award
    return None


def calculate_awards(session=None) -> list[dict]:
    """Recalculate yearly pin records and return dashboard-ready summaries."""
    owns_session = session is None
    session = session or SessionLocal()
    summaries = []
    try:
        students = (session.query(Student)
            .options(selectinload(Student.submissions), selectinload(Student.pin_history), selectinload(Student.award_override))
                .order_by(Student.full_name, Student.email))
        for student in students:
            by_year = defaultdict(float)
            for submission in student.submissions:
                by_year[submission.school_year] += submission.hours_submitted
            cumulative = 0.0
            yearly = []
            earned_awards = {record.pin_awarded for record in student.pin_history}
            for academic_year in sorted(by_year):
                hours = by_year[academic_year]
                cumulative += hours
                pin = award_for_hours(cumulative, earned_awards)
                if pin:
                    record = (session.query(PinHistory)
                              .filter_by(student_id=student.id, academic_year=academic_year)
                              .filter(PinHistory.pin_awarded != "Owl")
                              .order_by(PinHistory.id)
                              .first())
                    if record is None:
                        record = PinHistory(student_id=student.id, academic_year=academic_year, pin_awarded=pin)
                        session.add(record)
                    else:
                        record.pin_awarded = pin
                    yearly.append({"academic_year": academic_year, "hours": hours, "pin": pin})
                    earned_awards.add(pin)
            if student.owl_selected and yearly:
                owl_year = yearly[-1]["academic_year"]
                owl_record = (session.query(PinHistory)
                              .filter_by(student_id=student.id, academic_year=owl_year, pin_awarded="Owl")
                              .order_by(PinHistory.id)
                              .first())
                if owl_record is None:
                    owl_record = PinHistory(student_id=student.id, academic_year=owl_year,
                                            pin_awarded="Owl", owl_requested=False)
                    student.pin_history.append(owl_record)
                    session.add(owl_record)
            hour_pin = yearly[-1]["pin"] if yearly else ""
            owl_history = any(record.pin_awarded == "Owl" for record in student.pin_history)
            current_pin = "Owl" if student.owl_selected or owl_history else hour_pin
            standard_override = None
            owl_override = False
            if student.award_override:
                standard_override = student.award_override.standard_award
                owl_override = student.award_override.owl_awarded or student.award_override.award == "Owl"
                if standard_override is None and student.award_override.award in AWARD_ORDER:
                    standard_override = student.award_override.award
                if standard_override:
                    current_pin = standard_override
                if owl_override:
                    current_pin = "Owl"
            award_label = current_pin
            if standard_override and owl_override:
                award_label = f"{standard_override} + Owl"
            summaries.append({"student": student, "total_hours": cumulative, "yearly": yearly,
                              "current_pin": current_pin, "hour_pin": hour_pin,
                              "award_label": award_label,
                              "admin_award_override": standard_override,
                              "admin_owl_override": owl_override})
        if owns_session:
            session.commit()
        return summaries
    except Exception:
        if owns_session:
            session.rollback()
        raise
    finally:
        if owns_session:
            session.close()
