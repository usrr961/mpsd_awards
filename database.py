"""Shared SQLAlchemy models and session helpers for the CSA application."""

from datetime import datetime, timezone
from typing import Optional

from sqlalchemy import Boolean, DateTime, Float, ForeignKey, Integer, String, Text, create_engine, inspect, text
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship, sessionmaker

from config import DATABASE_URL


class Base(DeclarativeBase):
    pass


class Student(Base):
    __tablename__ = "students"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    email: Mapped[str] = mapped_column(String(255), unique=True, index=True)
    student_id: Mapped[Optional[str]] = mapped_column(String(100), nullable=True)
    full_name: Mapped[str] = mapped_column(String(255), default="")
    grade_level: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    counselor_email: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    parent_email: Mapped[Optional[str]] = mapped_column(String(255), nullable=True)
    owl_rubric_score: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    owl_selected: Mapped[bool] = mapped_column(Boolean, default=False)
    owl_override: Mapped[Optional[bool]] = mapped_column(Boolean, nullable=True)
    admin_award_override: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="ACTIVE")

    submissions: Mapped[list["Submission"]] = relationship(back_populates="student", cascade="all, delete-orphan")
    pin_history: Mapped[list["PinHistory"]] = relationship(back_populates="student", cascade="all, delete-orphan")
    award_override: Mapped[Optional["AwardOverride"]] = relationship(back_populates="student", uselist=False, cascade="all, delete-orphan")


class Submission(Base):
    __tablename__ = "submissions"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    student_id: Mapped[int] = mapped_column(ForeignKey("students.id"), index=True)
    school_year: Mapped[str] = mapped_column(String(20), index=True)
    hours_submitted: Mapped[float] = mapped_column(Float, default=0.0)
    file_upload_url: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))

    student: Mapped[Student] = relationship(back_populates="submissions")


class PinHistory(Base):
    __tablename__ = "pin_history"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    student_id: Mapped[int] = mapped_column(ForeignKey("students.id"), index=True)
    academic_year: Mapped[str] = mapped_column(String(20), index=True)
    pin_awarded: Mapped[str] = mapped_column(String(50))
    owl_writeup: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    owl_requested: Mapped[bool] = mapped_column(Boolean, default=False)

    student: Mapped[Student] = relationship(back_populates="pin_history")


class AwardOverride(Base):
    __tablename__ = "award_overrides"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    student_id: Mapped[int] = mapped_column(ForeignKey("students.id"), unique=True, index=True)
    award: Mapped[str] = mapped_column(String(50))
    standard_award: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    owl_awarded: Mapped[bool] = mapped_column(Boolean, default=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    student: Mapped[Student] = relationship(back_populates="award_override")


engine = create_engine(DATABASE_URL, future=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, expire_on_commit=False)


def init_db() -> None:
    Base.metadata.create_all(engine)
    # create_all does not add columns to an existing SQLite database.
    if "owl_rubric_score" not in {column["name"] for column in inspect(engine).get_columns("students")}:
        with engine.begin() as connection:
            connection.execute(text("ALTER TABLE students ADD COLUMN owl_rubric_score INTEGER"))
    columns = {column["name"] for column in inspect(engine).get_columns("students")}
    with engine.begin() as connection:
        if "parent_email" not in columns:
            connection.execute(text("ALTER TABLE students ADD COLUMN parent_email VARCHAR(255)"))
        if "owl_selected" not in columns:
            connection.execute(text("ALTER TABLE students ADD COLUMN owl_selected BOOLEAN DEFAULT 0"))
        if "owl_override" not in columns:
            connection.execute(text("ALTER TABLE students ADD COLUMN owl_override BOOLEAN"))
        if "admin_award_override" not in columns:
            connection.execute(text("ALTER TABLE students ADD COLUMN admin_award_override VARCHAR(50)"))
    override_columns = {column["name"] for column in inspect(engine).get_columns("award_overrides")}
    with engine.begin() as connection:
        if "standard_award" not in override_columns:
            connection.execute(text("ALTER TABLE award_overrides ADD COLUMN standard_award VARCHAR(50)"))
        if "owl_awarded" not in override_columns:
            connection.execute(text("ALTER TABLE award_overrides ADD COLUMN owl_awarded BOOLEAN DEFAULT 0"))
    # Preserve overrides created by older versions while moving future changes
    # into the dedicated override table.
    with SessionLocal() as session:
        existing_override_ids = {row.student_id for row in session.query(AwardOverride.student_id).all()}
        legacy = session.query(Student).filter(Student.admin_award_override.isnot(None)).all()
        for student in legacy:
            if student.id not in existing_override_ids:
                session.add(AwardOverride(student_id=student.id, award=student.admin_award_override))
                existing_override_ids.add(student.id)
        legacy_owl = session.query(Student).filter(Student.owl_override.is_(True)).all()
        for student in legacy_owl:
            if student.id not in existing_override_ids:
                session.add(AwardOverride(student_id=student.id, award="Owl", owl_awarded=True))
                existing_override_ids.add(student.id)
        for override in session.query(AwardOverride).all():
            if override.award == "Owl" and not override.owl_awarded:
                override.owl_awarded = True
            if override.standard_award is None and override.award in {"First Year", "Bronze", "Silver", "Gold"}:
                override.standard_award = override.award
        if legacy or legacy_owl:
            session.commit()


def get_session():
    return SessionLocal()


init_db()
