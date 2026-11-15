"""Backward-compatible imports for older code using models.py."""

from database import Base, PinHistory, SessionLocal, Student, Submission, engine

__all__ = ["Base", "Student", "Submission", "PinHistory", "engine", "SessionLocal"]