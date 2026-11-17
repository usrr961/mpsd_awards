"""PyQt6 desktop dashboard for the Community Service Awards system."""

import csv
import json
import math
import re
from collections import Counter, defaultdict
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook
from PyQt6.QtCore import QTimer, QUrl, Qt
from PyQt6.QtGui import QAction, QColor
from PyQt6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QFileDialog, QFormLayout,
    QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPushButton, QPlainTextEdit, QSpinBox, QSplitter, QTabWidget,
    QTableWidget, QTableWidgetItem, QToolButton, QVBoxLayout, QWidget, QDialogButtonBox,
)
from PyQt6.QtGui import QDesktopServices
from sqlalchemy.orm import selectinload

from award_engine import AWARD_ORDER, calculate_awards
from config import (AWARD_MESSAGE_FILE, DASHBOARD_DEFAULT_HEIGHT, DASHBOARD_DEFAULT_WIDTH,
                    EXPORT_DIRECTORY, FILE_PRODUCER, FILE_PUBLISHER, OWL_TOP_PERCENTAGE,
                    SMTP_PASSWORD, SMTP_PORT,
                    SMTP_SERVER, SMTP_TIMEOUT, SMTP_USE_SSL, SMTP_USE_TLS, SENDER_EMAIL,
                    SMTP_USERNAME, WEB_RESOURCES_FILE)
from database import AwardOverride, PinHistory, SessionLocal, Student, Submission
from email_dispatcher import (build_counselor_google_form_url, generate_counselor_token,
                              send_email, test_smtp_connection)
from importer import (
    active_graduation_years, import_counselor_form_responses, import_form_responses,
    import_student_directory, rollover_students,
)
from web_server import LANPortal, local_ip_addresses


class GradeDialog(QDialog):
    def __init__(self, parent, title):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.checks = []
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Choose the grade levels to include:"))
        for grade in (9, 10, 11, 12):
            check = QCheckBox(f"Grade {grade}")
            check.setChecked(True)
            self.checks.append(check)
            layout.addWidget(check)
        buttons = QHBoxLayout()
        export = QPushButton("Use Selected Grades")
        export.clicked.connect(self.accept)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        buttons.addWidget(cancel)
        buttons.addWidget(export)
        layout.addLayout(buttons)

    def grades(self):
        return {check.text().split()[-1] for check in self.checks if check.isChecked()}


class CollapsibleSection(QWidget):
    def __init__(self, title, parent=None, expanded=True):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.toggle = QToolButton()
        self.toggle.setText(title)
        self.toggle.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.toggle.setArrowType(Qt.ArrowType.DownArrow if expanded else Qt.ArrowType.RightArrow)
        self.toggle.setCheckable(True)
        self.toggle.setChecked(expanded)
        self.toggle.clicked.connect(self._toggle_content)
        layout.addWidget(self.toggle)
        self.content = QWidget()
        self.content.setVisible(expanded)
        layout.addWidget(self.content)

    def _toggle_content(self, expanded):
        self.content.setVisible(expanded)
        self.toggle.setArrowType(Qt.ArrowType.DownArrow if expanded else Qt.ArrowType.RightArrow)


class EmailComposerDialog(QDialog):
    def __init__(self, parent, title, rows, recipient_field, subject, body, smtp_settings):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(820, 700)
        self.rows = rows
        self.smtp_settings = smtp_settings
        self.fields = sorted({key for row in rows for key in row})
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.recipient = QComboBox()
        self.recipient.addItems(self.fields)
        if recipient_field in self.fields:
            self.recipient.setCurrentText(recipient_field)
        self.subject = QLineEdit(subject)
        self.body = QPlainTextEdit()
        self.body.setPlainText(body)
        self.preview_row = QComboBox()
        self.preview_row.addItems([self._label(row) for row in rows])
        form.addRow("Recipient column", self.recipient)
        form.addRow("Subject template", self.subject)
        form.addRow("Preview row", self.preview_row)
        layout.addLayout(form)
        layout.addWidget(QLabel("Message template. Use {column_name} tokens from the recipient data."))
        layout.addWidget(self.body, 1)
        preview_group = QGroupBox("Rendered email preview")
        preview_layout = QVBoxLayout(preview_group)
        self.preview = QPlainTextEdit()
        self.preview.setReadOnly(True)
        preview_layout.addWidget(self.preview)
        layout.addWidget(preview_group, 1)
        buttons = QDialogButtonBox()
        send = buttons.addButton("Send to All", QDialogButtonBox.ButtonRole.AcceptRole)
        cancel = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
        send.clicked.connect(self.send_all)
        cancel.clicked.connect(self.reject)
        layout.addWidget(buttons)
        self.recipient.currentTextChanged.connect(self.update_preview)
        self.subject.textChanged.connect(self.update_preview)
        self.body.textChanged.connect(self.update_preview)
        self.preview_row.currentIndexChanged.connect(self.update_preview)
        self.update_preview()

    @staticmethod
    def _label(row):
        return f"{row.get('student_name') or row.get('recipient_email') or 'Unnamed'} <{row.get('recipient_email', '')}>"

    @staticmethod
    def _render(template, row):
        return re.sub(r"\{([^{}]+)\}", lambda match: str(row.get(match.group(1), "")), template)

    def update_preview(self):
        if not self.rows:
            self.preview.setPlainText("No rows are available for this send.")
            return
        row = self.rows[self.preview_row.currentIndex()]
        recipient = self._render("{" + self.recipient.currentText() + "}", row).strip()
        subject = self._render(self.subject.text(), row)
        body = self._render(self.body.toPlainText(), row)
        self.preview.setPlainText(f"To: {recipient}\nSubject: {subject}\n\n{body}")

    def send_all(self):
        recipient_field = self.recipient.currentText()
        rendered = []
        for row in self.rows:
            recipient = str(row.get(recipient_field, "")).strip()
            if not recipient:
                continue
            rendered.append((recipient, self._render(self.subject.text(), row), self._render(self.body.toPlainText(), row)))
        if not rendered:
            QMessageBox.warning(self, "No recipients", f"No rows contain a value in '{recipient_field}'.")
            return
        answer = QMessageBox.question(self, "Confirm send", f"Send {len(rendered)} email(s) now?")
        if answer != QMessageBox.StandardButton.Yes:
            return
        failures = []
        for recipient, subject, body in rendered:
            try:
                send_email(recipient, subject, body, settings=self.smtp_settings)
            except Exception as error:
                failures.append(f"{recipient}: {error}")
        if failures:
            QMessageBox.critical(self, "Send completed with errors", "\n".join(failures[:10]))
        else:
            QMessageBox.information(self, "Send complete", f"Sent {len(rendered)} email(s).")
            self.accept()


class CSADashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSA Management System - Montclair Public Schools")
        self.summaries = []
        self.rubric_percentage = float(OWL_TOP_PERCENTAGE)
        self._launch_size_applied = False
        self.portal = LANPortal()
        self._build_ui()
        available_height = QApplication.primaryScreen().availableGeometry().height()
        self._launch_height = min(DASHBOARD_DEFAULT_HEIGHT, max(300, available_height - 40))
        self.setMinimumHeight(self._launch_height)
        self.resize(DASHBOARD_DEFAULT_WIDTH, self._launch_height)
        self.refresh_all()
        self._sync_timer = QTimer(self)
        self._sync_timer.setInterval(2000)
        self._sync_timer.timeout.connect(self.refresh_all)
        self._sync_timer.start()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._launch_size_applied:
            QTimer.singleShot(0, lambda: QTimer.singleShot(0, self._apply_launch_size))

    def _apply_launch_size(self):
        if not self._launch_size_applied:
            self.resize(DASHBOARD_DEFAULT_WIDTH, self._launch_height)
            self._launch_size_applied = True

    def _build_ui(self):
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #eef2f5; color: #17202a; }
            QGroupBox { background: white; border: 1px solid #c8d3dc; border-radius: 4px; margin-top: 12px; padding: 12px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; font-weight: bold; color: #183b56; }
            QPushButton { background: #2f6f9f; color: white; border: 0; padding: 6px 8px; min-height: 28px; font-size: 11px; border-radius: 3px; }
            QPushButton:hover { background: #4b91c1; }
            QLineEdit, QPlainTextEdit, QComboBox, QSpinBox { background: white; color: #17202a; border: 1px solid #aebdca; padding: 5px; }
            QTabWidget::pane { border: 1px solid #c8d3dc; background: #eef2f5; }
            QTabBar::tab { background: #d6e0e7; padding: 10px 18px; margin-right: 2px; }
            QTabBar::tab:selected { background: #2f6f9f; color: white; }
            QHeaderView::section { background: #d6e0e7; color: #17202a; padding: 7px; font-weight: bold; border: 0; }
            QTableWidget { background: white; alternate-background-color: #f6f9fb; gridline-color: #d7e0e8; }
        """)
        central = QWidget()
        root = QVBoxLayout(central)
        header = QWidget()
        header.setStyleSheet("background:#183b56;")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(16, 8, 12, 8)
        title = QLabel("CSA Management System  |  Montclair Public Schools")
        title.setStyleSheet("color:white; font-size:22px; font-weight:bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        about_button = QPushButton("About")
        about_button.setStyleSheet("QPushButton { background:#315d7c; color:white; border:1px solid #9bc1d8; padding:7px 14px; border-radius:4px; } QPushButton:hover { background:#467d9e; }")
        about_button.clicked.connect(self.show_about)
        header_layout.addWidget(about_button)
        resources_button = QPushButton("Web Resources")
        resources_button.setStyleSheet("QPushButton { background:#39765f; color:white; border:1px solid #a8dec3; padding:7px 14px; border-radius:4px; } QPushButton:hover { background:#4d9477; }")
        resources_button.clicked.connect(self.show_web_resources)
        header_layout.addWidget(resources_button)
        root.addWidget(header)
        self.tabs = QTabWidget()
        self.tabs.addTab(self._imports_tab(), "1. Data Imports")
        self.tabs.addTab(self._overview_tab(), "2. Student Overview")
        self.tabs.addTab(self._history_tab(), "3. Pin History")
        self.tabs.addTab(self._counselor_tab(), "4. Counselor Owl Hub")
        self.tabs.addTab(self._stats_tab(), "5. Award Stats")
        self.tabs.addTab(self._email_tab(), "6. Email Center")
        root.addWidget(self.tabs)
        self.status = QLabel("SQLite database")
        root.addWidget(self.status)
        self.setCentralWidget(central)

        refresh_action = QAction("Refresh", self)
        refresh_action.triggered.connect(self.refresh_all)
        self.menuBar().addAction(refresh_action)

    def show_about(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("About MPS CSA Email Blast Manager")
        dialog.setMinimumWidth(560)
        layout = QVBoxLayout(dialog)
        about = QLabel(
            "<h2>MPS CSA Email Blast Manager</h2>"
            "<p><b>Developer:</b> Rayan Nasir "
            "(<a href='https://github.com/usrr961'>@usrr961</a>)</p>"
            "<p><b>Version:</b> 2.4.0</p>"
            "<p>Automated utility built to streamline and manage individualized school "
            "email blasts for Community Service Awards (CSA). Dynamically syncs high "
            "school student directories and manages custom communications.</p>"
            "<hr>"
            "<p><b>CSA Google Sheet:</b><br>"
            "<a href='https://docs.google.com/spreadsheets/d/1QqvtbFZAvZf2HWrLOggJQt6xNI_hBJuO5mM6B8dEukQ/edit'>"
            "Open the Community Service Awards Google Sheet</a></p>"
        )
        about.setOpenExternalLinks(False)
        about.linkActivated.connect(lambda url: QDesktopServices.openUrl(QUrl(url)) )
        about.setWordWrap(True)
        layout.addWidget(about)
        data_actions = QGroupBox("Database Data")
        data_layout = QVBoxLayout(data_actions)
        data_layout.addWidget(QLabel("Use these actions to reset the local database or restore the sample dataset included with the application."))
        data_buttons = QHBoxLayout()
        data_buttons.addWidget(self._button("Clear All Student Data", self.clear_database))
        data_buttons.addWidget(self._button("Load Current Sample Data", self.load_sample_data))
        data_layout.addLayout(data_buttons)
        layout.addWidget(data_actions)
        close_button = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close_button.rejected.connect(dialog.reject)
        layout.addWidget(close_button)
        dialog.exec()

    def clear_database(self):
        answer = QMessageBox.warning(
            self,
            "Clear database",
            "This permanently deletes all students, service submissions, and pin history from the local database. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        session = SessionLocal()
        try:
            submission_count = session.query(Submission).delete(synchronize_session=False)
            pin_count = session.query(PinHistory).delete(synchronize_session=False)
            session.query(AwardOverride).delete(synchronize_session=False)
            student_count = session.query(Student).delete(synchronize_session=False)
            session.commit()
        except Exception as error:
            session.rollback()
            QMessageBox.critical(self, "Clear failed", str(error))
            return
        finally:
            session.close()
        self.refresh_all()
        QMessageBox.information(self, "Database cleared", f"Deleted {student_count} students, {submission_count} submissions, and {pin_count} pin records.")

    def load_sample_data(self):
        answer = QMessageBox.warning(
            self,
            "Load sample data",
            "This replaces all current local database data with the sample dataset included with the application. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        data_directory = Path(__file__).resolve().parent / "data"
        roster_path = data_directory / "sample_roster.csv"
        hours_path = data_directory / "sample_form_responses.csv"
        counselor_path = data_directory / "sample_counselor_form_responses.csv"
        session = SessionLocal()
        try:
            session.query(Submission).delete(synchronize_session=False)
            session.query(PinHistory).delete(synchronize_session=False)
            session.query(AwardOverride).delete(synchronize_session=False)
            session.query(Student).delete(synchronize_session=False)
            session.commit()
            roster_count = import_student_directory(roster_path, session=session)
            session.flush()
            hours_count = import_form_responses(hours_path, session=session)
            counselor_count = import_counselor_form_responses(counselor_path, session=session)
            session.commit()
        except Exception as error:
            session.rollback()
            QMessageBox.critical(self, "Sample load failed", str(error))
            return
        finally:
            session.close()
        self.refresh_all()
        QMessageBox.information(self, "Sample data loaded", f"Loaded {roster_count} roster records, {hours_count} service submissions, and {counselor_count} counselor responses.")

    def show_web_resources(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("CSA Web Resources")
        dialog.setMinimumWidth(700)
        layout = QVBoxLayout(dialog)

        portal_group = QGroupBox("Shared Student Portal")
        portal_layout = QVBoxLayout(portal_group)
        description = QLabel("Start a shared web portal for devices connected to the same Wi-Fi network.")
        description.setWordWrap(True)
        portal_layout.addWidget(description)

        portal_settings = QFormLayout()
        portal_pin = QLineEdit(self.portal.pin)
        portal_pin.setEchoMode(QLineEdit.EchoMode.Password)
        portal_pin.setMaxLength(64)
        portal_settings.addRow("Portal password", portal_pin)
        portal_layout.addLayout(portal_settings)

        link_row = QHBoxLayout()
        portal_link = QLineEdit()
        portal_link.setReadOnly(True)
        copy_link = QPushButton("Copy Link")
        copy_link.setEnabled(False)
        link_row.addWidget(portal_link, 1)
        link_row.addWidget(copy_link)
        portal_layout.addLayout(link_row)

        pin_row = QHBoxLayout()
        portal_pin_display = QLineEdit()
        portal_pin_display.setReadOnly(True)
        copy_pin = QPushButton("Copy Password")
        copy_pin.setEnabled(False)
        pin_row.addWidget(portal_pin_display, 1)
        pin_row.addWidget(copy_pin)
        portal_layout.addLayout(pin_row)

        portal_toggle = QPushButton("Stop Shared Portal" if self.portal.server else "Start Shared Portal")
        portal_layout.addWidget(portal_toggle)

        def update_portal_details():
            if self.portal.server:
                addresses = local_ip_addresses()
                url = f"http://{addresses[0]}:{self.portal.port}" if addresses else f"http://localhost:{self.portal.port}"
                portal_link.setText(url)
                copy_link.setEnabled(True)
                portal_pin_display.setText(self.portal.pin)
                copy_pin.setEnabled(True)
                portal_toggle.setText("Stop Shared Portal")
            else:
                portal_link.clear()
                portal_link.setPlaceholderText("Start the portal to generate a link")
                copy_link.setEnabled(False)
                portal_pin_display.clear()
                portal_pin_display.setPlaceholderText("Portal password")
                copy_pin.setEnabled(False)
                portal_toggle.setText("Start Shared Portal")

        def save_pin():
            try:
                self.portal.set_pin(portal_pin.text())
                if self.portal.server:
                    update_portal_details()
                return True
            except ValueError as error:
                QMessageBox.warning(dialog, "Portal Password", str(error))
                return False

        def toggle_from_modal():
            portal_toggle.setEnabled(False)
            try:
                if self.portal.server:
                    self.portal.stop()
                else:
                    if not save_pin():
                        return
                    self.portal.start()
            except Exception as error:
                QMessageBox.critical(dialog, "Shared Portal Error", str(error))
            finally:
                update_portal_details()
                portal_toggle.setEnabled(True)

        portal_toggle.clicked.connect(toggle_from_modal)
        portal_pin.editingFinished.connect(save_pin)
        copy_link.clicked.connect(lambda: QApplication.clipboard().setText(portal_link.text()))
        copy_pin.clicked.connect(lambda: QApplication.clipboard().setText(portal_pin_display.text()))
        update_portal_details()
        layout.addWidget(portal_group)

        export_group = QGroupBox("Desktop File Exports")
        export_layout = QGridLayout(export_group)
        export_layout.addWidget(self._button("Counselor Word + Excel", self.export_counselor_mail_merge), 0, 0)
        export_layout.addWidget(self._button("Student Word + Excel", self.export_student_mail_merge), 0, 1)
        export_layout.addWidget(self._button("Senior Word + Excel", self.export_senior_mail_merge), 0, 2)
        export_layout.addWidget(self._button("Student Awards CSV", self.export_awards), 1, 0)
        export_layout.addWidget(self._button("Counselor Autocrat CSV", self.export_counselor), 1, 1)
        export_layout.addWidget(self._button("Owl Rubric CSV", self.export_rubric), 1, 2)
        export_layout.addWidget(self._button("Pin History CSV", self.export_history), 2, 0)
        layout.addWidget(export_group)

        layout.addWidget(QLabel("Open a configured CSA web resource:"))
        try:
            resources = json.loads(Path(WEB_RESOURCES_FILE).read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            layout.addWidget(QLabel(f"Could not load web resources: {error}"))
            resources = {}
        for label, url in resources.items():
            link = QLabel(f"<a href='{url}'>{label}</a>")
            link.setOpenExternalLinks(False)
            link.linkActivated.connect(lambda clicked_url: QDesktopServices.openUrl(QUrl(clicked_url)))
            link.setStyleSheet("QLabel { background:white; border:1px solid #b7c9d5; padding:10px; color:#1c5d8a; }")
            layout.addWidget(link)
        close_button = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close_button.rejected.connect(dialog.reject)
        layout.addWidget(close_button)
        dialog.exec()

    def _button(self, text, slot):
        button = QPushButton(text)
        button.clicked.connect(slot)
        return button

    def _imports_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        actions = QGroupBox("Import and Rollover Actions")
        grid = QGridLayout(actions)
        grid.addWidget(self._button("Import Hours CSV/XLSX", self.import_hours), 0, 0)
        grid.addWidget(self._button("Import Student Roster", self.import_roster), 0, 1)
        grid.addWidget(self._button("Import Counselor Responses", self.import_counselor), 0, 2)
        grid.addWidget(self._button("Run School-Year Rollover", self.rollover), 1, 0)
        grid.addWidget(QLabel("Active graduation years: " + ", ".join(map(str, active_graduation_years()))), 1, 1, 1, 2)
        layout.addWidget(actions)
        expected = QGroupBox("CSV Templates and Expected Columns")
        text = QPlainTextEdit()
        text.setReadOnly(True)
        text.setPlainText(
            "Templates are in the data folder.\n\n"
            "Hours import: Email, School Year, Hours Submitted, File Upload URL\n"
            "Roster import: Email, Student ID, Full Name, Grade Level, Counselor Email, Graduation Year, Owl Rubric Score\n"
            "Counselor responses: Timestamp, Username, Student ID, Student Name, Owl Pin Blurb, JWT (DO NOT EDIT), Confirm\n\n"
            "Rollover advances active students one grade, graduates grade 12, and appends new roster students."
        )
        box = QVBoxLayout(expected)
        box.addWidget(text)
        layout.addWidget(expected)
        layout.addStretch()
        return tab

    def _overview_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        filters = QGroupBox("Student Directory Filters")
        row = QHBoxLayout(filters)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search by name, student ID, or email...")
        self.search.textChanged.connect(self.refresh_overview)
        self.grade_filter = QComboBox()
        self.grade_filter.addItems(["All Grades", "9", "10", "11", "12"])
        self.grade_filter.currentTextChanged.connect(self.refresh_overview)
        self.award_filter = QComboBox()
        self.award_filter.addItems(["All Awards", "First Year", "Bronze", "Silver", "Gold", "Owl", "None"])
        self.award_filter.currentTextChanged.connect(self.refresh_overview)
        self.status_filter = QComboBox()
        self.status_filter.addItems(["Active Students Only", "All Students", "Graduated Only"])
        self.status_filter.currentTextChanged.connect(self.refresh_overview)
        for label, widget in (("Search", self.search), ("Grade", self.grade_filter), ("Award", self.award_filter), ("Status", self.status_filter)):
            row.addWidget(QLabel(label))
            row.addWidget(widget)
        row.addWidget(self._button("Refresh", self.refresh_all))
        layout.addWidget(filters)
        batch = QGroupBox("Batch Award Overrides")
        batch_row = QHBoxLayout(batch)
        batch_row.addWidget(self._button("Select Visible", lambda: self._overview_select_all(True)))
        batch_row.addWidget(self._button("Clear Selection", lambda: self._overview_select_all(False)))
        self.overview_standard_batch = QComboBox()
        self.overview_standard_batch.addItems(["Keep/Calculated", "Reset Override", *AWARD_ORDER])
        self.overview_owl_batch = QComboBox()
        self.overview_owl_batch.addItems(["Keep Owl", "Award Owl", "Remove Owl"])
        batch_row.addWidget(QLabel("Standard"))
        batch_row.addWidget(self.overview_standard_batch)
        batch_row.addWidget(QLabel("Owl"))
        batch_row.addWidget(self.overview_owl_batch)
        batch_row.addWidget(self._button("Apply to Selected", self.apply_overview_batch))
        layout.addWidget(batch)
        self.student_table = self._table(["Select", "Email", "Student ID", "Full Name", "Grade", "4-Year Hours", "Award", "Standard Override", "Owl Override", "Counselor Write-up"])
        self.student_table.itemChanged.connect(self._overview_selection_changed)
        self.student_table.cellDoubleClicked.connect(self.show_student_detail)
        layout.addWidget(self.student_table)
        return tab

    def _history_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        top = QHBoxLayout()
        top.addWidget(QLabel("Every recorded pin award, by student and academic year"))
        top.addStretch()
        top.addWidget(self._button("Export Pin History CSV", self.export_history))
        layout.addLayout(top)
        self.history_table = self._table(["Email", "Student ID", "Full Name", "Grade", "Academic Year", "Pin Awarded", "Year Hours", "Owl Write-up", "Requested"])
        layout.addWidget(self.history_table)
        return tab

    def _counselor_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        top = QGridLayout()
        top.addWidget(QLabel("Generate counselor requests or send them directly with the Email Center composer."), 0, 0, 1, 3)
        top.addWidget(self._button("Export Counselor Autocrat CSV", self.export_counselor), 1, 0)
        top.addWidget(self._button("Send Counselor Emails", self.compose_counselor), 1, 1)
        top.addWidget(self._button("Import Google Form CSV", self.import_counselor), 1, 2)
        layout.addLayout(top)
        self.counselor_table = self._table(["Student", "Student ID", "Grade", "Counselor Email", "Write-up Status"])
        layout.addWidget(self.counselor_table)
        return tab

    def _stats_tab(self):
        tab = QWidget()
        tab.setStyleSheet(
            "QPushButton { padding: 3px 6px; min-height: 20px; font-size: 10px; }"
        )
        layout = QVBoxLayout(tab)
        actions_section = CollapsibleSection("Award controls", expanded=True)
        actions = QGridLayout(actions_section.content)
        self.percent = QSpinBox()
        self.percent.setRange(0, 100)
        self.percent.setValue(int(self.rubric_percentage))
        actions.addWidget(QLabel("Award Owl to top"), 0, 0)
        percentage_controls = QWidget()
        percentage_layout = QHBoxLayout(percentage_controls)
        percentage_layout.setContentsMargins(0, 0, 0, 0)
        percentage_layout.addWidget(self.percent)
        percentage_layout.addWidget(QLabel("% of eligible Gold students"))
        actions.addWidget(percentage_controls, 0, 1)
        simulation_buttons = QWidget()
        simulation_layout = QHBoxLayout(simulation_buttons)
        simulation_layout.setContentsMargins(0, 0, 0, 0)
        simulation_layout.addWidget(self._button("Simulate, reserve overrides", self.simulate_owl_reserving_overrides))
        simulation_layout.addWidget(self._button("Simulate full slots", self.simulate_owl_full_slots))
        actions.addWidget(simulation_buttons, 0, 2)
        actions.addWidget(self._button("Reset All Award Overrides", self.reset_all_award_overrides), 1, 0, 1, 3)
        layout.addWidget(actions_section)
        filters = CollapsibleSection("Awarded This Year Filters", expanded=True)
        filter_grid = QGridLayout(filters.content)
        self.stats_search = QLineEdit()
        self.stats_search.setPlaceholderText("Search name, student ID, or email...")
        self.stats_grade_filter = QComboBox()
        self.stats_grade_filter.addItems(["All Grades", "9", "10", "11", "12"])
        self.stats_award_filter = QComboBox()
        self.stats_award_filter.addItems(["Awarded This Year", "All Awards", *AWARD_ORDER, "None"])
        self.stats_status_filter = QComboBox()
        self.stats_status_filter.addItems(["Active Students Only", "All Students", "Graduated Only"])
        for widget in (self.stats_grade_filter, self.stats_award_filter, self.stats_status_filter):
            widget.currentTextChanged.connect(self.refresh_stats)
        self.stats_search.textChanged.connect(self.refresh_stats)
        filter_grid.addWidget(QLabel("Search"), 0, 0)
        filter_grid.addWidget(self.stats_search, 0, 1)
        filter_grid.addWidget(QLabel("Grade"), 0, 2)
        filter_grid.addWidget(self.stats_grade_filter, 0, 3)
        filter_grid.addWidget(QLabel("Award"), 1, 0)
        filter_grid.addWidget(self.stats_award_filter, 1, 1)
        filter_grid.addWidget(QLabel("Status"), 1, 2)
        filter_grid.addWidget(self.stats_status_filter, 1, 3)
        layout.addWidget(filters)
        self.stats_label = QLabel()
        layout.addWidget(self.stats_label)
        self.rubric_table = self._table(["Score", "Student ID", "Student Name", "Grade", "Total Hours", "Award This Year", "Keep Current", *AWARD_ORDER, "Owl"])
        self.rubric_table.itemChanged.connect(self._rubric_selection_changed)
        self.rubric_table.cellClicked.connect(self._stats_cell_clicked)
        layout.addWidget(self.rubric_table)
        message_group = CollapsibleSection("Award Message", expanded=True)
        message_layout = QVBoxLayout(message_group.content)
        self.message_editor = QPlainTextEdit()
        self.message_editor.setPlainText(self._award_message())
        message_layout.addWidget(self.message_editor)
        save_message = self._button("Save Award Message", self.save_message)
        message_layout.addWidget(save_message)
        layout.addWidget(message_group)
        return tab

    def _email_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        settings_group = QGroupBox("SMTP account")
        form = QFormLayout(settings_group)
        self.smtp_server = QLineEdit(SMTP_SERVER)
        self.smtp_port = QSpinBox()
        self.smtp_port.setRange(1, 65535)
        self.smtp_port.setValue(SMTP_PORT)
        self.smtp_username = QLineEdit(SMTP_USERNAME)
        self.smtp_password = QLineEdit(SMTP_PASSWORD)
        self.smtp_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.smtp_from = QLineEdit(SENDER_EMAIL)
        self.smtp_timeout = QSpinBox()
        self.smtp_timeout.setRange(1, 300)
        self.smtp_timeout.setValue(SMTP_TIMEOUT)
        self.file_producer = QLineEdit(FILE_PRODUCER)
        self.file_publisher = QLineEdit(FILE_PUBLISHER)
        self.smtp_tls = QCheckBox("Use STARTTLS")
        self.smtp_tls.setChecked(SMTP_USE_TLS)
        self.smtp_ssl = QCheckBox("Use SSL")
        self.smtp_ssl.setChecked(SMTP_USE_SSL)
        form.addRow("SMTP server", self.smtp_server)
        form.addRow("SMTP port", self.smtp_port)
        form.addRow("Username", self.smtp_username)
        form.addRow("Password", self.smtp_password)
        form.addRow("From address", self.smtp_from)
        form.addRow("Timeout (seconds)", self.smtp_timeout)
        form.addRow("File producer", self.file_producer)
        form.addRow("File publisher", self.file_publisher)
        security = QHBoxLayout()
        security.addWidget(self.smtp_tls)
        security.addWidget(self.smtp_ssl)
        security.addStretch()
        form.addRow("Security", security)
        settings_buttons = QHBoxLayout()
        settings_buttons.addWidget(self._button("Test SMTP Connection", self.test_smtp))
        settings_buttons.addWidget(QLabel("Values default from CSA_SMTP_* environment variables and apply to new composers."))
        settings_buttons.addStretch()
        layout.addWidget(settings_group)
        layout.addLayout(settings_buttons)
        actions = QGroupBox("Autocrat replacements")
        action_layout = QGridLayout(actions)
        action_layout.addWidget(QLabel("Open a composer, choose the recipient column, edit the subject/body, preview a row, then send all."), 0, 0, 1, 3)
        action_layout.addWidget(self._button("Send Counselor Emails", self.compose_counselor), 1, 0)
        action_layout.addWidget(self._button("Send Student Award Emails", self.compose_student_awards), 1, 1)
        action_layout.addWidget(self._button("Send Senior Award Emails", self.compose_senior_awards), 1, 2)
        action_layout.addWidget(self._button("Counselor Word + Excel", self.export_counselor_mail_merge), 2, 0)
        action_layout.addWidget(self._button("Student Word + Excel", self.export_student_mail_merge), 2, 1)
        action_layout.addWidget(self._button("Senior Word + Excel", self.export_senior_mail_merge), 2, 2)
        layout.addWidget(actions)
        layout.addStretch()
        return tab

    def _table(self, headers):
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setWordWrap(False)
        table.verticalHeader().setDefaultSectionSize(28)
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setSectionResizeMode(0, table.horizontalHeader().ResizeMode.Stretch)
        return table

    def refresh_all(self):
        self.summaries = self._load_summaries()
        self.refresh_overview()
        self.refresh_history()
        self.refresh_counselors()
        self.refresh_stats()

    def _load_summaries(self):
        session = SessionLocal()
        try:
            rows = calculate_awards(session)
            session.commit()
            return rows
        finally:
            session.close()

    @staticmethod
    def _writeup(student):
        return any(record.owl_writeup for record in student.pin_history if record.pin_awarded == "Owl")

    def _matches(self, student, award):
        query = self.search.text().lower().strip()
        if query and query not in f"{student.email} {student.student_id or ''} {student.full_name}".lower():
            return False
        if self.grade_filter.currentText() != "All Grades" and str(student.grade_level) != self.grade_filter.currentText():
            return False
        if self.award_filter.currentText() != "All Awards" and award != self.award_filter.currentText():
            return False
        if self.status_filter.currentText() == "Active Students Only" and student.status != "ACTIVE":
            return False
        if self.status_filter.currentText() == "Graduated Only" and student.status != "GRADUATED":
            return False
        return True

    def refresh_overview(self):
        if not hasattr(self, "student_table"):
            return
        self.student_table.setRowCount(0)
        for summary in self.summaries:
            student = summary["student"]
            award = summary["current_pin"] or "None"
            if not self._matches(student, award):
                continue
            row = self.student_table.rowCount()
            self.student_table.insertRow(row)
            values = ["", student.email, student.student_id or "", student.full_name, student.grade_level or "",
                      f"{summary['total_hours']:.1f}", award,
                      summary.get("admin_award_override") or "Calculated",
                      "Yes" if summary.get("admin_owl_override") else "No",
                      "Received" if self._writeup(student) else ("Pending" if award == "Owl" else "-")]
            for col, value in enumerate(values):
                self.student_table.setItem(row, col, QTableWidgetItem(str(value)))
            select = self.student_table.item(row, 0)
            select.setFlags(select.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            select.setCheckState(Qt.CheckState.Unchecked)
            if award == "Owl" and not self._writeup(student):
                for col in range(self.student_table.columnCount()):
                    self.student_table.item(row, col).setBackground(QColor("#fff2bf"))
        active = sum(summary["student"].status == "ACTIVE" for summary in self.summaries)
        self.status.setText(f"SQLite Connected | Active Students: {active} | Displayed: {self.student_table.rowCount()}")

    def refresh_history(self):
        if not hasattr(self, "history_table"):
            return
        self.history_table.setRowCount(0)
        for student in self._students():
            hours_by_year = defaultdict(float)
            for submission in student.submissions:
                hours_by_year[submission.school_year] += submission.hours_submitted
            for pin in sorted(student.pin_history, key=lambda record: (record.academic_year, record.id)):
                values = [student.email, student.student_id or "", student.full_name, student.grade_level or "",
                          pin.academic_year, pin.pin_awarded, f"{hours_by_year[pin.academic_year]:.1f}",
                          "Received" if pin.owl_writeup else ("Pending" if pin.pin_awarded == "Owl" else "-"),
                          "Yes" if pin.owl_requested else "No"]
                self._add_row(self.history_table, values)

    def refresh_counselors(self):
        if not hasattr(self, "counselor_table"):
            return
        self.counselor_table.setRowCount(0)
        for summary in self._owl_candidates():
            student = summary["student"]
            self._add_row(self.counselor_table, [student.full_name, student.student_id or "", student.grade_level or "",
                                                 student.counselor_email or "", "Received" if self._writeup(student) else "Pending"])

    def _add_row(self, table, values):
        row = table.rowCount()
        table.insertRow(row)
        for col, value in enumerate(values):
            table.setItem(row, col, QTableWidgetItem(str(value)))

    def _students(self):
        session = SessionLocal()
        try:
            return (session.query(Student)
                    .options(selectinload(Student.submissions), selectinload(Student.pin_history))
                    .order_by(Student.full_name, Student.email).all())
        finally:
            session.close()

    def _owl_candidates(self):
        return [summary for summary in self.summaries if summary["current_pin"] == "Owl" and summary["student"].status == "ACTIVE"]

    def show_student_detail(self, row, _column):
        email = self.student_table.item(row, 1).text()
        summary = next((item for item in self.summaries if item["student"].email == email), None)
        if summary is None:
            return
        student = summary["student"]
        dialog = QDialog(self)
        dialog.setWindowTitle("Student Overview")
        dialog.setMinimumWidth(560)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(f"<b>{student.full_name}</b> | {student.email}"))
        layout.addWidget(QLabel(f"ID: {student.student_id or '-'} | Grade: {student.grade_level or '-'} | Status: {student.status}"))
        layout.addWidget(QLabel(f"Cumulative hours: {summary['total_hours']:.1f} | Calculated/current award: {summary['award_label'] or 'None'}"))
        layout.addWidget(QLabel("Award override"))
        standard = QComboBox()
        standard.addItem("Calculated/Reset", "")
        for award in AWARD_ORDER:
            standard.addItem(award, award)
        current_standard = summary.get("admin_award_override") or ""
        if current_standard:
            standard.setCurrentIndex(max(0, standard.findData(current_standard)))
        owl = QCheckBox("Award Owl pin")
        owl.setChecked(bool(summary.get("admin_owl_override")))
        layout.addWidget(standard)
        layout.addWidget(owl)

        def save_override():
            session = SessionLocal()
            try:
                stored = session.query(Student).filter_by(email=email).one_or_none()
                if stored is None:
                    return
                chosen_standard = standard.currentData()
                chosen_owl = owl.isChecked()
                override = session.query(AwardOverride).filter_by(student_id=stored.id).first()
                if not chosen_standard and not chosen_owl:
                    if override:
                        session.delete(override)
                    stored.admin_award_override = None
                    stored.owl_override = None
                    stored.owl_selected = False
                else:
                    if override is None:
                        override = AwardOverride(student_id=stored.id, award=chosen_standard or "Owl")
                        session.add(override)
                    override.standard_award = chosen_standard or None
                    override.owl_awarded = chosen_owl
                    override.award = chosen_standard or "Owl"
                    stored.admin_award_override = None
                    stored.owl_override = True if chosen_owl else None
                    stored.owl_selected = chosen_owl
                session.commit()
            finally:
                session.close()
            dialog.accept()
            self.refresh_all()

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Close)
        buttons.accepted.connect(save_override)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec()

    def _owl_rows(self, grades=None):
        eligible = [summary for summary in self.summaries if summary.get("hour_pin") == "Gold" and summary["student"].status == "ACTIVE"]
        if grades:
            eligible = [summary for summary in eligible if str(summary["student"].grade_level) in grades]
        eligible.sort(key=lambda summary: (summary["student"].owl_rubric_score is None,
                                           -(summary["student"].owl_rubric_score or 0), summary["student"].full_name.lower()))
        percentage = self.percent.value()
        target = math.ceil(len(eligible) * percentage / 100) if eligible else 0
        rows = []
        for index, summary in enumerate(eligible):
            score = summary["student"].owl_rubric_score
            rows.append((summary, score, index + 1, (index + 1) * 100 / len(eligible) if eligible else 0,
                         score is not None and index < target))
        return rows, target, percentage

    def _stats_award(self, summary):
        current_year = self._current_academic_year()
        current = next((year for year in summary["yearly"] if year["academic_year"] == current_year), None)
        if current is None:
            return ""
        return summary.get("award_label") or (summary["current_pin"] if summary["current_pin"] == "Owl" else current["pin"])

    def _current_academic_year(self):
        years = [year["academic_year"] for summary in self.summaries for year in summary["yearly"]]
        return max(years) if years else ""

    def _stats_matches(self, summary):
        student = summary["student"]
        award = self._stats_award(summary)
        query = self.stats_search.text().lower().strip()
        if query and query not in f"{student.email} {student.student_id or ''} {student.full_name}".lower():
            return False
        if self.stats_grade_filter.currentText() != "All Grades" and str(student.grade_level) != self.stats_grade_filter.currentText():
            return False
        if self.stats_status_filter.currentText() == "Active Students Only" and student.status != "ACTIVE":
            return False
        if self.stats_status_filter.currentText() == "Graduated Only" and student.status != "GRADUATED":
            return False
        award_filter = self.stats_award_filter.currentText()
        if award_filter == "Awarded This Year" and not award:
            return False
        if (award_filter not in {"Awarded This Year", "All Awards"}
            and award_filter not in award.split(" + ")):
            return False
        return True

    def refresh_stats(self):
        if not hasattr(self, "rubric_table"):
            return
        self.summaries = self._load_summaries()
        counts = Counter(self._stats_award(summary) or "None" for summary in self.summaries)
        rows, target, percentage = self._owl_rows()
        visible = [summary for summary in self.summaries if self._stats_matches(summary)]
        self.stats_label.setText(" | ".join(f"{award}: {counts.get(award, 0)}" for award in ("Owl", "Gold", "Silver", "Bronze", "First Year", "None"))
                                   + f"    | Eligible Gold: {len(rows)} | Owl slots at {percentage}%: {target} | Showing awards: {len(visible)}")
        self.rubric_table.blockSignals(True)
        self.rubric_table.setRowCount(0)
        for summary in visible:
            student = summary["student"]
            award = self._stats_award(summary)
            has_override = summary.get("admin_award_override") or summary.get("admin_owl_override")
            display_name = f"{student.full_name} *" if has_override else student.full_name
            values = ["No data" if student.owl_rubric_score is None else student.owl_rubric_score,
                      student.student_id or "", display_name, student.grade_level or "",
                      f"{summary['total_hours']:.1f}", award or "None", ""] + ["" for _ in AWARD_ORDER] + [""]
            self._add_row(self.rubric_table, values)
            row = self.rubric_table.rowCount() - 1
            keep = self.rubric_table.item(row, 6)
            keep.setFlags(keep.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            keep.setCheckState(Qt.CheckState.Checked if not has_override else Qt.CheckState.Unchecked)
            for offset, option in enumerate((*AWARD_ORDER, "Owl"), start=7):
                cell = self.rubric_table.item(row, offset)
                flags = cell.flags() | Qt.ItemFlag.ItemIsUserCheckable
                cell.setFlags(flags)
                selected = (summary.get("admin_award_override") == option
                            if option in AWARD_ORDER else summary.get("admin_owl_override", False))
                cell.setCheckState(Qt.CheckState.Checked if selected else Qt.CheckState.Unchecked)
        self.rubric_table.blockSignals(False)

    def _rubric_selection_changed(self, item):
        if item.column() < 6:
            return
        selected = item.checkState() == Qt.CheckState.Checked
        selected_rows = {index.row() for index in self.rubric_table.selectionModel().selectedRows()}
        selected_rows.add(item.row())
        student_ids = [self.rubric_table.item(row, 1).text() for row in selected_rows]
        standard_award = None
        owl_awarded = None
        if 7 <= item.column() <= 10 and selected:
            standard_award = AWARD_ORDER[item.column() - 7]
        elif item.column() == 11:
            owl_awarded = selected
        session = SessionLocal()
        try:
            students = session.query(Student).filter(Student.student_id.in_(student_ids)).all()
            for student in students:
                stored_override = (session.query(AwardOverride)
                                   .filter_by(student_id=student.id)
                                   .order_by(AwardOverride.id)
                                   .first())
                if item.column() == 6:
                    if stored_override is not None:
                        session.delete(stored_override)
                    student.admin_award_override = None
                    student.owl_override = None
                    student.owl_selected = False
                    continue
                existing_standard = stored_override.standard_award if stored_override else None
                if stored_override and not existing_standard and stored_override.award in AWARD_ORDER:
                    existing_standard = stored_override.award
                existing_owl = bool(stored_override and (stored_override.owl_awarded or stored_override.award == "Owl"))
                chosen_standard = standard_award if standard_award is not None else existing_standard
                chosen_owl = owl_awarded if owl_awarded is not None else existing_owl
                if chosen_standard is None and not chosen_owl:
                    if stored_override is not None:
                        session.delete(stored_override)
                else:
                    if stored_override is None:
                        stored_override = AwardOverride(student_id=student.id, award=chosen_standard or "Owl")
                        session.add(stored_override)
                    stored_override.standard_award = chosen_standard or None
                    stored_override.owl_awarded = chosen_owl
                    stored_override.award = chosen_standard or "Owl"
                student.admin_award_override = None
                student.owl_override = chosen_owl if owl_awarded is not None else None
                if owl_awarded is not None:
                    student.owl_selected = chosen_owl
                session.commit()
        finally:
            session.close()
        self.summaries = self._load_summaries()
        self.refresh_overview()
        self.refresh_counselors()
        self.refresh_stats()

    def _stats_cell_clicked(self, row, column):
        if column != 2:
            return
        student_id = self.rubric_table.item(row, 1).text()
        session = SessionLocal()
        try:
            student = session.query(Student).filter_by(student_id=student_id).one_or_none()
            if student is None:
                return
            override = (session.query(AwardOverride)
                        .filter_by(student_id=student.id)
                        .order_by(AwardOverride.id)
                        .first())
            if override is None:
                return
            answer = QMessageBox.question(
                self, "Reset award override",
                f"Reset {student.full_name} to the calculated award?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
            session.delete(override)
            student.admin_award_override = None
            student.owl_override = None
            session.commit()
        finally:
            session.close()
        self.refresh_all()

    def reset_all_award_overrides(self):
        answer = QMessageBox.question(
            self, "Reset all award overrides",
            "Reset every administrator award override to the calculated award?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        session = SessionLocal()
        try:
            session.query(AwardOverride).delete(synchronize_session=False)
            session.query(Student).update({Student.admin_award_override: None, Student.owl_override: None,
                                           Student.owl_selected: False},
                                          synchronize_session=False)
            session.commit()
        finally:
            session.close()
        self.refresh_all()

    def _sync_rubric_defaults(self, reserve_overrides=True):
        eligible = [summary for summary in self.summaries
                    if summary.get("hour_pin") == "Gold" and summary["student"].status == "ACTIVE"]
        eligible.sort(key=lambda summary: (summary["student"].owl_rubric_score is None,
                                           -(summary["student"].owl_rubric_score or 0),
                                           summary["student"].full_name.lower()))
        target = math.ceil(len(eligible) * self.percent.value() / 100) if eligible else 0
        manual_owl_ids = {summary["student"].id for summary in eligible if summary.get("admin_owl_override")}
        available_slots = max(0, target - len(manual_owl_ids)) if reserve_overrides else target
        candidates = [summary for summary in eligible
                  if summary["student"].id not in manual_owl_ids
                  and summary["student"].owl_rubric_score is not None]
        rubric_ids = {summary["student"].id for summary in candidates[:available_slots]}
        session = SessionLocal()
        changed = False
        try:
            for student in session.query(Student).filter_by(status="ACTIVE").all():
                summary = next((item for item in self.summaries if item["student"].id == student.id), None)
                if not summary or not summary.get("admin_owl_override"):
                    rubric_selected = student.id in rubric_ids
                    if student.owl_selected != rubric_selected:
                        student.owl_selected = rubric_selected
                        changed = True
            if changed:
                session.commit()
        finally:
            session.close()

    def _run_owl_simulation(self, reserve_overrides):
        self.summaries = self._load_summaries()
        self._sync_rubric_defaults(reserve_overrides=reserve_overrides)
        self.refresh_all()

    def simulate_owl_reserving_overrides(self):
        self._run_owl_simulation(reserve_overrides=True)

    def simulate_owl_full_slots(self):
        self._run_owl_simulation(reserve_overrides=False)

    def _choose_file(self, caption, filters):
        path, _ = QFileDialog.getOpenFileName(self, caption, str(Path.home()), filters)
        return path

    def import_hours(self):
        path = self._choose_file("Import service hours", "CSV/XLSX (*.csv *.xlsx *.xlsm)")
        if path:
            self._run_import(import_form_responses, path, "service submissions")

    def import_roster(self):
        path = self._choose_file("Import student roster", "CSV/XLSX (*.csv *.xlsx *.xlsm)")
        if path:
            self._run_import(import_student_directory, path, "roster records")

    def import_counselor(self):
        path = self._choose_file("Import counselor Google Form responses", "CSV/XLSX (*.csv *.xlsx *.xlsm)")
        if path:
            self._run_import(import_counselor_form_responses, path, "counselor responses")

    def _run_import(self, function, path, label):
        try:
            count = function(path)
            QMessageBox.information(self, "Import Complete", f"Imported {count} {label}.")
            self.refresh_all()
        except Exception as error:
            QMessageBox.critical(self, "Import Failed", str(error))

    def rollover(self):
        path = self._choose_file("Select the new roster for school-year rollover", "CSV/XLSX (*.csv *.xlsx *.xlsm)")
        if not path:
            return
        try:
            result = rollover_students(path)
            QMessageBox.information(self, "Rollover Complete",
                                    f"Updated {result['updated']} students by one grade.\n"
                                    f"Graduated {result['graduated']} grade 12 students.\n"
                                    f"Added {result['added']} new students.\n\n"
                                    "Awards are recalculated against the current academic-year submissions.")
            self.refresh_all()
        except Exception as error:
            QMessageBox.critical(self, "Rollover Failed", str(error))

    def _grades(self, title):
        dialog = GradeDialog(self, title)
        return dialog.grades() if dialog.exec() else set()

    def _award_message(self):
        return Path(AWARD_MESSAGE_FILE).read_text(encoding="utf-8") if Path(AWARD_MESSAGE_FILE).exists() else ""

    def edit_message(self):
        self.tabs.setCurrentIndex(4)
        self.message_editor.setFocus()

    def save_message(self):
        path = Path(AWARD_MESSAGE_FILE)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(self.message_editor.toPlainText().rstrip() + "\n", encoding="utf-8")
        self.status.setText(f"Award message saved: {path}")

    def _smtp_settings(self):
        return {
            "server": self.smtp_server.text().strip(),
            "port": self.smtp_port.value(),
            "username": self.smtp_username.text().strip(),
            "password": self.smtp_password.text(),
            "from_email": self.smtp_from.text().strip(),
            "use_tls": self.smtp_tls.isChecked(),
            "use_ssl": self.smtp_ssl.isChecked(),
            "timeout": self.smtp_timeout.value(),
        }

    @staticmethod
    def _add_merge_field(paragraph, field_name):
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instruction = OxmlElement("w:instrText")
        instruction.set(qn("xml:space"), "preserve")
        instruction.text = f" MERGEFIELD {field_name} "
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        placeholder = OxmlElement("w:t")
        placeholder.text = f"«{field_name}»"
        separate_run = OxmlElement("w:r")
        separate_run.append(placeholder)
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run = paragraph.add_run()
        run._r.append(begin)
        run._r.append(instruction)
        run._r.append(separate)
        run._r.append(separate_run)
        run._r.append(end)

    @classmethod
    def _add_template_line(cls, document, line):
        paragraph = document.add_paragraph()
        cursor = 0
        for match in re.finditer(r"\{([^{}]+)\}", line):
            if match.start() > cursor:
                paragraph.add_run(line[cursor:match.start()])
            cls._add_merge_field(paragraph, match.group(1))
            cursor = match.end()
        if cursor < len(line):
            paragraph.add_run(line[cursor:])
        return paragraph

    def _write_mail_merge_package(self, name, rows, subject, body):
        if not rows:
            QMessageBox.information(self, "No recipients", "There are no matching recipients for this mail merge.")
            return
        default_folder = Path.home() / "Documents"
        if not default_folder.exists():
            default_folder = Path.home()
        selected_folder = QFileDialog.getExistingDirectory(
            self, "Choose folder for Word and Excel mail-merge files", str(default_folder)
        )
        if not selected_folder:
            return
        output = Path(selected_folder)
        workbook_path = output / f"csa_{name}.xlsx"
        document_path = output / f"csa_{name}.docx"
        fields = list(rows[0])
        workbook = Workbook()
        workbook.properties.creator = self.file_producer.text().strip()
        workbook.properties.lastModifiedBy = self.file_publisher.text().strip()
        workbook.properties.title = f"CSA {name.replace('_', ' ').title()} Mail Merge"
        workbook.properties.subject = "Community Service Awards mail merge data"
        workbook.properties.description = f"Produced by {self.file_producer.text().strip()}; published by {self.file_publisher.text().strip()}"
        sheet = workbook.active
        sheet.title = "Mail Merge Data"
        sheet.append(fields)
        for row in rows:
            sheet.append([row.get(field, "") for field in fields])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            width = min(max(max(len(str(cell.value or "")) for cell in column) + 2, 12), 55)
            sheet.column_dimensions[column[0].column_letter].width = width
        workbook.save(workbook_path)

        document = Document()
        document.core_properties.title = f"CSA {name.replace('_', ' ').title()} Mail Merge"
        document.core_properties.author = self.file_producer.text().strip()
        document.core_properties.last_modified_by = self.file_publisher.text().strip()
        document.core_properties.subject = "Community Service Awards mail merge template"
        document.core_properties.comments = f"Producer: {self.file_producer.text().strip()} | Publisher: {self.file_publisher.text().strip()}"
        document.add_heading("Community Service Awards", level=1)
        self._add_template_line(document, f"Email subject: {subject}")
        document.add_paragraph()
        for line in body.splitlines() or [""]:
            self._add_template_line(document, line)
        document.add_paragraph()
        document.add_paragraph("Mail merge recipient field: ")
        self._add_merge_field(document.paragraphs[-1], "recipient_email")
        document.save(document_path)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(workbook_path.resolve())))
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(document_path.resolve())))
        self.status.setText(f"Created Word and Excel mail merge files in {output}")
        dialog = QMessageBox(self)
        dialog.setWindowTitle("Mail merge files created")
        dialog.setTextFormat(Qt.TextFormat.RichText)
        dialog.setText(
            f"Created and opened:<br>{document_path}<br>{workbook_path}<br><br>"
            "In Word, choose <b>Mailings &gt; Start Mail Merge &gt; E-mail Messages</b>, "
            "select the Excel file, and use <b>recipient_email</b> as the To field.<br><br>"
            "<a href='https://www.youtube.com/watch?v=rGG-In7PAEk'>"
            "Watch the Word/Excel mail merge tutorial</a>"
        )
        dialog.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        dialog.setStandardButtons(QMessageBox.StandardButton.Close)
        dialog.exec()

    def export_counselor_mail_merge(self):
        rows = self._mail_merge_counselor_rows()
        self._write_mail_merge_package(
            "counselors", rows, "CSA Owl Pin write-up for {student_name}",
            "Dear Counselor,\n\n{student_name} has earned an Owl Pin. Please submit a one-paragraph write-up here:\n{google_form_url}\n\nThank you,\nCSA Team",
        )

    def export_student_mail_merge(self):
        self._write_mail_merge_package(
            "students", self._student_email_rows(), "CSA award result for {student_name}",
            "Dear {student_name},\n\n{award_message}\n\nYou have qualified for the {award} Community Service Award with {total_hours} service hours.\n\nCSA Team",
        )

    def export_senior_mail_merge(self):
        self._write_mail_merge_package(
            "seniors", self._student_email_rows(seniors_only=True), "Senior Community Service Award for {student_name}",
            "Dear {student_name},\n\n{award_message}\n\nThis message recognizes your Community Service Award as a senior. Your recorded award is {award}, with {total_hours} service hours.\n\nCSA Team",
        )

    def _mail_merge_counselor_rows(self):
        rows = []
        for summary in self._owl_candidates():
            student = summary["student"]
            if not student.counselor_email:
                continue
            token = generate_counselor_token(student.id, student.counselor_email)
            rows.append({
                "recipient_email": student.counselor_email,
                "student_name": student.full_name,
                "student_id": student.student_id or "",
                "grade": student.grade_level or "",
                "service_pin": "Owl",
                "google_form_url": build_counselor_google_form_url(student, token),
                "jwt": token,
                "request_status": "Received" if self._writeup(student) else "Pending",
                "email_subject": f"CSA Owl Pin write-up for {student.full_name}",
            })
        return rows

    def test_smtp(self):
        if self.smtp_tls.isChecked() and self.smtp_ssl.isChecked():
            QMessageBox.warning(self, "SMTP security", "Choose STARTTLS or SSL, not both.")
            return
        try:
            test_smtp_connection(self._smtp_settings())
            QMessageBox.information(self, "SMTP connection", "SMTP connection and authentication succeeded.")
        except Exception as error:
            QMessageBox.critical(self, "SMTP connection failed", str(error))

    def _open_email_composer(self, title, rows, recipient_field, subject, body):
        if not rows:
            QMessageBox.information(self, "No recipients", "There are no matching recipients for this email preset.")
            return
        if self.smtp_tls.isChecked() and self.smtp_ssl.isChecked():
            QMessageBox.warning(self, "SMTP security", "Choose STARTTLS or SSL, not both.")
            return
        dialog = EmailComposerDialog(self, title, rows, recipient_field, subject, body, self._smtp_settings())
        dialog.exec()

    def compose_counselor(self):
        rows = self._mail_merge_counselor_rows()
        self._open_email_composer(
            "Counselor Owl Request Composer", rows, "recipient_email",
            "CSA Owl Pin write-up for {student_name}",
            "Dear Counselor,\n\n{student_name} has earned an Owl Pin. Please submit a one-paragraph write-up here:\n{google_form_url}\n\nThank you,\nCSA Team",
        )

    def _student_email_rows(self, seniors_only=False):
        rows = []
        for summary in self.summaries:
            student = summary["student"]
            if student.status != "ACTIVE" or not student.email or not summary["current_pin"]:
                continue
            if seniors_only and str(student.grade_level) != "12":
                continue
            rows.append({
                "recipient_email": student.email,
                "parent_email": student.parent_email or "",
                "student_email": student.email,
                "student_name": student.full_name,
                "student_id": student.student_id or "",
                "grade": student.grade_level or "",
                "service_pin": summary.get("award_label", summary["current_pin"]),
                "award": summary.get("award_label", summary["current_pin"]),
                "total_hours": f"{summary['total_hours']:.1f}",
                "award_message": self._award_message(),
                "counselor_email": student.counselor_email or "",
                "email_subject": f"CSA award result for {student.full_name}",
            })
        return rows

    def compose_student_awards(self):
        self._open_email_composer(
            "Student Award Email Composer", self._student_email_rows(), "recipient_email",
            "CSA award result for {student_name}",
            "Dear {student_name},\n\n{award_message}\n\nYou have qualified for the {award} Community Service Award with {total_hours} service hours.\n\nCSA Team",
        )

    def compose_senior_awards(self):
        self._open_email_composer(
            "Senior Award Email Composer", self._student_email_rows(seniors_only=True), "recipient_email",
            "Senior Community Service Award for {student_name}",
            "Dear {student_name},\n\n{award_message}\n\nThis message recognizes your Community Service Award as a senior. Your recorded award is {award}, with {total_hours} service hours.\n\nCSA Team",
        )

    def export_awards(self):
        grades = self._grades("Choose grades for student award Autocrat export")
        if grades:
            self._save_awards(grades)

    def _save_awards(self, grades):
        path, _ = QFileDialog.getSaveFileName(self, "Export student award Autocrat CSV", "csa_autocrat_awards.csv", "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["Recipient Email", "Student Name", "Student ID", "Grade", "Service Pin", "Total Hours", "Award Message", "Counselor Email"])
            for summary in self.summaries:
                student = summary["student"]
                if str(student.grade_level) in grades and summary["current_pin"]:
                    writer.writerow([student.email, student.full_name, student.student_id or "", student.grade_level or "",
                                     summary.get("award_label", summary["current_pin"]), summary["total_hours"], self._award_message(), student.counselor_email or ""])
        self.status.setText(f"Exported student Autocrat CSV: {path}")

    def export_counselor(self):
        grades = self._grades("Choose grades for counselor Autocrat export")
        if not grades:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export counselor Autocrat CSV", "csa_counselor_autocrat.csv", "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["Counselor Email", "Student ID", "Student Name", "Grade", "Service Pin", "Google Form URL", "JWT", "Request Status"])
            for summary in self._owl_candidates():
                student = summary["student"]
                if str(student.grade_level) not in grades or not student.counselor_email:
                    continue
                token = generate_counselor_token(student.id, student.counselor_email)
                writer.writerow([student.counselor_email, student.student_id or "", student.full_name, student.grade_level or "",
                                 "Owl", build_counselor_google_form_url(student, token), token,
                                 "Received" if self._writeup(student) else "Pending"])
        self.status.setText(f"Exported counselor Autocrat CSV: {path}")

    def export_rubric(self):
        grades = self._grades("Choose grades for Owl rubric Autocrat export")
        if not grades:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export Owl rubric Autocrat CSV", "csa_owl_rubric_autocrat.csv", "CSV (*.csv)")
        if not path:
            return
        rows, target, percentage = self._owl_rows(grades)
        with open(path, "w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["Rank", "Student ID", "Student Name", "Email", "Grade", "Owl Rubric Score", "Eligible Gold Students", "Top Percentage", "Owl Slots", "Group % of Eligible", "Top % of Eligible", "Group Gets Owl Pin"])
            for summary, score, rank, group_percent, gets_owl in rows:
                student = summary["student"]
                writer.writerow([rank, student.student_id or "", student.full_name, student.email, student.grade_level or "",
                                 "No data" if score is None else score, len(rows), percentage, target,
                                 f"{group_percent:.2f}%", f"{group_percent:.2f}%", "Yes" if gets_owl else "No"])
        self.status.setText(f"Exported Owl rubric Autocrat CSV: {path}")

    def export_history(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export pin history CSV", "csa_pin_history.csv", "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["Email", "Student ID", "Student Name", "Grade", "Academic Year", "Pin Awarded", "Year Hours", "Owl Write-up", "Requested"])
            for student in self._students():
                hours_by_year = defaultdict(float)
                for submission in student.submissions:
                    hours_by_year[submission.school_year] += submission.hours_submitted
                for pin in student.pin_history:
                    writer.writerow([student.email, student.student_id or "", student.full_name, student.grade_level or "",
                                     pin.academic_year, pin.pin_awarded, hours_by_year[pin.academic_year],
                                     "Received" if pin.owl_writeup else "Pending" if pin.pin_awarded == "Owl" else "-",
                                     "Yes" if pin.owl_requested else "No"])
        self.status.setText(f"Exported pin history CSV: {path}")

    def closeEvent(self, event):
        if self.portal.server:
            self.portal.stop()
        super().closeEvent(event)

    def _overview_select_all(self, selected):
        if not hasattr(self, "student_table"):
            return
        self.student_table.blockSignals(True)
        for row in range(self.student_table.rowCount()):
            self.student_table.item(row, 0).setCheckState(
                Qt.CheckState.Checked if selected else Qt.CheckState.Unchecked)
        self.student_table.blockSignals(False)

    def _overview_selection_changed(self, _item):
        return

    def apply_overview_batch(self):
        selected_ids = []
        for row in range(self.student_table.rowCount()):
            if self.student_table.item(row, 0).checkState() == Qt.CheckState.Checked:
                selected_ids.append(self.student_table.item(row, 2).text())
        if not selected_ids:
            QMessageBox.information(self, "Batch Override", "Select at least one student.")
            return
        standard_choice = self.overview_standard_batch.currentText()
        standard = "" if standard_choice in {"Keep/Calculated", "Reset Override"} else standard_choice
        owl_choice = self.overview_owl_batch.currentText()
        owl = None if owl_choice == "Keep Owl" else owl_choice == "Award Owl"
        session = SessionLocal()
        try:
            students = session.query(Student).filter(Student.student_id.in_(selected_ids)).all()
            for student in students:
                override = session.query(AwardOverride).filter_by(student_id=student.id).first()
                if standard_choice == "Reset Override":
                    if override:
                        session.delete(override)
                    student.admin_award_override = None
                    student.owl_override = None
                    student.owl_selected = False
                    continue
                existing_owl = bool(override and (override.owl_awarded or override.award == "Owl"))
                chosen_owl = existing_owl if owl is None else owl
                existing_standard = (override.standard_award if override else None)
                if override and existing_standard is None and override.award in AWARD_ORDER:
                    existing_standard = override.award
                chosen_standard = existing_standard if standard_choice == "Keep/Calculated" else standard
                if not chosen_standard and not chosen_owl:
                    if override:
                        session.delete(override)
                else:
                    if override is None:
                        override = AwardOverride(student_id=student.id, award=chosen_standard or "Owl")
                        session.add(override)
                    override.standard_award = chosen_standard or None
                    override.owl_awarded = chosen_owl
                    override.award = chosen_standard or "Owl"
                student.admin_award_override = None
                student.owl_override = True if chosen_owl else None
                student.owl_selected = chosen_owl
            session.commit()
        except Exception as error:
            session.rollback()
            QMessageBox.critical(self, "Batch Override", str(error))
            return
        finally:
            session.close()
        self.refresh_all()


def run():
    app = QApplication.instance() or QApplication([])
    window = CSADashboard()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(run())
