# Community Service Awards Manager

Desktop application for managing a public schools community Service Awards database, including service-hour imports, award calculations, Owl Pin selection, counselor requests, email notifications, exports, and a LAN web portal.

## Features

- PyQt6 desktop dashboard
- SQLite database with automatic initialization
- Import student rosters and service-hour submissions from CSV/XLSX
- Import counselor Owl Pin form responses
- Calculate cumulative awards:
  - First Year: 10 hours
  - Bronze: 20 hours
  - Silver: 30 hours
  - Gold: 40 hours
- Track pin history by academic year
- Select Owl Pin recipients using rubric scores
- Apply individual or batch award overrides
- Send counselor and student emails through SMTP
- Export Autocrat-compatible CSV files
- Export Word/Excel mail-merge packages
- Optional PIN-protected LAN web portal

## Requirements

- Python 3.9 or newer
- macOS, Windows, or Linux
- Dependencies listed in `requirements.txt`

## Setup

### macOS

Run:

```bash
./run/"Setup CSA.command"
```

Or manually:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

### Windows

Double-click:

```text
run\Setup CSA.bat
```

Or manually:

```bat
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Launching

### macOS

```bash
./run/"Launch CSA.command"
```

### Windows

Double-click:

```text
run\Launch CSA.bat
```

### Linux or command line

```bash
.venv/bin/python app.py
```

Windows:

```bat
.venv\Scripts\python.exe app.py
```

## Main Workflow

1. Import the student roster from the **Data Imports** tab.
2. Import service-hour submissions.
3. Run **School-Year Rollover** when advancing to a new school year.
4. Review calculated awards in **Student Overview**.
5. Use **Award Stats** to review Gold students and simulate Owl selections.
6. Export counselor requests or send counselor emails.
7. Import completed counselor form responses.
8. Export student awards, counselor data, Owl rubric results, or pin history.
9. Use the **Email Center** for SMTP delivery or mail-merge exports.

## Data Files

Templates and sample files are located in `data/`.

Supported input formats:

- CSV
- XLSX
- XLSM

Expected columns include:

### Roster

```text
Email
Student ID
Full Name
Grade Level
Counselor Email
Graduation Year
Owl Rubric Score
```

### Service Hours

```text
Email
School Year
Hours Submitted
File Upload URL
```

### Counselor Responses

```text
Timestamp
Username
Student ID
Student Name
Owl Pin Blurb
JWT (DO NOT EDIT)
Confirm
```

## Command-Line Import

Import a roster and service-hour export without opening the GUI:

```bash
.venv/bin/python import_data.py roster.xlsx form_responses.csv
```

The command imports the records and recalculates awards.

## Configuration

Application settings are defined in `config.py`.

SMTP settings can be supplied through environment variables:

```bash
export CSA_SMTP_SERVER="smtp.example.com"
export CSA_SMTP_PORT="587"
export CSA_SMTP_USERNAME="username"
export CSA_SMTP_PASSWORD="password"
export CSA_SENDER_EMAIL="awards@example.com"
export CSA_SMTP_USE_TLS="true"
export CSA_SMTP_USE_SSL="false"
export CSA_SMTP_TIMEOUT="30"
```

Additional settings include:

- Award thresholds
- Owl selection percentage
- Counselor Google Form URL
- Dashboard dimensions
- Export metadata

Replace the default JWT secret in `config.py` before using the application outside a trusted local environment.

## Database

The application creates:

```text
csa_database.db
```

in the repository root.

The database stores:

- Students
- Service submissions
- Pin history
- Award overrides

Back up this file before making major changes. Deleting it removes all locally stored application data.

## LAN Web Portal

The desktop dashboard can start a shared portal for devices on the same local network.

From **Web Resources**:

1. Start the shared portal.
2. Share the displayed LAN URL and access PIN.
3. Users sign in with a display name and PIN.
4. Administrators can review students and apply award overrides.

The portal accepts LAN connections only and automatically expires sessions after 12 hours.

## Project Structure

```text
.
├── app.py                  # Desktop application entry point
├── pyqt_dashboard.py       # Main PyQt6 dashboard
├── web_server.py           # LAN web portal
├── database.py             # SQLAlchemy models and database setup
├── award_engine.py         # Award calculation logic
├── importer.py             # CSV/XLSX import functions
├── import_data.py          # Command-line import utility
├── email_dispatcher.py     # SMTP and counselor email helpers
├── google_contacts_sync.py # Google Contacts integration
├── config.py               # Application configuration
├── data/                   # Templates, samples, and message content
└── run/                    # Setup and launch scripts
```

## License

Copyright © 2026 Rayan Nasir. All rights reserved.
