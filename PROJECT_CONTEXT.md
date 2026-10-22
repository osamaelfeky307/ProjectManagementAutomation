# Project Management Automation — Project Context

> **EYouth Coordination WhatsApp Messenger**  
> A desktop Python application for automating WhatsApp communication with **instructors** and **coordinators** at eYouth training centers.

---

## 📋 Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [File Map](#file-map)
- [Module Breakdown](#module-breakdown)
- [Data Flow](#data-flow)
- [Configuration & Templates](#configuration--templates)
- [Dependencies](#dependencies)
- [Key Concepts](#key-concepts)
- [Data Files](#data-files)

---

## Overview

Project Management Automation is a **dark-themed desktop GUI** (CustomTkinter) that lets eYouth coordinators:

1. **Load a session schedule** from an Excel file (`.xlsx`)
2. **Send templated WhatsApp messages** to instructors and coordinators via [pywhatkit](https://github.com/Ankit404butfound/PyWhatKit) (which drives WhatsApp Web in the browser)
3. **Schedule automated triggers** (e.g., day-before reminders, pre-session alerts, post-session thanks) that fire at computed offsets from each session's time
4. **Generate coordinator messages** in Arabic, summarising assigned slots, estimated pay, and confirmation details
5. **Export logs** to Excel (`.xlsx`) after a sending run

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                            main.py (App)                            │
│  ┌──────────────┐  ┌────────────────────┐  ┌──────────────────────┐ │
│  │ ScheduleMixin│  │  CoordinatorApp    │  │  Instructor UI       │ │
│  │(schedule_    │  │  (coordinator_ui)  │  │  (embedded in main)  │ │
│  │ manager.py)  │  └────────────────────┘  └──────────────────────┘ │
│  └──────────────┘           │                         │             │
│         │                   ▼                         ▼             │
│         │         ┌──────────────────┐     ┌───────────────────┐   │
│         │         │coordinator_logic │     │    engine.py      │   │
│         │         │coordinator_parser│     │  (MessageEngine)  │   │
│         │         └──────────────────┘     └───────────────────┘   │
└─────────────────────────────────────────────────────────────────────┘
                                     │
                    ┌────────────────▼────────────────┐
                    │         pywhatkit               │
                    │   (opens WhatsApp Web in        │
                    │    browser, types & sends)      │
                    └─────────────────────────────────┘
```

The `main.py` class (`App`) inherits from both `ctk.CTk` and `ScheduleMixin`.  
Coordinator workflow is handled by the separate `CoordinatorApp` window launched from `main.py`.

---

## File Map

```
ProjectManagementAutomation/
│
├── main.py                      # 🚀 Entry point — Instructor messaging App + TourGuide
├── app_config.py                # 🎨 Color palette, default templates, file paths
├── engine.py                    # 📨 MessageEngine: pywhatkit wrapper with batching & delays
├── schedule_manager.py          # ⏰ ScheduleMixin: automated trigger scheduling logic + UI
├── ui_components.py             # 🧩 Reusable widgets: StatusDot, StyledButton, SidebarButton
│
├── coordinator_ui.py            # 🗂️ CoordinatorApp window + choose_recipient_type dialog
├── coordinator_logic.py         # ⚙️ Business logic: message generation, conflict detection
├── coordinator_parser.py        # 📑 Excel/CSV parsing, column mapping, phone normalisation
│
├── generate_sample.py           # 🛠️ One-off script: creates sample_sessions.xlsx for testing
├── refactor.py                  # 🔧 Dev utility: code-refactor helper (not runtime)
│
├── templates.json               # 💬 Persisted instructor message templates
├── coordinator_templates.json   # 💬 Persisted coordinator message templates (Arabic)
├── triggers.json                # ⏱️ Persisted schedule trigger definitions
├── schedules.json               # 📆 Persisted scheduled session data
├── instructor_data.json         # 📇 Persisted instructor contact directory
│
├── sample_sessions.xlsx         # 📊 Sample instructor session data (for testing)
├── sample_sessins.xlsx          # 📊 Duplicate / older sample (typo in filename)
├── Instructor_Template.xlsx     # 📋 Empty template for users to fill in
├── Coordination_Log_2026-05-08_2139.xlsx  # 📝 Example coordination export log
│
├── requirements.txt             # 📦 Python package dependencies
├── main.spec                    # 🏗️ PyInstaller build spec
├── build/ dist/                 # 🏗️ PyInstaller output directories
└── __pycache__/                 # Python bytecode cache
```

---

## Module Breakdown

### `main.py` — Application Entry Point

| Symbol | Purpose |
|--------|---------|
| `App(ctk.CTk, ScheduleMixin)` | Root application window; owns sidebar navigation, session table, send controls |
| `TourGuide` | Step-by-step on-boarding overlay (pulsing highlight box) |
| `normalize_phone(phone_str)` | Converts various Egyptian phone formats → E.164 (`+201XXXXXXXXX`) |
| `STATUS_*` constants | Message-state machine: `Not Sent → Sending → Sent/Failed → Acknowledged → Confirmed` |
| `STATUS_COLORS` | Maps each status to a color from the palette |

**App Tabs / Sections:**
- **Dashboard** — quick stats, recent activity  
- **Sessions** — spreadsheet-like table of loaded sessions with per-row status dots  
- **Templates** — CRUD editor for message templates  
- **Schedule** — visual grid of scheduled sends (from `ScheduleMixin`)  
- **Settings** — delay, wait-time, sound, engine params  

---

### `engine.py` — `MessageEngine`

Core sending layer. Wraps `pywhatkit.sendwhatmsg_instantly()` with:

| Method | Description |
|--------|-------------|
| `send_message(phone, message)` | Single send; returns `(bool, detail_str)` |
| `send_batch(recipients, template, callbacks)` | Background thread; iterates list with random delays |
| `_confirm_send_and_close_tab()` | `pyautogui.press("enter")` then `Ctrl+W` to close the tab |
| `stop()` / `reset()` | Threading stop-event for cancellation |

**Template placeholders:** `{name}`, `{group_code}`, `{time}`  
**Delay:** Random `20–40 s` between messages (configurable) to avoid WhatsApp rate-limiting.

---

### `schedule_manager.py` — `ScheduleMixin`

Mixed into `App`. Provides a visual **Schedule Manager** dialog:

- Loads sessions from `schedules.json`
- Displays a **grid**: rows = session slots, columns = triggers
- Each cell shows a `StatusDot` and can be toggled for scheduling
- **Triggers** are time offsets (e.g., `-1440 min` = day before, `-15 min` = pre-session, `+120 min` = post-thanks)
- Background thread fires messages at the computed absolute time
- Trigger definitions stored in `triggers.json` and editable via `_open_trigger_settings()`

---

### `coordinator_ui.py` — `CoordinatorApp`

Separate `CTkToplevel` window for the **Coordinator workflow**:

| Feature | Description |
|---------|-------------|
| `choose_recipient_type()` | Modal dialog: "Instructors" or "Coordinators" |
| File loading | Browse for sessions `.xlsx`; optionally load a phone directory file |
| Column mapping | Auto-detects columns; lets user remap if needed |
| Preview tab | Shows generated Arabic messages before sending |
| Analysis tab | Slot counts, estimated pay per coordinator |
| Conflict detection | Flags coordinators assigned to overlapping slots |
| Send panel | Progress bar, per-recipient status dots, abort button |
| Export | Saves messages and analysis to Excel |

**Status states:** `Ready → Sending → Sent / Failed / Skipped`

---

### `coordinator_logic.py` — Business Logic

| Function | Description |
|----------|-------------|
| `generate_coordinator_messages(rows, ...)` | Builds one Arabic WhatsApp message per coordinator from session rows |
| `calculate_slot_analysis(rows, ...)` | Computes slot counts and estimated earnings |
| `detect_conflicts(rows)` | Finds coordinators with overlapping time slots |
| `preview_messages(recipients)` | Returns formatted preview strings |
| `send_messages(recipients, engine, callbacks)` | Iterates and sends; calls engine.send_message |
| `save_outputs(recipients, analysis, log)` | Writes output Excel files |
| `load_coordinator_templates()` / `save_coordinator_templates()` | JSON persistence for Arabic templates |

**Arabic weekday mapping:** Full and abbreviated English day names → Arabic (السبت, الأحد, …)

**Default closing message** (Arabic):  
> "برجاء التواجد قبل المعاد بـ 30 دقيقة…" (Please be there 30 minutes early…)

---

### `coordinator_parser.py` — Data Parsing

| Symbol | Description |
|--------|-------------|
| `CANONICAL_COLUMNS` | Dict mapping internal keys → sets of accepted Excel header aliases |
| `load_data(file, coord_file?)` | Reads `.xlsx`; detects session sheet; merges coordinator directory |
| `validate_columns(headers)` | Returns missing required columns |
| `clean_phone_number(phone)` | Strips formatting, normalises to E.164 |
| `value_to_text(v)` | Converts Excel cell value → cleaned string |
| `_person_key(name, phone)` | Deduplication key: `(normalised_name, normalised_phone)` |
| `load_coordinator_directory(file)` | Reads a separate coordinator phone-book file |

**Auto-detected columns include:** `group_code`, `center_code`, `room_code`, `area`, `city`, `center_name`, `day`, `time`, `main_coordinator`, `main_phone`, `backup_coordinator`, `backup_phone`, `no_students`

---

### `app_config.py` — Configuration

| Export | Description |
|--------|-------------|
| `COLORS` | Full dark-mode palette dict (bg, accents, status colours) |
| `DEFAULT_TEMPLATES` | 7 instructor message templates (English) with `{name}`, `{group_code}`, `{time}` placeholders |
| `TEMPLATES_FILE` | Path to `templates.json` |
| `DATA_FILE` | Path to `instructor_data.json` |
| `TRIGGERS_FILE` | Path to `triggers.json` |
| `load_templates()` / `save_templates()` | JSON persistence helpers |

---

### `ui_components.py` — Shared Widgets

| Widget | Description |
|--------|-------------|
| `StatusDot` | Colored circular indicator (14 px by default); `set_color(hex)` |
| `StyledButton` | App-styled button: blue → purple hover, rounded, bold |
| `SidebarButton` | Transparent sidebar nav button; `set_active(bool)` toggles highlight |

---

## Data Flow

### Instructor Messaging

```
User loads .xlsx
      │
      ▼
normalize_phone() for each row
      │
      ▼
App session table (status = Not Sent)
      │
  [Send All]
      │
      ▼
MessageEngine.send_batch()
  ├─ Template rendered: {name}, {group_code}, {time}
  ├─ pywhatkit.sendwhatmsg_instantly()
  ├─ pyautogui Enter + Ctrl+W
  └─ on_progress callback → updates row status dot
      │
      ▼
Optional: Export log to Excel
```

### Coordinator Messaging

```
User loads sessions .xlsx (+ optional coordinator directory)
      │
      ▼
coordinator_parser.load_data()
  ├─ Auto-detects columns via CANONICAL_COLUMNS
  └─ Merges phone directory
      │
      ▼
coordinator_logic.generate_coordinator_messages()
  ├─ Groups rows by (coordinator name + phone)
  ├─ Renders Arabic template with section list
  └─ Returns recipients dict
      │
      ▼
CoordinatorApp preview / analysis tabs
      │
  [Send]
      │
      ▼
MessageEngine (same engine, same batch logic)
      │
      ▼
Save outputs: messages XLSX + analysis XLSX + log XLSX
```

---

## Configuration & Templates

### Instructor Templates (English)

Stored in `templates.json`. Defaults defined in `app_config.py`:

| # | Name | Trigger Timing |
|---|------|---------------|
| 1 | Day-Before Reminder | −24 h |
| 2 | Pre-Session Readiness | −15 min |
| 3 | Waiting Room Reminder | Session start |
| 4 | Attendance Reminder | Mid-session |
| 5 | Session Closing | End of session |
| 6 | Post-Session Thanks | +0 min after end |
| 7 | Session Time Exceeded | +N min after end |

### Coordinator Templates (Arabic)

Stored in `coordinator_templates.json`. Default:

| # | Name |
|---|------|
| 1 | Coordinator Confirmation (slot summary) |
| 2 | Zoom Reminder |

### Schedule Triggers

Stored in `triggers.json`. Each entry:
```json
{
  "key": "day_before",
  "offset_minutes": -1440,
  "template": "1. Day Before",
  "label": "📅 Day Before"
}
```

---

## Dependencies

```
customtkinter>=5.2.0   # Modern dark-mode Tkinter wrapper
pywhatkit>=5.4         # WhatsApp Web automation
openpyxl>=3.1.0        # Excel read/write
Pillow>=10.0.0         # Image support for CTk
pyautogui              # (runtime import) — press Enter, close tab
```

Install with:
```bash
pip install -r requirements.txt
pip install pyautogui   # not in requirements.txt but used at runtime
```

---

## Key Concepts

### Phone Normalisation

All phone numbers are normalised to **E.164** format before sending:

| Input | Output |
|-------|--------|
| `01012345678` | `+201012345678` |
| `1012345678` | `+201012345678` |
| `00201012345678` | `+201012345678` |
| `+201012345678` | `+201012345678` |

### Random Send Delays

Between every message send, the engine waits a **random 20–40 seconds** (defaults) to simulate human pacing and reduce the risk of WhatsApp flagging the account.

### Column Auto-Detection

`coordinator_parser.CANONICAL_COLUMNS` maps normalised (lowercase, no spaces) header variants to internal keys, enabling the app to work with many differently-named Excel column headers without user configuration.

### Conflict Detection

A coordinator is flagged as **conflicted** if they appear in two or more session rows that share the same **day** and **overlapping time slot**.

---

*Generated: 2026-07-28 | Project Management Automation v2*
