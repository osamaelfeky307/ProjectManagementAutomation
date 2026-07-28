"""
Configuration and constants for DECI Automation App.
"""
import json
import os

# --- Color Palette ---
COLORS = {
    "bg_dark": "#0f0f1a",
    "bg_card": "#1a1a2e",
    "bg_sidebar": "#12121f",
    "bg_input": "#252540",
    "accent_blue": "#4f6df5",
    "accent_purple": "#7c3aed",
    "accent_cyan": "#06b6d4",
    "accent_gradient_start": "#4f6df5",
    "accent_gradient_end": "#7c3aed",
    "text_primary": "#e8e8f0",
    "text_secondary": "#9ca3af",
    "text_muted": "#6b7280",
    "border": "#2a2a45",
    "status_red": "#ef4444",
    "status_yellow": "#f59e0b",
    "status_green": "#22c55e",
    "hover": "#2d2d4a",
    "danger": "#dc2626",
    "success": "#16a34a",
}

# --- Default Templates ---
DEFAULT_TEMPLATES = {
    "1. Day-Before Reminder": (
        "Hello Eng. *{name}*, 👋\n\n"
        "This is a reminder that you have a session for group *{group_code}* scheduled tomorrow at *{time}*. 📅\n"
        "Please make sure to join at least 15 minutes before the session starts to avoid assigning another trainer to the session.\n\n"
        "Thank you,\neyouth Coordination"
    ),
    "2. Pre-Session Readiness": (
        "Hello Eng. *{name}*, 🚨 \n\n"
        "Your session for group *{group_code}* at *{time}* will begin shortly.\n"
        "Please confirm that you are ready to start.\n\n"
        "• If you notice students are not focused or refusing to turn on their cameras, please inform me.\n"
        "• If there is any inappropriate behavior or disturbance from any student, please contact me immediately.\n"
        "• Please note that only 2 breaks are allowed during the session — no more and no less.\n"
        "• During breaks, please share a visible timer on screen.\n\n"
        "Thank you,\neyouth Coordination"
    ),
    "3. Waiting Room Reminder": (
        "Hello Eng. *{name}*, 👥🚨 \n\n"
        "Students for group *{group_code}* are currently waiting in the room.\n\n"
        "Please start admitting them immediately"
    ),
    "4. Attendance Reminder": (
        "Hello Eng. *{name}*, 📝\n\n"
        "Friendly reminder to take attendance for group *{group_code}* before ending the session."
    ),
    "5. Session Closing": (
        "Hello Eng. *{name}*, ⏰\n\n"
        "Your session for group *{group_code}* is nearing its scheduled end time.\n\n"
        "Please make sure to close the session exactly on time (not earlier and not later)"
    ),
    "6. Post-Session Thanks": (
        "Hello Eng. *{name}*, 🌟\n\n"
        "Thank you for completing today's session for group *{group_code}*.\n\n"
        "We appreciate your effort and commitment, and we hope the session went smoothly 👏🏼❤️"
    ),
    "7. Session Time Exceeded": (
        "Hello Eng. *{name}*, ⚠️\n\n"
        "Your session for group *{group_code}* has exceeded the scheduled end time.\n\n"
        "Please conclude the session and close the meeting as soon as possible "
    ),
}

TEMPLATES_FILE = os.path.join(os.path.dirname(__file__), "templates.json")
DATA_FILE = os.path.join(os.path.dirname(__file__), "instructor_data.json")
TRIGGERS_FILE = os.path.join(os.path.dirname(__file__), "triggers.json")


def load_templates():
    """Load templates from file or return defaults."""
    if os.path.exists(TEMPLATES_FILE):
        try:
            with open(TEMPLATES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return dict(DEFAULT_TEMPLATES)


def save_templates(templates: dict):
    """Persist templates to JSON file."""
    with open(TEMPLATES_FILE, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=2)
