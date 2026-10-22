"""
Configuration and constants for Project Management Automation App.
"""
import json
import os

# --- Color Palette ---
COLORS = {
    # Backgrounds
    "bg_dark": "#111113",         # Main app background (deepest)
    "bg_sidebar": "#161618",      # Sidebar background (slightly elevated)
    "bg_card": "#1c1c1f",         # Cards/panels background
    "bg_input": "#222225",        # Input fields
    "hover": "#2a2a2e",           # General hover state
    "border": "#28282c",          # Subtle borders

    # Text
    "text_primary": "#F2F2F2",    # High contrast primary text
    "text_secondary": "#A0A0A5",  # Muted secondary text
    "text_muted": "#707075",      # Disabled or very muted text

    # Brand Accents (Sleek Indigo/Purple)
    "accent_blue": "#5E6AD2",     # Primary brand color
    "accent_purple": "#7D7AFF",   # Accent / Hover for brand
    "accent_cyan": "#12A594",     # Muted Teal instead of bright cyan
    "accent_gradient_start": "#5E6AD2",
    "accent_gradient_end": "#7D7AFF",

    # Status Indicators (Refined, non-harsh)
    "status_red": "#E5484D",      # Linear-like red
    "status_yellow": "#FFB224",   # Soft warning yellow
    "status_green": "#30A46C",    # Soft success green
    "danger": "#E5484D",
    "success": "#30A46C",
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
