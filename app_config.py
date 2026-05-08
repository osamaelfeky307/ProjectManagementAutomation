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
        "Hello Eng. {name},\n\n"
        "This is a gentle reminder that you have a session for group {group_code} scheduled for tomorrow at {time}.\n"
        "Please ensure you are prepared.\n\n"
        "Thank you,\nEYouth coordination"
    ),
    "2. Pre-Session Readiness": (
        "Hello Eng. {name},\n\n"
        "Your session for group {group_code} at {time} is starting soon.\n"
        "Please confirm that you are ready to begin.\n\n"
        "Best regards,\nEYouth coordination"
    ),
    "3. Session Closing Reminder": (
        "Hello Eng. {name},\n\n"
        "Your session for group {group_code} is nearing its end.\n"
        "Please remember to close the session exactly on time—not before and not after.\n\n"
        "Thank you for your cooperation."
    ),
    "4. Post-Session Thank You": (
        "Dear Eng. {name},\n\n"
        "Thank you for completing today's {group_code} session. "
        "We appreciate your hard work and dedication.\n\n"
        "Best regards,\nEYouth coordination"
    ),
}

TEMPLATES_FILE = os.path.join(os.path.dirname(__file__), "templates.json")
DATA_FILE = os.path.join(os.path.dirname(__file__), "instructor_data.json")


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
