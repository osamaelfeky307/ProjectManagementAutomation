from __future__ import annotations
import datetime as dt
import json
import os
import random
import re
import time
from collections import defaultdict, Counter
import openpyxl
from engine import MessageEngine
from coordinator_parser import clean_phone_number, value_to_text, _person_key
import pandas as pd

DEFAULT_SLOT_PRICE = 250
DEFAULT_MAIN_LIMIT = 2
DEFAULT_BACKUP_LIMIT = 3
OUTPUT_MESSAGES_FILE = "coordinator_messages_output.xlsx"
OUTPUT_ANALYSIS_FILE = "coordinator_slot_analysis.xlsx"
OUTPUT_LOG_FILE = "whatsapp_sending_log.xlsx"
COORDINATOR_TEMPLATES_FILE = os.path.join(os.path.dirname(__file__), "coordinator_templates.json")
DEFAULT_COORDINATOR_CLOSING = (
    "برجاء التواجد قبل المعاد بـ 30 دقيقة للتأكد من إن المركز شغال وإن كل حاجة جاهزة قبل بداية السيشن.\n\n"
    "لو في أي مشكلة أو تعارض في المعاد، بلغنا فورًا.\n\n"
    "شكرًا ليك."
)
DEFAULT_COORDINATOR_TEMPLATES = {
    "1. Coordinator Confirmation": (
        "السلام عليكم يا {name}\n"
        "ده تأكيد توزيع السيشنز الخاصة بيك:\n\n"
        "{sections}\n\n"
        f"{DEFAULT_COORDINATOR_CLOSING}"
    ),
    "2. Zoom Reminder": (
        "متنساش تفتح ميتنج الزووم في ميعاده بالظبط. لو في أي مشكلة واجهتك، بلغني فوراً."
    )
}
ARABIC_WEEKDAYS = {
    "saturday": "السبت",
    "sat": "السبت",
    "sunday": "الأحد",
    "sun": "الأحد",
    "monday": "الإثنين",
    "mon": "الإثنين",
    "tuesday": "الثلاثاء",
    "tue": "الثلاثاء",
    "wednesday": "الأربعاء",
    "wed": "الأربعاء",
    "thursday": "الخميس",
    "thu": "الخميس",
    "friday": "الجمعة",
    "fri": "الجمعة",
}

def generate_coordinator_messages(rows, slot_price, main_limit, backup_limit, template_text=None):
    """Build one Arabic WhatsApp message per coordinator name + phone number."""
    recipients = {}
    skipped_logs = []

    for row in rows:
        for role, name_col, phone_col, gender_col in (
            ("Main", "main_coordinator", "main_phone", "main_gender"),
            ("Backup", "backup_coordinator", "backup_phone", "backup_gender"),
        ):
            name = row.get(name_col, "").strip()
            raw_phone = row.get(phone_col, "").strip()
            gender = row.get(gender_col, "").strip()
            # Skip placeholder names that aren't real coordinators
            if name.casefold() in {"بدون", "بدون منسق", "n/a", "na", "-", "--", "0", "none", ""}:
                continue
            if not name and not raw_phone:
                continue

            phone = clean_phone_number(raw_phone)
            if not name or not phone:
                skipped_logs.append(_log_row(name, raw_phone, "skipped", "Missing or invalid phone/name"))
                continue

            key = (name.casefold(), phone)
            recipients.setdefault(key, {"Name": name, "Phone": phone, "Gender": gender, "assignments": []})
            if gender and not recipients[key].get("Gender"):
                recipients[key]["Gender"] = gender
            recipients[key]["assignments"].append({"role": role, "row": row})

    messages = []
    for recipient in recipients.values():
        assignments = recipient["assignments"]
        original_main = [a for a in assignments if a["role"] == "Main"]
        original_backup = [a for a in assignments if a["role"] == "Backup"]
        role = "Both" if original_main and original_backup else "Main" if original_main else "Backup"

        # Backup coordinators receive the same message as their mains:
        # all assignments are treated as main assignments in the message body.
        all_as_main = list(assignments)
        for a in all_as_main:
            a = dict(a)  # don't mutate originals
        message = _build_coordinator_message(
            recipient["Name"],
            all_as_main,       # everything goes into the main section
            [],                # no separate backup section
            template_text,
            recipient.get("Gender", ""),
        )
        main_slots = len(_paid_slots(original_main))
        backup_slots = len(_paid_slots(original_backup))
        total_paid_slots = main_slots + backup_slots

        messages.append({
            "Name": recipient["Name"],
            "Phone": recipient["Phone"],
            "Gender": recipient.get("Gender", ""),
            "Role": role,
            "Main Slots": main_slots,
            "Backup Slots": backup_slots,
            "Total Paid Slots": total_paid_slots,
            "Amount": total_paid_slots * slot_price,
            "Message": message,
        })

    messages.sort(key=lambda item: item["Name"])
    return messages, skipped_logs

def calculate_slot_analysis(rows, slot_price, main_limit, backup_limit):
    """Calculate paid slots; a paid slot is role + day + time + center, not each row."""
    people = defaultdict(lambda: {"main_sessions": 0, "backup_sessions": 0, "main_slots": set(), "backup_slots": set()})
    conflicts = detect_conflicts(rows, main_limit, backup_limit)
    conflict_names = {c["Coordinator"].casefold() for c in conflicts}

    for row in rows:
        main = row.get("main_coordinator", "").strip()
        backup = row.get("backup_coordinator", "").strip()
        if main:
            people[main]["main_sessions"] += 1
            people[main]["main_slots"].add(_slot_key(row, "Main"))
        if backup:
            people[backup]["backup_sessions"] += 1
            people[backup]["backup_slots"].add(_slot_key(row, "Backup"))

    analysis = []
    for name, stats in sorted(people.items(), key=lambda item: item[0]):
        main_slots = len(stats["main_slots"])
        backup_slots = len(stats["backup_slots"])
        paid_slots = main_slots + backup_slots
        analysis.append({
            "Coordinator": name,
            "Main Sessions": stats["main_sessions"],
            "Main Slots": main_slots,
            "Backup Sessions": stats["backup_sessions"],
            "Backup Slots": backup_slots,
            "Paid Unique Slots": paid_slots,
            "Amount": paid_slots * slot_price,
            "Max Main in Same Slot": main_limit,
            "Max Backup in Same Slot": backup_limit,
            "Status": "Conflict" if name.casefold() in conflict_names else "OK",
        })
    return analysis

def detect_conflicts(rows, main_limit, backup_limit):
    """Find coordinators assigned to too many sessions in the same day/time/center slot."""
    buckets = defaultdict(list)
    for row in rows:
        for role, name_col in (("Main", "main_coordinator"), ("Backup", "backup_coordinator")):
            name = row.get(name_col, "").strip()
            if not name:
                continue
            buckets[(role, name.casefold(), row.get("day", ""), row.get("time", ""), _center_key(row))].append(row)

    conflicts = []
    for (role, _, day, slot_time, center), slot_rows in buckets.items():
        limit = main_limit if role == "Main" else backup_limit
        if len(slot_rows) <= limit:
            continue
        name = slot_rows[0]["main_coordinator"] if role == "Main" else slot_rows[0]["backup_coordinator"]
        conflicts.append({
            "Role": role,
            "Coordinator": name,
            "Day": day,
            "Time": slot_time,
            "Center": center,
            "Sessions": len(slot_rows),
            "Limit": limit,
        })
    return sorted(conflicts, key=lambda item: (item["Coordinator"], item["Day"], item["Time"]))

def preview_messages(messages):
    """Return a compact preview summary for confirmation before any send happens."""
    lines = [
        f"Number of recipients: {len(messages)}",
        "",
        "Recipients:",
    ]
    for msg in messages:
        lines.append(
            f"- {msg['Name']} | {msg['Phone']} | Role: {msg['Role']} | Slots: {msg['Total Paid Slots']}"
        )
    return "\n".join(lines)

def send_messages(messages, engine=None, min_delay=20, max_delay=40):
    """Send generated messages through the existing WhatsApp engine and log each result."""
    engine = engine or MessageEngine(min_delay=min_delay, max_delay=max_delay)
    logs = []
    for idx, msg in enumerate(messages):
        try:
            ok, detail = engine.send_message(msg["Phone"], msg["Message"])
            logs.append(_log_row(msg["Name"], msg["Phone"], "sent" if ok else "failed", "" if ok else detail))
        except Exception as exc:
            logs.append(_log_row(msg["Name"], msg["Phone"], "failed", str(exc)))

        if idx < len(messages) - 1:
            time.sleep(random.randint(min_delay, max_delay))
    return logs

def save_outputs(messages, analysis, log_rows, output_dir=None):
    """Write the coordinator message, slot analysis, and WhatsApp log workbooks."""
    output_dir = output_dir or os.getcwd()
    os.makedirs(output_dir, exist_ok=True)
    paths = {
        "messages": os.path.join(output_dir, OUTPUT_MESSAGES_FILE),
        "analysis": os.path.join(output_dir, OUTPUT_ANALYSIS_FILE),
        "log": os.path.join(output_dir, OUTPUT_LOG_FILE),
    }
    _write_workbook(paths["messages"], messages, [
        "Name", "Phone", "Role", "Main Slots", "Backup Slots", "Total Paid Slots", "Amount", "Message"
    ])
    _write_workbook(paths["analysis"], analysis, [
        "Coordinator", "Main Sessions", "Main Slots", "Backup Sessions", "Backup Slots", "Paid Unique Slots",
        "Amount", "Max Main in Same Slot", "Max Backup in Same Slot", "Status"
    ])
    _write_workbook(paths["log"], log_rows, ["Timestamp", "Name", "Phone", "Status", "Error"])
    return paths

def _normalize_gender(gender):
    text = _normalize_gender_text(gender)
    female_tokens = {"female", "f", "woman", "girl", "lady", "miss", "mrs", "انثى", "انثي", "بنت", "سيده", "ست", "مدام"}
    male_tokens = {"male", "m", "man", "boy", "ذكر", "زكر", "ولد", "راجل", "استاذ"}
    if any(token in text.split() for token in female_tokens) or any(token in text for token in {"female", "انثى", "انثي"}):
        return "female"
    if any(token in text.split() for token in male_tokens) or any(token in text for token in {"male", "ذكر", "زكر"}):
        return "male"
    return ""

def _normalize_gender_text(value):
    text = value_to_text(value).strip().casefold()
    text = re.sub(r"[\u064b-\u065f\u0670ـ]", "", text)
    text = text.translate(str.maketrans({
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ى": "ي",
        "ة": "ه",
    }))
    text = re.sub(r"[^a-z\u0600-\u06ff]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def _gender_forms(gender):
    normalized = _normalize_gender(gender)
    if normalized == "female":
        return {
            "gender": "female",
            "bek": "بيكي",
            "leek": "ليكي",
            "recipient_suffix": "ي",
        }
    return {
        "gender": "male" if normalized == "male" else "",
        "bek": "بيك",
        "leek": "ليك",
        "recipient_suffix": "",
    }

def _apply_gender_text(text, gender):
    if _normalize_gender(gender) != "female":
        return text
    text = re.sub(r"بيك(?!ي)", "بيكي", text)
    text = re.sub(r"عليك(?![مين])", "عليكي", text)
    text = re.sub(r"(?<!ع)ليك(?!ي)", "ليكي", text)
    text = re.sub(r"معاك(?!ي)", "معاكي", text)
    text = re.sub(r"(?<![اأإآء-ي])بك(?![اأإآء-ي])", "بكي", text)
    text = re.sub(r"(?<![اأإآء-ي])لك(?![اأإآء-ي])", "لكي", text)
    
    # Zoom Reminder specific replacements
    text = re.sub(r"متنساش", "متنسيش", text)
    text = re.sub(r"تفتح ", "تفتحي ", text)
    text = re.sub(r"بلغني", "بلغيني", text)
    text = re.sub(r"كلمني", "كلميني", text)
    return text

def _template_for_display(template_text):
    return (template_text or "").replace("{closing}", DEFAULT_COORDINATOR_CLOSING)

def _arabic_day(day):
    text = value_to_text(day)
    if not text:
        return ""

    def replace_part(match):
        part = match.group(0)
        return ARABIC_WEEKDAYS.get(part.casefold(), part)

    return re.sub(r"[A-Za-z]+", replace_part, text)

from template_utils import render_template as utils_render_template

def _render_template(template_text, placeholders, gender=""):
    return utils_render_template(template_text, placeholders, gender)

def _first_name(name):
    parts = value_to_text(name).strip().split()
    return parts[0] if parts else ""

def _display_person_names(values):
    return _join_unique(values)

def _build_coordinator_message(name, main_assignments, backup_assignments, template_text=None, gender=""):
    first_name = _first_name(name) or name
    main_section = _build_role_section("Coordinator", main_assignments) if main_assignments else ""
    backup_section = _build_role_section("Backup Coordinator", backup_assignments) if backup_assignments else ""
    sections = "\n\n".join(section for section in (main_section, backup_section) if section.strip())
    forms = _gender_forms(gender)
    closing = _apply_gender_text(DEFAULT_COORDINATOR_CLOSING, gender)
    template_text = template_text or DEFAULT_COORDINATOR_TEMPLATES["1. Coordinator Confirmation"]
    template_text = _template_for_display(template_text)
    placeholders = {
        "name": first_name,
        "first_name": first_name,
        "full_name": name,
        "gender": forms["gender"],
        "bek": forms["bek"],
        "leek": forms["leek"],
        "recipient_suffix": forms["recipient_suffix"],
        "sections": sections,
        "main_section": main_section,
        "backup_section": backup_section,
        "closing": closing,
    }
    message = _apply_gender_text(_render_template(template_text, placeholders, gender=gender), gender)
    return message.strip()

def _build_role_section(role_label, assignments):
    lines = [f"دورك كـ {role_label}:", ""]
    grouped = defaultdict(list)
    for assignment in assignments:
        row = assignment["row"]
        grouped[(row["day"], row["area"], row["city"], row["center_name"], row.get("center_code", ""))].append(row)

    for (day, area, city, center_name, _), rows in sorted(grouped.items()):
        display_day = _arabic_day(day)
        lines.append(f"ليوم {display_day} في {area} ({city}) بـ {center_name}، والخاصة بالمجموعات التالية:")
        lines.append("")
        lines.append(f"📅 مواعيد يوم {display_day}:")
        lines.append("")

        by_time = defaultdict(list)
        for row in rows:
            by_time[row["time"]].append(row)

        for slot_time, slot_rows in sorted(by_time.items()):
            lines.append(f"⏰ الساعة {slot_time}:")
            lines.append("")
            for group_code, room_code in _unique_pairs((r.get("group_code", ""), r.get("room_code", "")) for r in slot_rows):
                lines.append(f"Group: {group_code} (Room: {room_code})")
            lines.append("")
            lines.append("👥 المسؤولين عن المجموعات:")
            lines.append("Coordinator: " + (_display_person_names(r.get("main_coordinator", "") for r in slot_rows) or "لم يتم تحديده"))
            lines.append("Backup Coordinator: " + (_display_person_names(r.get("backup_coordinator", "") for r in slot_rows) or "لم يتم تحديده"))
            lines.append("")
    return "\n".join(lines)

def _slot_key(row, role):
    return (role, row.get("day", ""), row.get("time", ""), _center_key(row))

def _center_key(row):
    return row.get("center_code", "").strip() or row.get("center_name", "").strip()

def _paid_slots(assignments):
    return {_slot_key(a["row"], a["role"]) for a in assignments}

def _join_unique(values):
    seen = []
    for value in values:
        value = value.strip()
        if value and value not in seen:
            seen.append(value)
    return "، ".join(seen)

def _unique_pairs(pairs):
    seen = set()
    output = []
    for pair in pairs:
        if pair in seen:
            continue
        seen.add(pair)
        output.append(pair)
    return output

def _log_row(name, phone, status, error=""):
    return {
        "Timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Name": name or "",
        "Phone": phone or "",
        "Status": status,
        "Error": error or "",
    }

def _write_workbook(path, rows, headers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Output"
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header, "") for header in headers])
    for column_cells in ws.columns:
        width = min(max(len(value_to_text(cell.value)) for cell in column_cells) + 2, 70)
        ws.column_dimensions[column_cells[0].column_letter].width = width
    wb.save(path)

def generate_instructor_messages(rows):
    """Placeholder hook: the instructor GUI keeps its existing message flow."""
    return rows

def load_coordinator_templates() -> dict:
    if os.path.exists(COORDINATOR_TEMPLATES_FILE):
        try:
            with open(COORDINATOR_TEMPLATES_FILE, "r", encoding="utf-8") as f:
                templates = json.load(f)
            if isinstance(templates, dict) and templates:
                return templates
        except Exception:
            pass
    return dict(DEFAULT_COORDINATOR_TEMPLATES)

def save_coordinator_templates(templates: dict):
    with open(COORDINATOR_TEMPLATES_FILE, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=2)

def _is_null_coordinator(name):
    if not name:
        return True
    val = str(name).strip().lower()
    return val in {'0', '0.0', 'none', 'nan', '<na>', ''}

def _rows_to_audit_df(rows):
    df = pd.DataFrame(rows)
    for col in ['main_coordinator', 'backup_coordinator', 'center_code', 'day', 'time', 'group_code', 'room_code']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
        else:
            df[col] = ""
    return df

def calculate_slots_and_dues(rows, slot_price=250):
    if not rows:
        return [], {"Total Slots": 0, "Total Amount (EGP)": 0}
        
    df = _rows_to_audit_df(rows)
    
    orig_names = {}
    
    main_df = df[['main_coordinator', 'center_code', 'day', 'time']].copy()
    main_df['role'] = 'main'
    main_df.rename(columns={'main_coordinator': 'coordinator'}, inplace=True)
    
    backup_df = df[['backup_coordinator', 'center_code', 'day', 'time']].copy()
    backup_df['role'] = 'backup'
    backup_df.rename(columns={'backup_coordinator': 'coordinator'}, inplace=True)
    
    stacked = pd.concat([main_df, backup_df], ignore_index=True)
    
    def process_coord(name):
        if _is_null_coordinator(name):
            return None
        return name
        
    stacked['orig_name'] = stacked['coordinator'].apply(process_coord)
    stacked.dropna(subset=['orig_name'], inplace=True)
    stacked['coord_lower'] = stacked['orig_name'].str.lower()
    
    for _, row in stacked.iterrows():
        if row['coord_lower'] not in orig_names:
            orig_names[row['coord_lower']] = row['orig_name']
            
    slots = stacked.drop_duplicates(subset=['coord_lower', 'role', 'center_code', 'day', 'time'])
    
    if slots.empty:
        return [], {"Total Slots": 0, "Total Amount (EGP)": 0}
        
    counts = slots.groupby(['coord_lower', 'role']).size().unstack(fill_value=0)
    
    if 'main' not in counts.columns:
        counts['main'] = 0
    if 'backup' not in counts.columns:
        counts['backup'] = 0
        
    counts['Total Slots'] = counts['main'] + counts['backup']
    counts['Amount (EGP)'] = counts['Total Slots'] * slot_price
    
    analysis_list = []
    total_slots_all = 0
    total_amount_all = 0
    
    counts = counts.sort_values(by='Amount (EGP)', ascending=False)
    
    for coord_lower, row in counts.iterrows():
        total_slots_all += int(row['Total Slots'])
        total_amount_all += int(row['Amount (EGP)'])
        analysis_list.append({
            "Coordinator": orig_names[coord_lower],
            "Main Slots": int(row['main']),
            "Backup Slots": int(row['backup']),
            "Total Slots": int(row['Total Slots']),
            "Amount (EGP)": int(row['Amount (EGP)'])
        })
        
    totals_dict = {
        "Total Slots": total_slots_all,
        "Total Amount (EGP)": total_amount_all
    }
    
    return analysis_list, totals_dict

def detect_strict_conflicts(rows):
    if not rows:
        return {"coordinator_conflicts": [], "room_conflicts": []}
        
    df = _rows_to_audit_df(rows)
    
    coord_conflicts = []
    room_conflicts = []
    
    main_c = df[['main_coordinator', 'day', 'time', 'center_code']].copy()
    main_c.rename(columns={'main_coordinator': 'coordinator'}, inplace=True)
    
    backup_c = df[['backup_coordinator', 'day', 'time', 'center_code']].copy()
    backup_c.rename(columns={'backup_coordinator': 'coordinator'}, inplace=True)
    
    all_c = pd.concat([main_c, backup_c], ignore_index=True)
    all_c['coordinator'] = all_c['coordinator'].apply(lambda x: None if _is_null_coordinator(x) else x)
    all_c.dropna(subset=['coordinator'], inplace=True)
    
    if not all_c.empty:
        unique_c = all_c.drop_duplicates(subset=['coordinator', 'day', 'time', 'center_code'])
        c_groups = unique_c.groupby(['coordinator', 'day', 'time'])['center_code'].apply(list).reset_index()
        c_groups['center_count'] = c_groups['center_code'].apply(len)
        c_conflicts = c_groups[c_groups['center_count'] > 1]
        
        for _, row in c_conflicts.iterrows():
            coord_conflicts.append({
                "Coordinator": row['coordinator'],
                "Day": row['day'],
                "Time": row['time'],
                "Centers": row['center_code']
            })
            
    r_groups = df.groupby(['center_code', 'room_code', 'day', 'time'])['group_code'].unique().reset_index()
    r_groups['group_count'] = r_groups['group_code'].apply(len)
    r_conflicts = r_groups[r_groups['group_count'] > 1]
    
    for _, row in r_conflicts.iterrows():
        room_conflicts.append({
            "Center Code": row['center_code'],
            "Room Code": row['room_code'],
            "Day": row['day'],
            "Time": row['time'],
            "Groups": list(row['group_code'])
        })
        
    return {
        "coordinator_conflicts": coord_conflicts,
        "room_conflicts": room_conflicts
    }
