from __future__ import annotations
import os
import threading
import time
import random
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter.scrolledtext import ScrolledText
import customtkinter as ctk
from app_config import COLORS as C
from engine import MessageEngine
from ui_components import StatusDot, StyledButton, StyledEntry, StyledTextbox, StyledOptionMenu, CardFrame, SidebarButton
from collections import Counter, defaultdict
from coordinator_parser import clean_phone_number, value_to_text, load_data, validate_columns, _person_key
from file_errors import FileUploadError
from coordinator_logic import (
    load_coordinator_templates, save_coordinator_templates, generate_coordinator_messages,
    calculate_slot_analysis, detect_conflicts, preview_messages, send_messages, save_outputs,
    DEFAULT_SLOT_PRICE, DEFAULT_MAIN_LIMIT, DEFAULT_BACKUP_LIMIT, _template_for_display,
    DEFAULT_COORDINATOR_TEMPLATES, _log_row,
    calculate_slots_and_dues, detect_strict_conflicts,
)
from report_generator import generate_pdf_report, generate_xlsx_report

COORD_STATUS_READY = "Ready"
COORD_STATUS_SENDING = "Sending"
COORD_STATUS_SENT = "Sent"
COORD_STATUS_FAILED = "Failed"
COORD_STATUS_SKIPPED = "Skipped"
COORD_STATUS_COLORS = {
    COORD_STATUS_READY: C["status_yellow"],
    COORD_STATUS_SENDING: C["accent_cyan"],
    COORD_STATUS_SENT: C["status_green"],
    COORD_STATUS_FAILED: C["danger"],
    COORD_STATUS_SKIPPED: C["text_muted"],
}



def run_coordinator_wizard(parent=None):
    """GUI wizard for coordinator file selection, preview, optional sending, and outputs."""
    file_path = filedialog.askopenfilename(
        parent=parent,
        title="Choose sessions Excel or CSV file",
        filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("Excel", "*.xlsx *.xlsm"), ("CSV", "*.csv")],
    )
    if not file_path:
        return False

    coordinator_file_path = filedialog.askopenfilename(
        parent=parent,
        title="Choose coordinators data Excel or CSV file (Optional)",
        filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("Excel", "*.xlsx *.xlsm"), ("CSV", "*.csv")],
    )

    slot_price = simpledialog.askfloat(
        "Slot Price", "What is the slot price?", initialvalue=DEFAULT_SLOT_PRICE, minvalue=0, parent=parent
    )
    if slot_price is None:
        slot_price = DEFAULT_SLOT_PRICE
    main_limit = simpledialog.askinteger(
        "Main Coordinator Limit",
        "What is the maximum number of sessions a Main Coordinator can cover in the same slot?",
        initialvalue=DEFAULT_MAIN_LIMIT,
        minvalue=1,
        parent=parent,
    ) or DEFAULT_MAIN_LIMIT
    backup_limit = simpledialog.askinteger(
        "Backup Coordinator Limit",
        "What is the maximum number of sessions a Backup Coordinator can cover in the same slot?",
        initialvalue=DEFAULT_BACKUP_LIMIT,
        minvalue=1,
        parent=parent,
    ) or DEFAULT_BACKUP_LIMIT

    try:
        rows = validate_columns(load_data(file_path, coordinator_file_path))
        messages, skipped_logs = generate_coordinator_messages(rows, slot_price, main_limit, backup_limit)
        analysis = calculate_slot_analysis(rows, slot_price, main_limit, backup_limit)
        conflicts = detect_conflicts(rows, main_limit, backup_limit)
    except FileUploadError as exc:
        messagebox.showerror(
            exc.title,
            exc.message + ("\n\n" + "\n".join(f"• {s}" for s in exc.suggestions) if exc.suggestions else ""),
            parent=parent,
        )
        return False
    except Exception as exc:
        messagebox.showerror(
            "Coordinator Import Error",
            f"An unexpected error occurred while processing the files:\n\n{exc}\n\n"
            f"If this keeps happening, check that the files are valid and not open in another program.",
            parent=parent,
        )
        return False

    if not messages:
        skipped_text = "\n".join(
            f"{row.get('Name', '')} | {row.get('Phone', '')} | {row.get('Error', '')}"
            for row in skipped_logs
        ) or "No valid coordinator phone numbers were found."
        messagebox.showwarning(
            "No Valid Recipients",
            "No valid coordinator phone numbers were found.\n\nSkipped rows:\n" + skipped_text,
            parent=parent,
        )
        return False

    preview = preview_messages(messages)
    if conflicts:
        preview += "\n\nConflict warnings:\n" + "\n".join(
            f"- {c['Role']} {c['Coordinator']} has {c['Sessions']} sessions at {c['Day']} {c['Time']} "
            f"in {c['Center']} (limit {c['Limit']})"
            for c in conflicts
        )
    if skipped_logs:
        preview += f"\n\nSkipped invalid rows: {len(skipped_logs)}"

    _show_text_dialog(parent, "Coordinator Message Preview", preview)
    should_send = messagebox.askyesno("Send Messages", "Do you want to send these messages now? yes/no", parent=parent)

    log_rows = list(skipped_logs)
    if should_send and conflicts:
        should_send = messagebox.askyesno(
            "Conflicts Found",
            "Conflicts were found. Do you still want to continue sending? yes/no",
            parent=parent,
        )

    if should_send:
        messagebox.showinfo(
            "Sending Started",
            "WhatsApp sending will start now. Please keep WhatsApp Web signed in.",
            parent=parent,
        )
        log_rows.extend(send_messages(messages))

    messagebox.showinfo(
        "Coordinator Flow Complete",
        "Done. Messages, analysis, and sending log were kept inside the app flow; no output Excel files were generated.",
        parent=parent,
    )
    return True

def _show_text_dialog(parent, title, text):
    dlg = tk.Toplevel(parent) if parent else tk.Tk()
    dlg.title(title)
    dlg.geometry("780x560")
    if parent:
        dlg.transient(parent)
        dlg.grab_set()

    box = ScrolledText(dlg, wrap="word", font=("Segoe UI", 10))
    box.pack(fill="both", expand=True, padx=12, pady=(12, 8))
    box.insert("1.0", text)
    box.configure(state="disabled")

    btn = tk.Button(dlg, text="OK", width=14, command=dlg.destroy)
    btn.pack(pady=(0, 12))
    dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
    dlg.wait_window()

class CoordinatorApp(ctk.CTk):
    """Full coordinator workspace for upload, preview, editing, sending, and tracking."""

    def __init__(self):
        super().__init__()
        self.title("DECI Automation - Coordinator Messages")
        self.geometry("1280x780")
        self.minsize(1100, 650)
        self.configure(fg_color=C["bg_dark"])
        self.after(0, lambda: self.state("zoomed"))

        self.engine = MessageEngine()
        
        import db_manager
        db_manager.init_db()
        db_manager.clear_temporary_coordinator_states()
        self.date_filter_var = ctk.StringVar(value="All Saved")

        self.sessions_file = ""
        self.coordinators_file = ""
        self.output_dir = os.getcwd()
        self.rows = []
        self.messages = []
        self.analysis = []
        self.log_rows = []
        self.conflicts = []
        self.upload_summary_text = ""
        self.strict_analysis = []
        self.strict_totals = {}
        self.strict_conflicts = {"coordinator_conflicts": [], "room_conflicts": []}
        self.row_widgets = {}
        self._sending = False
        self._stop_event = threading.Event()
        self._search_after_id = None
        self.select_all_var = tk.BooleanVar(value=False)
        self.templates = load_coordinator_templates()

        self.slot_price_var = tk.StringVar(value=str(DEFAULT_SLOT_PRICE))
        self.main_limit_var = tk.StringVar(value=str(DEFAULT_MAIN_LIMIT))
        self.backup_limit_var = tk.StringVar(value=str(DEFAULT_BACKUP_LIMIT))
        self.delay_min_var = tk.StringVar(value="20")
        self.delay_max_var = tk.StringVar(value="40")
        self.wait_time_var = tk.StringVar(value="15")
        self.send_wait_var = tk.StringVar(value="8")
        self.close_time_var = tk.StringVar(value="10")
        self.search_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Upload the sessions file and coordinators data file to begin.")
        self.tpl_var = tk.StringVar(value=next(iter(self.templates), "1. Coordinator Confirmation"))

        self._build_layout()
        self._load_template_text()
        self._update_date_menu()
        self._render_table()

    def _build_layout(self):
        self.sidebar = ctk.CTkFrame(self, width=330, fg_color=C["bg_sidebar"], corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(
            self.sidebar,
            text="Coordinator\nMessages",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=C["text_primary"],
            justify="left",
        ).pack(anchor="w", padx=24, pady=(26, 8))

        ctk.CTkLabel(
            self.sidebar,
            text="Upload, review, edit, send, and track coordinator WhatsApp messages.",
            font=ctk.CTkFont(size=12),
            text_color=C["text_secondary"],
            wraplength=250,
            justify="left",
        ).pack(anchor="w", padx=24, pady=(0, 18))

        StyledButton(
            self.sidebar,
            text="Upload Files",
            width=250,
            color=C["accent_blue"],
            hover=C["accent_purple"],
            command=self._upload_files,
        ).pack(anchor="w", padx=24, pady=(0, 10))

        self._file_label = ctk.CTkLabel(
            self.sidebar,
            text="No files loaded.",
            font=ctk.CTkFont(size=11),
            text_color=C["text_muted"],
            wraplength=250,
            justify="left",
        )
        self._file_label.pack(anchor="w", padx=24, pady=(0, 18))

        tpl_frame = ctk.CTkFrame(self.sidebar, fg_color=C["bg_card"], corner_radius=8)
        tpl_frame.pack(fill="x", padx=18, pady=(0, 18))
        ctk.CTkLabel(
            tpl_frame,
            text="Templates",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=C["text_primary"],
        ).pack(anchor="w", padx=14, pady=(12, 8))
        self.tpl_menu = StyledOptionMenu(
            tpl_frame,
            variable=self.tpl_var,
            values=list(self.templates.keys()),
            width=245,
            fg_color=C["bg_input"],
            button_color=C["accent_blue"],
            command=lambda _: self._load_template_text(),
        )
        self.tpl_menu.pack(padx=14, pady=(0, 8))
        self.tpl_editor = StyledTextbox(tpl_frame, height=170, fg_color=C["bg_input"], wrap="word")
        self.tpl_editor.pack(fill="x", padx=14, pady=(0, 8))
        self.tpl_editor.bind("<KeyRelease>", lambda _e: self._update_template_hint())

        placeholder_row = ctk.CTkFrame(tpl_frame, fg_color="transparent")
        placeholder_row.pack(fill="x", padx=14, pady=(0, 8))
        for placeholder in ("{name}", "{sections}", "{g:مستعد:مستعدة}"):
            StyledButton(
                placeholder_row,
                text=placeholder,
                width=72,
                color=C["bg_input"],
                hover=C["hover"],
                command=lambda value=placeholder: self._insert_template_placeholder(value),
            ).pack(side="left", padx=2)

        tpl_btns = ctk.CTkFrame(tpl_frame, fg_color="transparent")
        tpl_btns.pack(fill="x", padx=14, pady=(0, 8))
        StyledButton(tpl_btns, text="Save", width=56, command=self._save_current_template).pack(side="left", padx=2)
        StyledButton(tpl_btns, text="New", width=52, color=C["success"], hover=C["status_green"], command=self._new_template).pack(side="left", padx=2)
        StyledButton(tpl_btns, text="Delete", width=64, color=C["danger"], hover="#991b1b", command=self._delete_template).pack(side="left", padx=2)

        tpl_btns2 = ctk.CTkFrame(tpl_frame, fg_color="transparent")
        tpl_btns2.pack(fill="x", padx=14, pady=(0, 14))
        StyledButton(tpl_btns2, text="Reset Defaults", width=112, color=C["text_muted"], hover=C["hover"], command=self._reset_templates).pack(side="left", padx=2)
        StyledButton(tpl_btns2, text="Regenerate", width=104, color=C["accent_cyan"], hover=C["accent_blue"], command=self._regenerate_from_template).pack(side="left", padx=2)
        self.template_hint = ctk.CTkLabel(
            tpl_frame,
            text="Use {name}, {first_name}, {full_name}, {sections}, {main_section}, {backup_section}, {closing}, {g:male:female}.",
            text_color=C["text_muted"],
            font=ctk.CTkFont(size=10),
            wraplength=245,
            justify="left",
        )
        self.template_hint.pack(anchor="w", padx=14, pady=(0, 12))



        self.send_selected_btn = StyledButton(
            self.sidebar,
            text="Send Selected",
            width=250,
            color=C["success"],
            hover=C["status_green"],
            command=self._send_selected,
        )
        self.send_selected_btn.pack(anchor="w", padx=24, pady=(0, 10))

        self.send_all_btn = StyledButton(
            self.sidebar,
            text="Send All Unsent",
            width=250,
            color=C["accent_cyan"],
            hover=C["accent_blue"],
            command=self._send_all,
        )
        self.send_all_btn.pack(anchor="w", padx=24, pady=(0, 10))

        self.stop_btn = StyledButton(
            self.sidebar,
            text="Stop Sending",
            width=250,
            color=C["danger"],
            hover="#991b1b",
            command=self._stop_sending,
        )
        self.stop_btn.pack(anchor="w", padx=24, pady=(0, 18))

        self.stats_label = ctk.CTkLabel(
            self.sidebar,
            text="Recipients: 0 | Sent: 0 | Failed: 0",
            text_color=C["text_secondary"],
            font=ctk.CTkFont(size=12),
            wraplength=250,
        )
        self.stats_label.pack(anchor="w", padx=24, pady=(0, 12))

        self.main_area = ctk.CTkFrame(self, fg_color=C["bg_dark"], corner_radius=0)
        self.main_area.pack(side="right", fill="both", expand=True)

        top = ctk.CTkFrame(self.main_area, fg_color=C["bg_dark"], height=74)
        top.pack(fill="x", padx=20, pady=(16, 8))
        top.pack_propagate(False)

        ctk.CTkLabel(
            top,
            text="Coordinator Dashboard",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=C["text_primary"],
        ).pack(side="left", padx=(0, 20))

        search = StyledEntry(
            top,
            textvariable=self.search_var,
            placeholder_text="Search name, phone, or role",
            width=280,
            fg_color=C["bg_input"],
        )
        search.pack(side="right", padx=(10, 0), pady=20)
        StyledButton(
            top,
            text="Settings",
            width=104,
            color=C["bg_input"],
            hover=C["hover"],
            command=self._open_settings_dialog,
        ).pack(side="right", padx=(10, 0), pady=20)
        self.search_var.trace_add("write", lambda *_: self._schedule_search())

        toolbar = ctk.CTkFrame(self.main_area, fg_color=C["bg_card"], height=48, corner_radius=8)
        toolbar.pack(fill="x", padx=20, pady=(0, 10))
        toolbar.pack_propagate(False)
        StyledButton(
            toolbar,
            text="Upload Files",
            width=118,
            color=C["accent_blue"],
            hover=C["accent_purple"],
            command=self._upload_files,
        ).pack(side="left", padx=(10, 6), pady=6)

        StyledButton(
            toolbar,
            text="Send Selected",
            width=122,
            color=C["success"],
            hover=C["status_green"],
            command=self._send_selected,
        ).pack(side="left", padx=6, pady=6)
        StyledButton(
            toolbar,
            text="Send All",
            width=104,
            color=C["accent_cyan"],
            hover=C["accent_blue"],
            command=self._send_all,
        ).pack(side="left", padx=6, pady=6)
        StyledButton(
            toolbar,
            text="Stop",
            width=82,
            color=C["danger"],
            hover="#991b1b",
            command=self._stop_sending,
        ).pack(side="left", padx=6, pady=6)
        StyledButton(
            toolbar,
            text="Log",
            width=82,
            color=C["bg_input"],
            hover=C["hover"],
            command=self._show_log_dialog,
        ).pack(side="left", padx=6, pady=6)
        StyledButton(
            toolbar,
            text="Summary",
            width=92,
            color=C["bg_input"],
            hover=C["hover"],
            command=self._show_upload_summary_dialog,
        ).pack(side="left", padx=6, pady=6)
        StyledButton(
            toolbar,
            text="Export Excel",
            width=100,
            color=C["bg_input"],
            hover=C["hover"],
            command=self._export_messages_quick,
        ).pack(side="left", padx=6, pady=6)

        self.date_filter_menu = StyledOptionMenu(
            toolbar,
            variable=self.date_filter_var,
            values=["All Saved"],
            width=120,
            fg_color=C["bg_input"],
            button_color=C["accent_blue"],
            command=self._on_date_filter_change,
        )
        self.delete_date_btn = ctk.CTkButton(toolbar, text="🗑️", width=30, fg_color="transparent",
                                             hover_color=C["danger"], text_color=C["text_primary"],
                                             command=self._delete_current_date)
        self.delete_date_btn.pack(side="right", padx=(0, 14), pady=6)

        self.date_filter_menu.pack(side="right", padx=(10, 2), pady=6)

        StyledButton(
            toolbar,
            text="Regenerate",
            width=112,
            color=C["accent_purple"],
            hover=C["accent_blue"],
            command=self._regenerate_from_template,
        ).pack(side="left", padx=6, pady=6)

        StyledButton(
            toolbar,
            text="Export Msgs",
            width=120,
            color=C["bg_input"],
            hover=C["hover"],
            command=self._export_messages_quick,
        ).pack(side="left", padx=6, pady=6)
        StyledButton(
            toolbar,
            text="Export Log",
            width=96,
            color=C["bg_input"],
            hover=C["hover"],
            command=self._export_log_quick,
        ).pack(side="left", padx=6, pady=6)

        self.status_label = ctk.CTkLabel(
            self.main_area,
            textvariable=self.status_var,
            text_color=C["accent_cyan"],
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        self.status_label.pack(fill="x", padx=20, pady=(0, 8))

        # ── Live progress panel — hidden until a batch send starts ──────────
        self.progress_panel = ctk.CTkFrame(self.main_area, fg_color=C["bg_card"], corner_radius=8)
        # Not packed here — revealed by _show_progress_panel() when sending starts
        _prog_inner = ctk.CTkFrame(self.progress_panel, fg_color="transparent")
        _prog_inner.pack(fill="x", padx=16, pady=(12, 8))
        self.progress_bar = ctk.CTkProgressBar(_prog_inner, height=14, corner_radius=7, progress_color=C["accent_cyan"])
        self.progress_bar.pack(fill="x", pady=(0, 8))
        self.progress_bar.set(0)
        _prog_row1 = ctk.CTkFrame(_prog_inner, fg_color="transparent")
        _prog_row1.pack(fill="x")
        self.progress_current_label = ctk.CTkLabel(
            _prog_row1, text="", text_color=C["text_primary"],
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w",
        )
        self.progress_current_label.pack(side="left")
        self.progress_counters_label = ctk.CTkLabel(
            _prog_row1, text="", text_color=C["text_secondary"],
            font=ctk.CTkFont(size=12), anchor="e",
        )
        self.progress_counters_label.pack(side="right")
        _prog_row2 = ctk.CTkFrame(_prog_inner, fg_color="transparent")
        _prog_row2.pack(fill="x", pady=(4, 0))
        self.progress_countdown_label = ctk.CTkLabel(
            _prog_row2, text="", text_color=C["accent_cyan"],
            font=ctk.CTkFont(size=11), anchor="w",
        )
        self.progress_countdown_label.pack(side="left")
        self.progress_eta_label = ctk.CTkLabel(
            _prog_row2, text="", text_color=C["text_muted"],
            font=ctk.CTkFont(size=11), anchor="e",
        )
        self.progress_eta_label.pack(side="right")

        self.tabview = ctk.CTkTabview(self.main_area, fg_color=C["bg_dark"])
        self.tabview.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.tab_messages = self.tabview.add("Messages")
        self.tab_analysis = self.tabview.add("Analysis")

        self.header_frame = ctk.CTkFrame(self.tab_messages, fg_color=C["bg_card"], height=38, corner_radius=0)
        self.header_frame.pack(fill="x", padx=0)
        self.header_frame.pack_propagate(False)
        ctk.CTkCheckBox(
            self.header_frame, text="", width=42,
            variable=self.select_all_var,
            command=self._toggle_select_all,
            checkbox_width=18, checkbox_height=18,
        ).pack(side="left", padx=3)
        for text, width in [
            ("Name", 190), ("Phone", 150), ("Role", 90), ("Main", 60),
            ("Backup", 70), ("Paid", 60), ("Amount", 80), ("Status", 100), ("Actions", 240),
        ]:
            ctk.CTkLabel(
                self.header_frame,
                text=text,
                width=width,
                anchor="w",
                text_color=C["text_secondary"],
                font=ctk.CTkFont(size=12, weight="bold"),
            ).pack(side="left", padx=3)

        self.table_frame = ctk.CTkScrollableFrame(self.tab_messages, fg_color=C["bg_dark"], corner_radius=0)
        self.table_frame.pack(fill="both", expand=True, padx=0, pady=(0, 10))

        preview_frame = ctk.CTkFrame(self.tab_messages, fg_color=C["bg_card"], height=170, corner_radius=8)
        preview_frame.pack(fill="x", padx=0, pady=(0, 16))
        preview_frame.pack_propagate(False)
        ctk.CTkLabel(
            preview_frame,
            text="Message Preview",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=C["text_primary"],
        ).pack(anchor="w", padx=12, pady=(10, 4))
        self.preview_box = StyledTextbox(preview_frame, fg_color=C["bg_input"], height=105, wrap="word")
        self.preview_box.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.preview_box.insert("1.0", "Select a row and click Preview.")
        self.preview_box.configure(state="disabled")

    def _add_setting_entry(self, parent, label, variable, bottom=6):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(0, bottom))
        ctk.CTkLabel(row, text=label, width=92, anchor="w", text_color=C["text_secondary"]).pack(side="left")
        StyledEntry(row, textvariable=variable, width=92, fg_color=C["bg_input"]).pack(side="right")

    def _open_settings_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Coordinator Settings")
        dlg.geometry("430x430")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        dlg.grab_set()

        content = ctk.CTkFrame(dlg, fg_color=C["bg_card"], corner_radius=8)
        content.pack(fill="both", expand=True, padx=18, pady=18)

        ctk.CTkLabel(
            content,
            text="Coordinator Settings",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=C["text_primary"],
        ).pack(anchor="w", padx=18, pady=(18, 12))

        settings = [
            ("Slot price", self.slot_price_var),
            ("Main limit", self.main_limit_var),
            ("Backup limit", self.backup_limit_var),
            ("Delay min", self.delay_min_var),
            ("Delay max", self.delay_max_var),
            ("Wait time", self.wait_time_var),
            ("Send wait", self.send_wait_var),
            ("Close time", self.close_time_var),
        ]
        for label, variable in settings:
            row = ctk.CTkFrame(content, fg_color="transparent")
            row.pack(fill="x", padx=18, pady=6)
            ctk.CTkLabel(row, text=label, width=130, anchor="w", text_color=C["text_secondary"]).pack(side="left")
            StyledEntry(row, textvariable=variable, width=160, fg_color=C["bg_input"]).pack(side="right")

        hint = ctk.CTkLabel(
            content,
            text="Slot price and limits affect generated analysis. Delay and wait time affect WhatsApp sending.",
            text_color=C["text_muted"],
            font=ctk.CTkFont(size=11),
            wraplength=360,
            justify="left",
        )
        hint.pack(anchor="w", padx=18, pady=(12, 8))

        def apply_settings(regenerate=False):
            try:
                self._get_float(self.slot_price_var, DEFAULT_SLOT_PRICE)
                self._get_int(self.main_limit_var, DEFAULT_MAIN_LIMIT)
                self._get_int(self.backup_limit_var, DEFAULT_BACKUP_LIMIT)
                self._get_int(self.delay_min_var, 20)
                self._get_int(self.delay_max_var, 40)
                self._get_int(self.wait_time_var, 15)
                self._get_int(self.send_wait_var, 8)
                self._get_int(self.close_time_var, 10)
            except ValueError as exc:
                messagebox.showerror("Invalid Settings", str(exc), parent=dlg)
                return
            self.status_var.set("Settings updated.")
            if regenerate and self.sessions_file and self.coordinators_file:
                self._regenerate_messages()
            dlg.destroy()

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.pack(fill="x", padx=18, pady=(12, 18))
        StyledButton(actions, text="Close", width=95, color=C["bg_input"], hover=C["hover"], command=dlg.destroy).pack(side="left")
        StyledButton(actions, text="Apply", width=95, color=C["accent_blue"], hover=C["accent_purple"], command=lambda: apply_settings(False)).pack(side="right", padx=(8, 0))
        StyledButton(actions, text="Apply + Regenerate", width=155, color=C["accent_cyan"], hover=C["accent_blue"], command=lambda: apply_settings(True)).pack(side="right")

    def _load_template_text(self):
        if not hasattr(self, "tpl_editor"):
            return
        name = self.tpl_var.get()
        self.tpl_editor.delete("1.0", "end")
        self.tpl_editor.insert("1.0", _template_for_display(self.templates.get(name, "")))
        self._update_template_hint()

    def _get_template_text(self):
        return self.tpl_editor.get("1.0", "end").strip()

    def _save_current_template(self):
        name = self.tpl_var.get().strip()
        if not name:
            return
        self.templates[name] = self._get_template_text()
        save_coordinator_templates(self.templates)
        self._refresh_template_menu(name)
        self.status_var.set(f"Template saved: {name}")

    def _new_template(self):
        name = simpledialog.askstring("New Template", "Template name:", parent=self)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        self.templates[name] = ""
        save_coordinator_templates(self.templates)
        self._refresh_template_menu(name)
        self._load_template_text()

    def _delete_template(self):
        name = self.tpl_var.get()
        if len(self.templates) <= 1:
            messagebox.showwarning("Cannot Delete", "Keep at least one coordinator template.", parent=self)
            return
        if not messagebox.askyesno("Delete Template", f"Delete template '{name}'?", parent=self):
            return
        self.templates.pop(name, None)
        save_coordinator_templates(self.templates)
        self._refresh_template_menu(next(iter(self.templates)))
        self._load_template_text()

    def _reset_templates(self):
        if not messagebox.askyesno(
            "Reset Templates",
            "Reset coordinator templates to defaults? Custom coordinator templates will be removed.",
            parent=self,
        ):
            return
        self.templates = dict(DEFAULT_COORDINATOR_TEMPLATES)
        save_coordinator_templates(self.templates)
        self._refresh_template_menu(next(iter(self.templates)))
        self._load_template_text()

    def _refresh_template_menu(self, selected=None):
        selected = selected or self.tpl_var.get()
        values = list(self.templates.keys())
        self.tpl_menu.configure(values=values)
        if selected not in self.templates:
            selected = values[0]
        self.tpl_var.set(selected)

    def _insert_template_placeholder(self, placeholder):
        self.tpl_editor.insert("insert", placeholder)
        self._update_template_hint()

    def _update_template_hint(self):
        if not hasattr(self, "template_hint"):
            return
        text = self._get_template_text() if hasattr(self, "tpl_editor") else ""
        if "{sections}" not in text and "{main_section}" not in text and "{backup_section}" not in text:
            self.template_hint.configure(
                text="Tip: add {sections} where the generated slot details should appear.",
                text_color=C["status_yellow"],
            )
        else:
            self.template_hint.configure(
                text="Use {name}, {first_name}, {full_name}, {gender}, {bek}, {leek}, {sections}. Closing text is visible and editable.",
                text_color=C["text_muted"],
            )

    def _regenerate_from_template(self):
        self._save_current_template()
        if not self.sessions_file or not self.coordinators_file:
            messagebox.showinfo("No Files Loaded", "Upload files before regenerating messages.", parent=self)
            return
        self._regenerate_messages()

    def _open_template_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Coordinator Templates")
        dlg.geometry("760x620")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        dlg.grab_set()

        content = ctk.CTkFrame(dlg, fg_color=C["bg_card"], corner_radius=8)
        content.pack(fill="both", expand=True, padx=18, pady=18)

        selected_var = tk.StringVar(value=self.tpl_var.get())
        top_row = ctk.CTkFrame(content, fg_color="transparent")
        top_row.pack(fill="x", padx=14, pady=(14, 8))
        ctk.CTkLabel(top_row, text="Template", width=85, anchor="w", text_color=C["text_secondary"]).pack(side="left")
        menu = StyledOptionMenu(
            top_row,
            variable=selected_var,
            values=list(self.templates.keys()),
            width=260,
            fg_color=C["bg_input"],
            button_color=C["accent_blue"],
        )
        menu.pack(side="left")

        editor = StyledTextbox(content, fg_color=C["bg_input"], wrap="word")
        editor.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        hint = ctk.CTkLabel(
            content,
            text="Placeholders: {name}/{first_name} = first name, {full_name}, {gender}, {bek}, {leek}, {sections}, {main_section}, {backup_section}. Closing text is visible and editable.",
            text_color=C["text_muted"],
            font=ctk.CTkFont(size=11),
            anchor="w",
        )
        hint.pack(fill="x", padx=14, pady=(0, 8))

        def load_selected(_=None):
            editor.delete("1.0", "end")
            editor.insert("1.0", _template_for_display(self.templates.get(selected_var.get(), "")))

        def refresh_menu(selected=None):
            values = list(self.templates.keys())
            menu.configure(values=values)
            selected_var.set(selected if selected in self.templates else values[0])
            load_selected()

        def save_selected():
            name = selected_var.get().strip()
            if not name:
                return
            self.templates[name] = editor.get("1.0", "end").strip()
            save_coordinator_templates(self.templates)
            self._refresh_template_menu(name)
            self._load_template_text()
            self.status_var.set(f"Template saved: {name}")

        def new_template():
            name = simpledialog.askstring("New Template", "Template name:", parent=dlg)
            if not name:
                return
            name = name.strip()
            if not name:
                return
            self.templates[name] = ""
            save_coordinator_templates(self.templates)
            self._refresh_template_menu(name)
            refresh_menu(name)

        def delete_template():
            name = selected_var.get()
            if len(self.templates) <= 1:
                messagebox.showwarning("Cannot Delete", "Keep at least one coordinator template.", parent=dlg)
                return
            if not messagebox.askyesno("Delete Template", f"Delete template '{name}'?", parent=dlg):
                return
            self.templates.pop(name, None)
            save_coordinator_templates(self.templates)
            refresh_menu(next(iter(self.templates)))
            self._refresh_template_menu(selected_var.get())
            self._load_template_text()

        def reset_templates():
            if not messagebox.askyesno("Reset Templates", "Reset coordinator templates to defaults?", parent=dlg):
                return
            self.templates = dict(DEFAULT_COORDINATOR_TEMPLATES)
            save_coordinator_templates(self.templates)
            refresh_menu(next(iter(self.templates)))
            self._refresh_template_menu(selected_var.get())
            self._load_template_text()

        def regenerate():
            save_selected()
            self.tpl_var.set(selected_var.get())
            self._load_template_text()
            if self.sessions_file and self.coordinators_file:
                self._regenerate_messages()

        menu.configure(command=load_selected)
        load_selected()

        btns = ctk.CTkFrame(content, fg_color="transparent")
        btns.pack(fill="x", padx=14, pady=(0, 14))
        StyledButton(btns, text="Save", width=90, command=save_selected).pack(side="left", padx=3)
        StyledButton(btns, text="New", width=82, color=C["success"], hover=C["status_green"], command=new_template).pack(side="left", padx=3)
        StyledButton(btns, text="Delete", width=88, color=C["danger"], hover="#991b1b", command=delete_template).pack(side="left", padx=3)
        StyledButton(btns, text="Reset Defaults", width=126, color=C["text_muted"], hover=C["hover"], command=reset_templates).pack(side="left", padx=3)
        StyledButton(btns, text="Save + Regenerate", width=150, color=C["accent_cyan"], hover=C["accent_blue"], command=regenerate).pack(side="right", padx=3)
        StyledButton(btns, text="Close", width=82, color=C["bg_input"], hover=C["hover"], command=dlg.destroy).pack(side="right", padx=3)

    def _upload_files(self):
        sessions_file = filedialog.askopenfilename(
            parent=self,
            title="Choose sessions Excel or CSV file",
            filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("Excel", "*.xlsx *.xlsm"), ("CSV", "*.csv")],
        )
        if not sessions_file:
            return
        coordinators_file = filedialog.askopenfilename(
            parent=self,
            title="Choose coordinators data Excel or CSV file (Optional)",
            filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("Excel", "*.xlsx *.xlsm"), ("CSV", "*.csv")],
        )

        self.sessions_file = sessions_file
        self.coordinators_file = coordinators_file
        self.output_dir = os.path.dirname(sessions_file)
        self._regenerate_messages()

    def _regenerate_messages(self):
        try:
            slot_price = self._get_float(self.slot_price_var, DEFAULT_SLOT_PRICE)
            main_limit = self._get_int(self.main_limit_var, DEFAULT_MAIN_LIMIT)
            backup_limit = self._get_int(self.backup_limit_var, DEFAULT_BACKUP_LIMIT)
            data = load_data(self.sessions_file, self.coordinators_file)
            self.rows = validate_columns(data)
            messages, skipped_logs = generate_coordinator_messages(
                self.rows,
                slot_price,
                main_limit,
                backup_limit,
                self._get_template_text(),
            )
            self.analysis = calculate_slot_analysis(self.rows, slot_price, main_limit, backup_limit)
            self.conflicts = detect_conflicts(self.rows, main_limit, backup_limit)
            self.strict_analysis, self.strict_totals = calculate_slots_and_dues(self.rows, slot_price)
            self.strict_conflicts = detect_strict_conflicts(self.rows)
        except FileUploadError as exc:
            messagebox.showerror(
                exc.title,
                exc.message + ("\n\n" + "\n".join(f"• {s}" for s in exc.suggestions) if exc.suggestions else ""),
                parent=self,
            )
            self.status_var.set(f"Import failed: {exc.title}")
            return
        except Exception as exc:
            messagebox.showerror(
                "Coordinator Import Error",
                f"An unexpected error occurred while processing the files:\n\n{exc}\n\n"
                f"If this keeps happening, check that the files are valid and not open in another program.",
                parent=self,
            )
            self.status_var.set(f"Import failed: {exc}")
            return

        self.messages = []
        for msg in messages:
            item = dict(msg)
            item["Selected"] = False
            item["Status"] = COORD_STATUS_READY
            item["Error"] = ""
            self.messages.append(item)

        self.log_rows = list(skipped_logs)
        self.upload_summary_text = self._build_upload_summary(skipped_logs)
        self._file_label.configure(
            text=f"Sessions:\n{os.path.basename(self.sessions_file)}\n\nCoordinators:\n{os.path.basename(self.coordinators_file) if self.coordinators_file else 'None'}"
        )
        conflict_note = f" | Conflicts: {len(self.conflicts)}" if self.conflicts else ""
        self.status_var.set(f"Generated {len(self.messages)} coordinator messages.{conflict_note}")
        self._render_table()
        self._refresh_stats()
        if self.messages:
            self._show_message_in_preview(0)
        self._show_upload_summary_if_needed()
        if self.conflicts:
            self._show_conflicts()

        save_for_later = messagebox.askyesno("Save Sessions", "Files loaded and messages generated successfully.\n\nDo you want to save these sessions in the database for later?\n(If No, they will be deleted when you close the app)", parent=self)
        
        if save_for_later:
            from datetime import date
            today_str = date.today().strftime("%Y-%m-%d")
            target_date = simpledialog.askstring("Save Sessions", "Enter the target date (e.g., Friday 24 Oct, or 2026-07-29):", initialvalue=today_str, parent=self)
            if not target_date or not target_date.strip():
                target_date = "Temporary"
            else:
                target_date = target_date.strip()
        else:
            target_date = "Temporary"

        state_dict = {
            "rows": self.rows,
            "messages": self.messages,
            "analysis": self.analysis,
            "log_rows": self.log_rows,
            "conflicts": self.conflicts,
            "upload_summary_text": self.upload_summary_text,
            "strict_analysis": self.strict_analysis,
            "strict_totals": self.strict_totals,
            "strict_conflicts": self.strict_conflicts,
            "sessions_file": self.sessions_file,
            "coordinators_file": self.coordinators_file
        }
        import db_manager
        db_manager.save_coordinator_state(target_date, state_dict)
        self._update_date_menu()
        self.date_filter_var.set(target_date)


    def _update_date_menu(self):
        import db_manager
        dates = db_manager.get_all_coordinator_dates()
        vals = ["All Saved"] + dates
        if hasattr(self, 'date_filter_menu'):
            self.date_filter_menu.configure(values=vals)
            if self.date_filter_var.get() not in vals:
                self.date_filter_var.set("All Saved")
                
    def _delete_current_date(self):
        date_val = self.date_filter_var.get()
        if date_val == "All Saved":
            messagebox.showwarning("Cannot Delete", "Please select a specific date to delete.")
            return
            
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete all saved data for {date_val}?"):
            import db_manager
            db_manager.delete_date(date_val)
            self.date_filter_var.set("All Saved")
            self._update_date_menu()
            self._on_date_filter_change(None)

    def _on_date_filter_change(self, _=None):
        date_filter = self.date_filter_var.get()
        if date_filter != "All Saved":
            self._load_state_from_db(date_filter)
        else:
            self.messages = []
            self.rows = []
            self._render_table()
            self._file_label.configure(text="No files loaded.")
            self.status_var.set("Select a saved date or upload files.")

    def _load_state_from_db(self, target_date):
        import db_manager
        state = db_manager.get_coordinator_state(target_date)
        if state:
            self.rows = state.get("rows", [])
            self.messages = state.get("messages", [])
            self.analysis = state.get("analysis", [])
            self.log_rows = state.get("log_rows", [])
            self.conflicts = state.get("conflicts", [])
            self.upload_summary_text = state.get("upload_summary_text", "")
            self.strict_analysis = state.get("strict_analysis", [])
            self.strict_totals = state.get("strict_totals", {})
            self.strict_conflicts = state.get("strict_conflicts", {"coordinator_conflicts": [], "room_conflicts": []})
            self.sessions_file = state.get("sessions_file", "")
            self.coordinators_file = state.get("coordinators_file", "")
            
            s_file = os.path.basename(self.sessions_file) if self.sessions_file else "None"
            c_file = os.path.basename(self.coordinators_file) if self.coordinators_file else "None"
            self._file_label.configure(
                text=f"Loaded DB Date: {target_date}\nSessions: {s_file}\nCoordinators: {c_file}"
            )
            conflict_note = f" | Conflicts: {len(self.conflicts)}" if self.conflicts else ""
            self.status_var.set(f"Loaded {len(self.messages)} coordinator messages from DB.{conflict_note}")
            
            self._render_table()
            self._refresh_stats()
            if self.messages:
                self._show_message_in_preview(0)

    def _render_table(self):
        for child in self.table_frame.winfo_children():
            child.destroy()
        self.row_widgets.clear()

        query = self.search_var.get().strip().casefold()
        visible_indexes = [
            i for i, msg in enumerate(self.messages)
            if not query
            or query in msg.get("Name", "").casefold()
            or query in msg.get("Phone", "").casefold()
            or query in msg.get("Role", "").casefold()
        ]

        if not visible_indexes:
            ctk.CTkLabel(
                self.table_frame,
                text="No coordinator messages loaded.",
                text_color=C["text_muted"],
                font=ctk.CTkFont(size=14),
            ).pack(pady=40)
            self._refresh_stats()
            return

        for idx in visible_indexes:
            msg = self.messages[idx]
            base_color = C["bg_card"] if idx % 2 == 0 else C["bg_dark"]
            row = ctk.CTkFrame(self.table_frame, fg_color=base_color, height=46, corner_radius=0)
            row.pack(fill="x", pady=1)
            row.pack_propagate(False)

            # Hover highlight + click-to-preview on the entire row
            def _on_enter(e, r=row): r.configure(fg_color=C["hover"])
            def _on_leave(e, r=row, c=base_color): r.configure(fg_color=c)
            row.bind("<Enter>", _on_enter)
            row.bind("<Leave>", _on_leave)
            row.bind("<Button-1>", lambda e, i=idx: self._show_message_in_preview(i))

            selected_var = tk.BooleanVar(value=msg.get("Selected", False))
            cb = ctk.CTkCheckBox(
                row,
                text="",
                width=42,
                variable=selected_var,
                command=lambda i=idx, var=selected_var: self._toggle_selected(i, var),
            )
            cb.pack(side="left", padx=3)
            self._row_label(row, msg.get("Name", ""), 190)
            self._row_label(row, msg.get("Phone", ""), 150)
            self._row_label(row, msg.get("Role", ""), 90)
            self._row_label(row, str(msg.get("Main Slots", 0)), 60)
            self._row_label(row, str(msg.get("Backup Slots", 0)), 70)
            self._row_label(row, str(msg.get("Total Paid Slots", 0)), 60)
            self._row_label(row, str(msg.get("Amount", 0)), 80)

            status_frame = ctk.CTkFrame(row, fg_color="transparent", width=100)
            status_frame.pack(side="left", padx=3)
            status_frame.pack_propagate(False)
            dot = StatusDot(status_frame, color=COORD_STATUS_COLORS.get(msg.get("Status"), C["text_muted"]), size=12)
            dot.pack(side="left", padx=(0, 6), pady=17)
            status_lbl = ctk.CTkLabel(
                status_frame,
                text=msg.get("Status", COORD_STATUS_READY),
                text_color=COORD_STATUS_COLORS.get(msg.get("Status"), C["text_secondary"]),
                font=ctk.CTkFont(size=11),
            )
            status_lbl.pack(side="left")

            actions = ctk.CTkFrame(row, fg_color="transparent", width=240)
            actions.pack(side="left", padx=3)
            actions.pack_propagate(False)
            StyledButton(actions, text="Send", width=50, color=C["success"], hover=C["status_green"], command=lambda i=idx: self._send_single(i)).pack(side="left", padx=1, pady=5)
            StyledButton(actions, text="Preview", width=64, color=C["accent_blue"], command=lambda i=idx: self._show_message_in_preview(i)).pack(side="left", padx=1, pady=5)
            StyledButton(actions, text="Edit", width=54, color=C["accent_purple"], command=lambda i=idx: self._edit_message(i)).pack(side="left", padx=1, pady=5)
            StyledButton(actions, text="Reset", width=54, color=C["text_muted"], hover=C["hover"], command=lambda i=idx: self._reset_message(i)).pack(side="left", padx=1, pady=5)

            self.row_widgets[idx] = {"selected": selected_var, "dot": dot, "status": status_lbl, "frame": row}

        self._refresh_stats()
        self._render_analysis()

    def _row_label(self, parent, text, width):
        display = value_to_text(text)
        if len(display) > 24:
            display = display[:22] + "..."
        ctk.CTkLabel(
            parent,
            text=display,
            width=width,
            anchor="w",
            text_color=C["text_primary"],
            font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=3)

    def _toggle_selected(self, idx, variable):
        self.messages[idx]["Selected"] = variable.get()
        self._refresh_stats()

    def _show_message_in_preview(self, idx):
        if idx >= len(self.messages):
            return
        msg = self.messages[idx]
        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", msg.get("Message", ""))
        self.preview_box.configure(state="disabled")
        self.status_var.set(f"Previewing message for {msg.get('Name', '')}.")

    def _edit_message(self, idx):
        msg = self.messages[idx]
        dlg = ctk.CTkToplevel(self)
        dlg.title(f"Edit - {msg.get('Name', '')}")
        dlg.geometry("720x620")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        dlg.grab_set()

        content = ctk.CTkFrame(dlg, fg_color=C["bg_card"], corner_radius=8)
        content.pack(fill="both", expand=True, padx=18, pady=18)

        name_var = tk.StringVar(value=msg.get("Name", ""))
        phone_var = tk.StringVar(value=msg.get("Phone", ""))
        for label, variable in (("Name", name_var), ("Phone", phone_var)):
            row = ctk.CTkFrame(content, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=(14, 0))
            ctk.CTkLabel(row, text=label, width=70, anchor="w", text_color=C["text_secondary"]).pack(side="left")
            StyledEntry(row, textvariable=variable, fg_color=C["bg_input"]).pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(content, text="Message", text_color=C["text_secondary"], anchor="w").pack(anchor="w", padx=14, pady=(16, 4))
        message_box = StyledTextbox(content, fg_color=C["bg_input"], wrap="word")
        message_box.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        message_box.insert("1.0", msg.get("Message", ""))

        def save():
            msg["Name"] = name_var.get().strip()
            msg["Phone"] = clean_phone_number(phone_var.get()) or phone_var.get().strip()
            msg["Message"] = message_box.get("1.0", "end").strip()
            self._render_table()
            self._show_message_in_preview(idx)
            dlg.destroy()

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.pack(fill="x", padx=14, pady=(0, 14))
        StyledButton(actions, text="Cancel", width=110, color=C["bg_input"], hover=C["hover"], command=dlg.destroy).pack(side="left")
        StyledButton(actions, text="Save", width=110, color=C["success"], hover=C["status_green"], command=save).pack(side="right")

    def _reset_message(self, idx):
        self._set_message_status(idx, COORD_STATUS_READY, "")

    def _send_selected(self):
        targets = [i for i, msg in enumerate(self.messages) if msg.get("Selected") and msg.get("Status") != COORD_STATUS_SENT]
        if not targets:
            messagebox.showwarning("Nothing Selected", "Select coordinator rows to send.", parent=self)
            return
        self._send_targets(targets)

    def _send_all(self):
        targets = [i for i, msg in enumerate(self.messages) if msg.get("Status") != COORD_STATUS_SENT]
        if not targets:
            messagebox.showinfo("Done", "All coordinator messages are already sent.", parent=self)
            return
        self._send_targets(targets)

    def _send_targets(self, indexes):
        if self._sending:
            messagebox.showinfo("Sending", "A sending batch is already running.", parent=self)
            return
        if self.conflicts and not messagebox.askyesno(
            "Conflicts Found",
            "Conflicts were found. Do you still want to continue sending?",
            parent=self,
        ):
            return
        if not messagebox.askyesno("Send Messages", f"Send {len(indexes)} coordinator message(s) now?", parent=self):
            return

        try:
            min_delay = self._get_int(self.delay_min_var, 20)
            max_delay = self._get_int(self.delay_max_var, 40)
            wait_time = self._get_int(self.wait_time_var, 15)
            send_wait = self._get_int(self.send_wait_var, 8)
            close_time = self._get_int(self.close_time_var, 10)
        except ValueError as exc:
            messagebox.showerror("Invalid Settings", str(exc), parent=self)
            return
        if min_delay > max_delay:
            min_delay, max_delay = max_delay, min_delay

        self.engine.min_delay = min_delay
        self.engine.max_delay = max_delay
        self.engine.wait_time = wait_time
        self.engine.post_send_delay = send_wait
        self.engine.close_time = close_time
        self._stop_event.clear()
        self._sending = True
        self._set_send_buttons(False)
        self._update_status_color("info")
        self.status_var.set(f"Starting batch of {len(indexes)} message(s)\u2026")
        self._show_progress_panel(len(indexes))
        thread = threading.Thread(target=self._send_worker, args=(indexes, min_delay, max_delay), daemon=True)
        thread.start()

    def _send_worker(self, indexes, min_delay, max_delay):
        total = len(indexes)
        sent_ok = 0
        sent_failed = 0
        sent_skipped = 0
        processed = 0
        start_time = time.time()

        for pos, idx in enumerate(indexes):
            if self._stop_event.is_set():
                break
            msg = self.messages[idx]
            phone = clean_phone_number(msg.get("Phone", ""))
            name = msg.get("Name", "Unknown")

            if not phone:
                sent_skipped += 1
                processed += 1
                self.after(0, lambda i=idx: self._set_message_status(i, COORD_STATUS_SKIPPED, "Invalid phone number"))
                self.log_rows.append(_log_row(name, msg.get("Phone", ""), "skipped", "Invalid phone number"))
                prog = (total, processed, sent_ok, sent_failed, sent_skipped, name, 0, 0)
                self.after(0, lambda p=prog: self._update_send_progress(*p))
                continue

            msg["Phone"] = phone
            self.after(0, lambda i=idx: self._set_message_status(i, COORD_STATUS_SENDING, ""))
            prog = (total, processed, sent_ok, sent_failed, sent_skipped, name, 0, 0)
            self.after(0, lambda p=prog: self._update_send_progress(*p))
            self.after(0, lambda i=idx: self._scroll_to_row(i))

            ok, detail = self.engine.send_message(phone, msg.get("Message", ""))
            status = COORD_STATUS_SENT if ok else COORD_STATUS_FAILED
            error = "" if ok else detail

            if ok:
                sent_ok += 1
            else:
                sent_failed += 1
            processed += 1

            self.log_rows.append(_log_row(name, phone, "sent" if ok else "failed", error))
            self.after(0, lambda i=idx, s=status, e=error: self._set_message_status(i, s, e))

            if pos < len(indexes) - 1 and not self._stop_event.is_set():
                delay = random.randint(min_delay, max_delay)
                elapsed = time.time() - start_time
                avg_time = elapsed / processed if processed > 0 else (min_delay + max_delay) / 2
                remaining_msgs = total - processed

                for rem in range(delay, 0, -1):
                    if self._stop_event.is_set():
                        break
                    eta_secs = int(remaining_msgs * avg_time + rem)
                    prog = (total, processed, sent_ok, sent_failed, sent_skipped, name, rem, eta_secs)
                    self.after(0, lambda p=prog: self._update_send_progress(*p))
                    time.sleep(1)

        self.after(0, self._finish_sending)

    def _finish_sending(self):
        self._sending = False
        self._set_send_buttons(True)
        stopped = self._stop_event.is_set()
        sent = sum(1 for m in self.messages if m.get("Status") == COORD_STATUS_SENT)
        failed = sum(1 for m in self.messages if m.get("Status") == COORD_STATUS_FAILED)
        skipped = sum(1 for m in self.messages if m.get("Status") == COORD_STATUS_SKIPPED)
        if stopped:
            self._update_status_color("warn")
            self.status_var.set(f"Sending stopped \u2014 Sent: {sent} | Failed: {failed} | Skipped: {skipped}")
        else:
            self._update_status_color("success")
            self.status_var.set(f"Batch complete \u2014 Sent: {sent} | Failed: {failed} | Skipped: {skipped}")
            messagebox.showinfo(
                "Batch Complete",
                f"Coordinator batch complete.\nSent: {sent} | Failed: {failed} | Skipped: {skipped}\n\nStatuses and logs are available inside the app.",
                parent=self,
            )
        # Keep panel visible for a few seconds then hide
        self.after(5000, self._hide_progress_panel)

    def _stop_sending(self):
        self._stop_event.set()
        self.engine.stop()
        self._update_status_color("warn")
        self.status_var.set("Stopping after the current WhatsApp action finishes\u2026")

    def _set_send_buttons(self, enabled):
        state = "normal" if enabled else "disabled"
        self.send_selected_btn.configure(state=state)
        self.send_all_btn.configure(state=state)

    def _set_message_status(self, idx, status, error=""):
        if idx >= len(self.messages):
            return
        self.messages[idx]["Status"] = status
        self.messages[idx]["Error"] = error
        widgets = self.row_widgets.get(idx)
        color = COORD_STATUS_COLORS.get(status, C["text_muted"])
        if widgets:
            widgets["dot"].set_color(color)
            widgets["status"].configure(text=status, text_color=color)
        if error:
            self.status_var.set(error)
        self._refresh_stats()

    def _show_conflicts(self):
        lines = [
            f"{c['Role']} {c['Coordinator']} has {c['Sessions']} sessions at {c['Day']} {c['Time']} "
            f"in {c['Center']} (limit {c['Limit']})"
            for c in self.conflicts
        ]
        _show_text_dialog(self, "Coordinator Conflict Warnings", "\n".join(lines))

    def _render_analysis(self):
        for child in self.tab_analysis.winfo_children():
            child.destroy()
            
        if not self.strict_analysis and not self.analysis:
            lbl = ctk.CTkLabel(self.tab_analysis, text="No analysis data available. Upload files or select a date.", text_color=C["text_muted"], font=ctk.CTkFont(size=14))
            lbl.pack(pady=40)
            return

        top_bar = ctk.CTkFrame(self.tab_analysis, fg_color="transparent")
        top_bar.pack(fill="x", padx=20, pady=(16, 0))
        StyledButton(top_bar, text="Export Report", width=130, color=C["accent_blue"], hover=C["accent_purple"], command=self._show_report_export_dialog).pack(side="right")

        # ── Summary cards at top ──────────────────────────────────────────
        cards_frame = ctk.CTkFrame(self.tab_analysis, fg_color="transparent", height=80)
        cards_frame.pack(fill="x", padx=20, pady=(8, 8))
        cards_frame.pack_propagate(False)

        total_coordinators = len(self.strict_analysis)
        total_slots = self.strict_totals.get("Total Slots", 0)
        total_amount = self.strict_totals.get("Total Amount (EGP)", 0)
        coord_conflicts = len(self.strict_conflicts.get("coordinator_conflicts", []))
        room_conflicts = len(self.strict_conflicts.get("room_conflicts", []))
        total_conflicts = coord_conflicts + room_conflicts

        card_data = [
            ("👥 Coordinators", str(total_coordinators), C["accent_blue"]),
            ("📊 Total Slots", str(total_slots), C["accent_cyan"]),
            ("💰 Total Amount", f"{total_amount:,.0f} EGP", C["status_green"]),
            ("⚠️ Conflicts", str(total_conflicts), C["danger"] if total_conflicts > 0 else C["status_green"]),
        ]
        for title, value, color in card_data:
            card = ctk.CTkFrame(cards_frame, fg_color=C["bg_card"], corner_radius=10)
            card.pack(side="left", fill="both", expand=True, padx=6)
            ctk.CTkLabel(
                card, text=title, text_color=C["text_secondary"],
                font=ctk.CTkFont(size=11),
            ).pack(anchor="w", padx=14, pady=(10, 0))
            ctk.CTkLabel(
                card, text=value, text_color=color,
                font=ctk.CTkFont(size=22, weight="bold"),
            ).pack(anchor="w", padx=14, pady=(2, 10))

        # ── Scrollable body ───────────────────────────────────────────────
        body = ctk.CTkScrollableFrame(self.tab_analysis, fg_color=C["bg_dark"])
        body.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # ── Section 1: Slots & Dues Table ─────────────────────────────────
        section1 = ctk.CTkFrame(body, fg_color=C["bg_card"], corner_radius=10)
        section1.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(
            section1, text="📋 Coordinator Slots & Financial Dues",
            font=ctk.CTkFont(size=16, weight="bold"), text_color=C["text_primary"],
        ).pack(anchor="w", padx=16, pady=(14, 8))

        # Table header
        col_defs = [
            ("#", 40), ("Coordinator", 220), ("Main Slots", 90),
            ("Backup Slots", 100), ("Total Slots", 90), ("Amount (EGP)", 120),
        ]
        header_row = ctk.CTkFrame(section1, fg_color=C["bg_input"], height=36, corner_radius=6)
        header_row.pack(fill="x", padx=12, pady=(0, 2))
        header_row.pack_propagate(False)
        for col_name, col_width in col_defs:
            ctk.CTkLabel(
                header_row, text=col_name, width=col_width, anchor="w",
                text_color=C["text_secondary"], font=ctk.CTkFont(size=12, weight="bold"),
            ).pack(side="left", padx=6)

        # Table rows
        for i, entry in enumerate(self.strict_analysis):
            row_color = C["bg_card"] if i % 2 == 0 else C["bg_dark"]
            row_frame = ctk.CTkFrame(section1, fg_color=row_color, height=34, corner_radius=0)
            row_frame.pack(fill="x", padx=12, pady=0)
            row_frame.pack_propagate(False)
            values = [
                str(i + 1),
                entry.get("Coordinator", ""),
                str(entry.get("Main Slots", 0)),
                str(entry.get("Backup Slots", 0)),
                str(entry.get("Total Slots", 0)),
                f"{entry.get('Amount (EGP)', 0):,.0f}",
            ]
            for (_, col_width), val in zip(col_defs, values):
                ctk.CTkLabel(
                    row_frame, text=val, width=col_width, anchor="w",
                    text_color=C["text_primary"], font=ctk.CTkFont(size=12),
                ).pack(side="left", padx=6)

        # Grand total row
        total_row = ctk.CTkFrame(section1, fg_color=C["accent_blue"], height=38, corner_radius=6)
        total_row.pack(fill="x", padx=12, pady=(4, 12))
        total_row.pack_propagate(False)
        total_values = [
            "", "GRAND TOTAL", "", "",
            str(self.strict_totals.get("Total Slots", 0)),
            f"{self.strict_totals.get('Total Amount (EGP)', 0):,.0f}",
        ]
        for (_, col_width), val in zip(col_defs, total_values):
            ctk.CTkLabel(
                total_row, text=val, width=col_width, anchor="w",
                text_color="#ffffff", font=ctk.CTkFont(size=13, weight="bold"),
            ).pack(side="left", padx=6)

        # ── Section 2: Conflict Alerts ────────────────────────────────────
        section2 = ctk.CTkFrame(body, fg_color=C["bg_card"], corner_radius=10)
        section2.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(
            section2, text="🔍 Conflict Detection",
            font=ctk.CTkFont(size=16, weight="bold"), text_color=C["text_primary"],
        ).pack(anchor="w", padx=16, pady=(14, 8))

        coordinator_conflicts = self.strict_conflicts.get("coordinator_conflicts", [])
        room_conflicts_data = self.strict_conflicts.get("room_conflicts", [])

        if not coordinator_conflicts and not room_conflicts_data:
            no_conflict_frame = ctk.CTkFrame(section2, fg_color="#0d3320", corner_radius=8)
            no_conflict_frame.pack(fill="x", padx=14, pady=(0, 14))
            ctk.CTkLabel(
                no_conflict_frame, text="✅ No conflicts detected — all assignments are clean!",
                text_color=C["status_green"], font=ctk.CTkFont(size=14, weight="bold"),
            ).pack(padx=16, pady=14)
        else:
            # Coordinator Conflicts
            if coordinator_conflicts:
                cc_header = ctk.CTkFrame(section2, fg_color="#3b1010", corner_radius=8)
                cc_header.pack(fill="x", padx=14, pady=(0, 4))
                ctk.CTkLabel(
                    cc_header,
                    text=f"🔴 Coordinator Conflicts ({len(coordinator_conflicts)})",
                    text_color=C["danger"], font=ctk.CTkFont(size=14, weight="bold"),
                ).pack(anchor="w", padx=14, pady=10)

                for conflict in coordinator_conflicts:
                    centers = ", ".join(conflict.get("Centers", []))
                    alert_frame = ctk.CTkFrame(section2, fg_color=C["bg_input"], corner_radius=6)
                    alert_frame.pack(fill="x", padx=20, pady=2)
                    ctk.CTkLabel(
                        alert_frame,
                        text=(
                            f"⚠️  {conflict.get('Coordinator', '')} is assigned to centers "
                            f"[{centers}] on {conflict.get('Day', '')} at {conflict.get('Time', '')}"
                        ),
                        text_color=C["status_yellow"], font=ctk.CTkFont(size=12),
                        wraplength=950, justify="left",
                    ).pack(anchor="w", padx=12, pady=8)

            # Room Conflicts
            if room_conflicts_data:
                spacer_pad = (12, 4) if coordinator_conflicts else (0, 4)
                rc_header = ctk.CTkFrame(section2, fg_color="#3b2e10", corner_radius=8)
                rc_header.pack(fill="x", padx=14, pady=spacer_pad)
                ctk.CTkLabel(
                    rc_header,
                    text=f"🟡 Room Conflicts ({len(room_conflicts_data)})",
                    text_color=C["status_yellow"], font=ctk.CTkFont(size=14, weight="bold"),
                ).pack(anchor="w", padx=14, pady=10)

                for conflict in room_conflicts_data:
                    groups = ", ".join(conflict.get("Groups", []))
                    alert_frame = ctk.CTkFrame(section2, fg_color=C["bg_input"], corner_radius=6)
                    alert_frame.pack(fill="x", padx=20, pady=2)
                    ctk.CTkLabel(
                        alert_frame,
                        text=(
                            f"⚠️  Room {conflict.get('Room Code', '')} in Center "
                            f"{conflict.get('Center Code', '')} has groups [{groups}] on "
                            f"{conflict.get('Day', '')} at {conflict.get('Time', '')}"
                        ),
                        text_color=C["status_yellow"], font=ctk.CTkFont(size=12),
                        wraplength=950, justify="left",
                    ).pack(anchor="w", padx=12, pady=8)

            # Bottom padding
            ctk.CTkFrame(section2, fg_color="transparent", height=8).pack()

    def _show_log_dialog(self):
        if not self.log_rows:
            messagebox.showinfo("WhatsApp Log", "No send/skip log entries yet.", parent=self)
            return
        headers = ["Timestamp", "Name", "Phone", "Status", "Error"]
        lines = [" | ".join(headers), "-" * 100]
        for row in self.log_rows:
            lines.append(" | ".join(value_to_text(row.get(header, "")) for header in headers))
        _show_text_dialog(self, "WhatsApp Sending Log", "\n".join(lines))

    def _show_upload_summary_dialog(self):
        if not self.upload_summary_text:
            messagebox.showinfo("Upload Summary", "Upload files first to generate a summary.", parent=self)
            return
        _show_text_dialog(self, "Upload Summary", self.upload_summary_text)

    def _show_upload_summary_if_needed(self):
        if not self.upload_summary_text:
            return
        unique_role_count, message_count, skipped_count = self._summary_counts()
        if unique_role_count != message_count or skipped_count:
            self._show_upload_summary_dialog()

    def _summary_counts(self):
        unique_role_people = set()
        for row in self.rows:
            for role, name_col in (("Main", "main_coordinator"), ("Backup", "backup_coordinator")):
                name = row.get(name_col, "").strip()
                if name:
                    unique_role_people.add((role, _person_key(name)))
        skipped_names = {_person_key(row.get("Name", "")) for row in self.log_rows if row.get("Status") == "skipped" and row.get("Name")}
        return len(unique_role_people), len(self.messages), len(skipped_names)

    def _build_upload_summary(self, skipped_logs):
        role_people = {}
        person_roles = defaultdict(set)
        duplicate_role_counts = Counter()

        for row in self.rows:
            for role, name_col in (("Main", "main_coordinator"), ("Backup", "backup_coordinator")):
                name = row.get(name_col, "").strip()
                if not name:
                    continue
                key = _person_key(name)
                role_people[(role, key)] = name
                person_roles[key].add(role)
                duplicate_role_counts[(role, key)] += 1

        generated_names = {_person_key(msg.get("Name", "")): msg.get("Name", "") for msg in self.messages}
        skipped_by_name = {}
        for row in skipped_logs:
            key = _person_key(row.get("Name", ""))
            if key:
                skipped_by_name[key] = row

        both_role_names = [
            role_people.get(("Main", key)) or role_people.get(("Backup", key)) or generated_names.get(key, key)
            for key, roles in person_roles.items()
            if len(roles) > 1
        ]
        repeated_role_names = [
            f"{role}: {role_people[(role, key)]} ({count} rows)"
            for (role, key), count in duplicate_role_counts.items()
            if count > 1
        ]
        missing_generated = [
            role_people[(role, key)]
            for (role, key) in role_people
            if key not in generated_names
        ]

        lines = [
            f"Unique Main/Backup role names in sessions: {len(role_people)}",
            f"Unique people in sessions: {len(person_roles)}",
            f"Generated message recipients: {len(self.messages)}",
            f"Skipped/missing phone recipients: {len(skipped_by_name)}",
            "",
        ]

        if both_role_names:
            lines.append("People appearing as BOTH Main and Backup, so they get ONE combined message:")
            lines.extend(f"- {name}" for name in sorted(set(both_role_names)))
            lines.append("")

        if repeated_role_names:
            lines.append("Repeated role entries across session rows:")
            lines.extend(f"- {item}" for item in sorted(repeated_role_names))
            lines.append("")

        if missing_generated:
            lines.append("Names found in sessions but not generated as recipients:")
            for name in sorted(set(missing_generated)):
                skipped = skipped_by_name.get(_person_key(name), {})
                reason = skipped.get("Error", "No matching valid phone found")
                phone = skipped.get("Phone", "")
                lines.append(f"- {name} | phone: {phone or 'missing'} | {reason}")
            lines.append("")

        if skipped_by_name:
            lines.append("Skipped details:")
            for row in skipped_by_name.values():
                lines.append(f"- {row.get('Name', '')} | {row.get('Phone', '') or 'missing'} | {row.get('Error', '')}")

        return "\n".join(lines).strip()

    def _refresh_stats(self):
        total = len(self.messages)
        selected = sum(1 for msg in self.messages if msg.get("Selected"))
        sent = sum(1 for msg in self.messages if msg.get("Status") == COORD_STATUS_SENT)
        failed = sum(1 for msg in self.messages if msg.get("Status") == COORD_STATUS_FAILED)
        skipped = sum(1 for msg in self.messages if msg.get("Status") == COORD_STATUS_SKIPPED)
        self.stats_label.configure(text=f"Recipients: {total} | Selected: {selected} | Sent: {sent} | Failed: {failed} | Skipped: {skipped}")

    def _get_int(self, variable, default):
        text = variable.get().strip()
        if not text:
            return default
        value = int(text)
        if value < 0:
            raise ValueError("Numeric settings cannot be negative.")
        return value

    def _get_float(self, variable, default):
        text = variable.get().strip()
        if not text:
            return default
        value = float(text)
        if value < 0:
            raise ValueError("Slot price cannot be negative.")
        return value

    # ── Progress panel helpers ────────────────────────────────────────────────

    def _show_progress_panel(self, total: int):
        """Reveal the progress panel and reset all widgets."""
        self.progress_bar.set(0)
        self.progress_counters_label.configure(text=f"0 of {total} messages processed")
        self.progress_current_label.configure(text="Preparing\u2026")
        self.progress_countdown_label.configure(text="")
        self.progress_eta_label.configure(text="")
        self.progress_panel.pack(fill="x", padx=20, pady=(0, 8), before=self.header_frame)

    def _hide_progress_panel(self):
        """Collapse the progress panel."""
        self.progress_panel.pack_forget()

    def _update_send_progress(
        self, total: int, processed: int, sent_ok: int, sent_failed: int,
        sent_skipped: int, current_name: str = "", countdown: int = 0, eta_secs: int = 0,
    ):
        """Refresh all progress-panel widgets. Called from the UI thread via after()."""
        if total > 0:
            self.progress_bar.set(processed / total)
        self.progress_counters_label.configure(
            text=f"\u2713 Sent: {sent_ok}  \u2717 Failed: {sent_failed}  \u2298 Skipped: {sent_skipped}  /  {processed} of {total}"
        )
        if current_name:
            if countdown > 0:
                self.progress_current_label.configure(text=f"Last sent: {current_name}")
            else:
                self.progress_current_label.configure(text=f"\U0001f4e4 Sending to: {current_name}")
        if countdown > 0:
            self.progress_countdown_label.configure(text=f"\u23f3 Next message in: {countdown}s")
            if eta_secs > 0:
                mins, secs = divmod(eta_secs, 60)
                eta_text = f"{mins}m {secs}s" if mins else f"{secs}s"
                self.progress_eta_label.configure(text=f"ETA: \u2248{eta_text}")
            else:
                self.progress_eta_label.configure(text="")
        else:
            self.progress_countdown_label.configure(text="\U0001f4e4 Sending\u2026")
            self.progress_eta_label.configure(text="")

    def _scroll_to_row(self, idx: int):
        """Try to scroll the table so the active row is visible."""
        widgets = self.row_widgets.get(idx)
        if not widgets:
            return
        frame = widgets.get("frame")
        if not frame:
            return
        try:
            frame.update_idletasks()
            canvas = self.table_frame._parent_canvas
            total_h = self.table_frame.winfo_height()
            y = frame.winfo_y()
            if total_h > 0:
                canvas.yview_moveto(max(0.0, (y - 40) / total_h))
        except Exception:
            pass

    # ── New UX helpers ────────────────────────────────────────────────────────

    def _send_single(self, idx: int):
        """Send a single row without requiring it to be checked."""
        if self._sending:
            messagebox.showinfo("Busy", "A batch is already running. Stop it first.", parent=self)
            return
        msg = self.messages[idx]
        name = msg.get("Name", "")
        if msg.get("Status") == COORD_STATUS_SENT:
            if not messagebox.askyesno("Already Sent", f"{name} was already sent. Send again?", parent=self):
                return
        if not messagebox.askyesno("Send Single", f"Send message to {name}?", parent=self):
            return
        self._send_targets([idx])

    def _toggle_select_all(self):
        """Toggle all visible rows to match the header checkbox state."""
        checked = self.select_all_var.get()
        query = self.search_var.get().strip().casefold()
        for i, msg in enumerate(self.messages):
            visible = (
                not query
                or query in msg.get("Name", "").casefold()
                or query in msg.get("Phone", "").casefold()
                or query in msg.get("Role", "").casefold()
            )
            if visible:
                self.messages[i]["Selected"] = checked
                widgets = self.row_widgets.get(i)
                if widgets:
                    widgets["selected"].set(checked)
        self._refresh_stats()

    def _update_status_color(self, level: str = "info"):
        """Switch the status-label text colour to reflect severity."""
        colors = {
            "info": C["accent_cyan"],
            "warn": C["status_yellow"],
            "success": C["status_green"],
            "error": C["danger"],
        }
        self.status_label.configure(text_color=colors.get(level, C["accent_cyan"]))

    def _schedule_search(self):
        """Debounce search: wait 200 ms after the last keystroke before re-rendering."""
        if self._search_after_id:
            self.after_cancel(self._search_after_id)
        self._search_after_id = self.after(200, self._render_table)

    def _export_messages_quick(self):
        """Export messages + analysis + log xlsx files into a user-selected folder."""
        if not self.messages:
            messagebox.showinfo("No Messages", "Generate messages first by uploading files.", parent=self)
            return
        
        out_dir = filedialog.askdirectory(title="Select Folder to Save Exports", parent=self)
        if not out_dir:
            return

        try:
            paths = save_outputs(self.messages, self.analysis, self.log_rows, out_dir)
            self._update_status_color("success")
            self.status_var.set(f"Exported \u2192 {os.path.basename(out_dir)}/")
        except Exception as exc:
            self._update_status_color("error")
            self.status_var.set(f"Export failed: {exc}")

    def _export_log_quick(self):
        """Export just the sending log xlsx file directly."""
        if not self.log_rows:
            messagebox.showinfo("No Log", "No log entries yet. Send some messages first.", parent=self)
            return
        try:
            paths = save_outputs(self.messages, self.analysis, self.log_rows, self.output_dir)
            self._update_status_color("success")
            self.status_var.set(f"Log exported \u2192 {os.path.basename(paths['log'])}")
        except Exception as exc:
            self._update_status_color("error")
            self.status_var.set(f"Log export failed: {exc}")

    def _show_report_export_dialog(self):
        if not self.strict_analysis:
            messagebox.showinfo("No Data", "No analysis data to export. Please upload files first.", parent=self)
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title("Export Analysis Report")
        dlg.geometry("380x330")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        dlg.grab_set()

        content = ctk.CTkFrame(dlg, fg_color=C["bg_card"], corner_radius=8)
        content.pack(fill="both", expand=True, padx=18, pady=18)

        ctk.CTkLabel(content, text="Export Options", font=ctk.CTkFont(size=18, weight="bold"), text_color=C["text_primary"]).pack(anchor="w", padx=14, pady=(14, 10))

        # Format
        format_var = tk.StringVar(value="PDF")
        ctk.CTkLabel(content, text="Format:", text_color=C["text_secondary"]).pack(anchor="w", padx=14, pady=(4, 2))
        f_frame = ctk.CTkFrame(content, fg_color="transparent")
        f_frame.pack(fill="x", padx=14, pady=(0, 10))
        ctk.CTkRadioButton(f_frame, text="PDF Document", variable=format_var, value="PDF", text_color=C["text_primary"]).pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(f_frame, text="Excel Workbook", variable=format_var, value="XLSX", text_color=C["text_primary"]).pack(side="left")

        # Detail Level
        detail_var = tk.StringVar(value="Detailed")
        ctk.CTkLabel(content, text="Detail Level:", text_color=C["text_secondary"]).pack(anchor="w", padx=14, pady=(4, 2))
        d_frame = ctk.CTkFrame(content, fg_color="transparent")
        d_frame.pack(fill="x", padx=14, pady=(0, 16))
        ctk.CTkRadioButton(d_frame, text="Detailed", variable=detail_var, value="Detailed", text_color=C["text_primary"]).pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(d_frame, text="Brief", variable=detail_var, value="Brief", text_color=C["text_primary"]).pack(side="left")

        def perform_export():
            from datetime import datetime
            fmt = format_var.get()
            is_detailed = detail_var.get() == "Detailed"
            
            ext = ".pdf" if fmt == "PDF" else ".xlsx"
            filetypes = [("PDF Document", "*.pdf")] if fmt == "PDF" else [("Excel Workbook", "*.xlsx")]
            
            filepath = filedialog.asksaveasfilename(
                title="Save Analysis Report",
                defaultextension=ext,
                initialfile=f"Analysis_Report_{datetime.now().strftime('%Y%m%d')}{ext}",
                filetypes=filetypes,
                parent=dlg
            )
            if not filepath:
                return

            try:
                if fmt == "PDF":
                    generate_pdf_report(self.strict_analysis, self.strict_totals, self.strict_conflicts, is_detailed, filepath)
                else:
                    generate_xlsx_report(self.strict_analysis, self.strict_totals, self.strict_conflicts, is_detailed, filepath)
                messagebox.showinfo("Export Success", f"Report saved successfully to:\n{filepath}", parent=dlg)
                dlg.destroy()
            except Exception as exc:
                messagebox.showerror("Export Failed", f"Failed to generate report:\n{exc}", parent=dlg)

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.pack(fill="x", padx=14, pady=(10, 14))
        StyledButton(actions, text="Cancel", width=90, color=C["bg_input"], hover=C["hover"], command=dlg.destroy).pack(side="left")
        StyledButton(actions, text="Export", width=120, color=C["accent_blue"], hover=C["accent_purple"], command=perform_export).pack(side="right")
