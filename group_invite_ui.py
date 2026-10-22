"""
Group Invite UI - GroupInviteApp window for sending WhatsApp group invite messages.
Redesigned to match the CoordinatorApp dashboard layout.
"""
from __future__ import annotations
import os
import threading
import time
import random
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter.scrolledtext import ScrolledText
import customtkinter as ctk
import datetime

from app_config import COLORS as C
from engine import MessageEngine
from ui_components import StatusDot, StyledButton, StyledEntry, StyledTextbox, StyledOptionMenu, CardFrame, play_sound
from file_errors import FileUploadError

from group_invite_logic import (
    load_group_invite_templates, save_group_invite_templates,
    parse_sheet, build_all_messages, send_batch,
    load_progress, clear_progress, save_log, DEFAULT_GROUP_INVITE_TEMPLATES
)

STATUS_READY = "Ready"
STATUS_SENDING = "Sending"
STATUS_SENT = "Sent"
STATUS_FAILED = "Failed"
STATUS_SKIPPED = "Skipped"
STATUS_COLORS = {
    STATUS_READY: C["status_yellow"],
    STATUS_SENDING: C["accent_cyan"],
    STATUS_SENT: C["status_green"],
    STATUS_FAILED: C["danger"],
    STATUS_SKIPPED: C["text_muted"],
}

class GroupInviteApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PM Automation - Group Invite Sender")
        self.geometry("1280x780")
        self.minsize(1100, 650)
        self.configure(fg_color=C["bg_dark"])
        self.after(0, lambda: self.state("zoomed"))

        self.engine = MessageEngine()
        
        self.sheet_file = ""
        self.rows = []
        self.messages = []
        self.log_rows = []
        
        self.row_widgets = {}
        self._sending = False
        self._stop_event = threading.Event()
        self.select_all_var = tk.BooleanVar(value=False)
        self.templates = load_group_invite_templates()

        self.search_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Upload the students sheet to begin.")
        self.tpl_var = tk.StringVar(value=next(iter(self.templates), "1. Coaching Welcome (Arabic)"))

        # Settings variables
        self.f_coordinator = tk.StringVar(value="أسامة الفقي")
        self.f_program = tk.StringVar(value="NTI - Creative (i-lance)")
        self.f_stage = tk.StringVar(value="Coaching")
        self.f_start_day = tk.StringVar(value="الأحد")
        self.f_start_date = tk.StringVar(value="4 أكتوبر 2026")
        self.f_start_time = tk.StringVar(value="6 صباحاً")
        self.f_end_time = tk.StringVar(value="9 صباحاً")
        self.f_duration = tk.StringVar(value="10")
        self.f_link = tk.StringVar(value="https://chat.whatsapp.com/")

        self.delay_min_var = tk.StringVar(value="45")
        self.delay_max_var = tk.StringVar(value="90")
        self.daily_limit_var = tk.StringVar(value="25")

        self._build_layout()
        self._load_template_text()
        self._render_table()

    def _build_layout(self):
        # Outer sidebar shell
        sidebar_shell = ctk.CTkFrame(self, width=330, fg_color=C["bg_sidebar"], corner_radius=0)
        sidebar_shell.pack(side="left", fill="y")
        sidebar_shell.pack_propagate(False)

        # Scrollable inner container
        self.sidebar = ctk.CTkScrollableFrame(
            sidebar_shell,
            fg_color=C["bg_sidebar"],
            scrollbar_button_color=C["bg_card"],
            scrollbar_button_hover_color=C["hover"],
            corner_radius=0,
            width=308,
        )
        self.sidebar.pack(fill="both", expand=True)

        ctk.CTkLabel(
            self.sidebar,
            text="Group Invites",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=C["text_primary"],
            justify="left",
        ).pack(anchor="w", padx=24, pady=(26, 8))

        ctk.CTkLabel(
            self.sidebar,
            text="Upload sheet, review, edit, send, and track group invites.",
            font=ctk.CTkFont(size=12),
            text_color=C["text_secondary"],
            wraplength=250,
            justify="left",
        ).pack(anchor="w", padx=24, pady=(0, 18))

        StyledButton(
            self.sidebar,
            text="Upload Sheet",
            width=250,
            color=C["accent_blue"],
            hover=C["accent_purple"],
            command=self._upload_files,
        ).pack(anchor="w", padx=24, pady=(0, 10))

        self._file_label = ctk.CTkLabel(
            self.sidebar,
            text="No sheet loaded.",
            font=ctk.CTkFont(size=11),
            text_color=C["text_muted"],
            wraplength=250,
            justify="left",
        )
        self._file_label.pack(anchor="w", padx=24, pady=(0, 18))

        # Templates Card
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
        self.tpl_editor.bind("<KeyRelease>", lambda _e: self._update_preview())

        placeholder_row = ctk.CTkFrame(tpl_frame, fg_color="transparent")
        placeholder_row.pack(fill="x", padx=14, pady=(0, 8))
        for placeholder in ("{name}", "{program}", "{stage}", "{link}"):
            StyledButton(
                placeholder_row,
                text=placeholder,
                width=50,
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
        StyledButton(tpl_btns2, text="Regenerate", width=104, color=C["accent_cyan"], hover=C["accent_blue"], command=self._regenerate_messages).pack(side="left", padx=2)
        self.template_hint = ctk.CTkLabel(
            tpl_frame,
            text="Placeholders: {name}, {program}, {stage}, {link}, {coordinator_name}, {start_day}, {start_date}, {start_time}, {end_time}, {duration_weeks}",
            text_color=C["text_muted"],
            font=ctk.CTkFont(size=10),
            wraplength=245,
            justify="left",
        )
        self.template_hint.pack(anchor="w", padx=14, pady=(0, 12))

        # Sending Buttons
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
        self.stop_btn.pack(anchor="w", padx=24, pady=(0, 10))

        self.clear_progress_btn = StyledButton(
            self.sidebar,
            text="🗑 Clear Progress",
            width=250,
            color=C["bg_input"],
            hover=C["hover"],
            command=self._clear_progress,
        )
        self.clear_progress_btn.pack(anchor="w", padx=24, pady=(0, 18))

        self.stats_label = ctk.CTkLabel(
            self.sidebar,
            text="Recipients: 0 | Sent: 0 | Failed: 0",
            text_color=C["text_secondary"],
            font=ctk.CTkFont(size=12),
            wraplength=250,
        )
        self.stats_label.pack(anchor="w", padx=24, pady=(0, 12))

        # Main Area
        self.main_area = ctk.CTkFrame(self, fg_color=C["bg_dark"], corner_radius=0)
        self.main_area.pack(side="right", fill="both", expand=True)

        top = ctk.CTkFrame(self.main_area, fg_color=C["bg_dark"], height=74)
        top.pack(fill="x", padx=20, pady=(16, 8))
        top.pack_propagate(False)

        ctk.CTkLabel(
            top,
            text="Group Invites Dashboard",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=C["text_primary"],
        ).pack(side="left", padx=(0, 20))

        search = StyledEntry(
            top,
            textvariable=self.search_var,
            placeholder_text="Search name or phone",
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
        self.search_var.trace_add("write", lambda *_: self._filter_table())

        toolbar = ctk.CTkFrame(self.main_area, fg_color=C["bg_card"], height=48, corner_radius=8)
        toolbar.pack(fill="x", padx=20, pady=(0, 10))
        toolbar.pack_propagate(False)
        StyledButton(
            toolbar,
            text="Upload Sheet",
            width=118,
            color=C["accent_blue"],
            hover=C["accent_purple"],
            command=self._upload_files,
        ).pack(side="left", padx=(10, 6), pady=6)
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
            text="Regenerate",
            width=112,
            color=C["accent_purple"],
            hover=C["accent_blue"],
            command=self._regenerate_messages,
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

        # Progress panel
        self.progress_panel = ctk.CTkFrame(self.main_area, fg_color=C["bg_card"], corner_radius=8)
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

        self.tabview = ctk.CTkTabview(self.main_area, fg_color=C["bg_dark"])
        self.tabview.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.tab_messages = self.tabview.add("Messages")

        self.header_frame = ctk.CTkFrame(self.tab_messages, fg_color=C["bg_card"], height=38, corner_radius=0)
        self.header_frame.pack(fill="x", padx=0)
        self.header_frame.pack_propagate(False)
        for text, width in [
            ("#", 40), ("Name", 190), ("Phone", 150), ("Status", 100), ("Actions", 240),
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

    def _open_settings_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Group Invites Settings")
        dlg.geometry("500x600")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        dlg.grab_set()

        content = ctk.CTkScrollableFrame(dlg, fg_color=C["bg_card"], corner_radius=8)
        content.pack(fill="both", expand=True, padx=18, pady=18)

        ctk.CTkLabel(
            content,
            text="Group Invites Settings",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=C["text_primary"],
        ).pack(anchor="w", padx=18, pady=(18, 12))

        ctk.CTkLabel(content, text="Template Variables", font=ctk.CTkFont(weight="bold"), text_color=C["text_primary"]).pack(anchor="w", padx=18, pady=(10,5))
        
        settings_vars = [
            ("Coordinator Name", self.f_coordinator),
            ("Program", self.f_program),
            ("Stage", self.f_stage),
            ("Start Day", self.f_start_day),
            ("Start Date", self.f_start_date),
            ("Start Time", self.f_start_time),
            ("End Time", self.f_end_time),
            ("Weeks", self.f_duration),
            ("Group Link", self.f_link),
        ]
        for label, variable in settings_vars:
            row = ctk.CTkFrame(content, fg_color="transparent")
            row.pack(fill="x", padx=18, pady=4)
            ctk.CTkLabel(row, text=label, width=130, anchor="w", text_color=C["text_secondary"]).pack(side="left")
            StyledEntry(row, textvariable=variable, width=280, fg_color=C["bg_input"]).pack(side="right")

        ctk.CTkLabel(content, text="Anti-Ban Settings", font=ctk.CTkFont(weight="bold"), text_color=C["text_primary"]).pack(anchor="w", padx=18, pady=(20,5))

        settings_antiban = [
            ("Delay min (sec)", self.delay_min_var),
            ("Delay max (sec)", self.delay_max_var),
            ("Daily limit", self.daily_limit_var),
        ]
        for label, variable in settings_antiban:
            row = ctk.CTkFrame(content, fg_color="transparent")
            row.pack(fill="x", padx=18, pady=4)
            ctk.CTkLabel(row, text=label, width=130, anchor="w", text_color=C["text_secondary"]).pack(side="left")
            StyledEntry(row, textvariable=variable, width=280, fg_color=C["bg_input"]).pack(side="right")

        def apply_settings(regenerate=False):
            try:
                int(self.delay_min_var.get())
                int(self.delay_max_var.get())
                int(self.daily_limit_var.get())
            except ValueError:
                messagebox.showerror("Invalid Settings", "Delays and limit must be integers.", parent=dlg)
                return
            self.status_var.set("Settings updated.")
            if regenerate and self.rows:
                self._regenerate_messages()
            else:
                self._update_preview()
            dlg.destroy()

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.pack(fill="x", padx=18, pady=(12, 18))
        StyledButton(actions, text="Close", width=95, color=C["bg_input"], hover=C["hover"], command=dlg.destroy).pack(side="left")
        StyledButton(actions, text="Apply + Regenerate", width=155, color=C["accent_cyan"], hover=C["accent_blue"], command=lambda: apply_settings(True)).pack(side="right")
        StyledButton(actions, text="Apply", width=95, color=C["accent_blue"], hover=C["accent_purple"], command=lambda: apply_settings(False)).pack(side="right", padx=(8, 8))

    def _load_template_text(self):
        if not hasattr(self, "tpl_editor"):
            return
        name = self.tpl_var.get()
        self.tpl_editor.delete("1.0", "end")
        self.tpl_editor.insert("1.0", self.templates.get(name, ""))
        self._update_preview()

    def _get_template_text(self):
        text = self.tpl_editor.get("1.0", "end").strip()
        try:
            # Fix surrogate pairs generated by tkinter for emojis
            return text.encode('utf-16', 'surrogatepass').decode('utf-16')
        except Exception:
            return text

    def _save_current_template(self):
        name = self.tpl_var.get().strip()
        if not name:
            return
        self.templates[name] = self._get_template_text()
        save_group_invite_templates(self.templates)
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
        save_group_invite_templates(self.templates)
        self._refresh_template_menu(name)
        self._load_template_text()

    def _delete_template(self):
        name = self.tpl_var.get()
        if len(self.templates) <= 1:
            messagebox.showwarning("Cannot Delete", "Keep at least one template.", parent=self)
            return
        if not messagebox.askyesno("Delete Template", f"Delete template '{name}'?", parent=self):
            return
        self.templates.pop(name, None)
        save_group_invite_templates(self.templates)
        self._refresh_template_menu(next(iter(self.templates)))
        self._load_template_text()

    def _reset_templates(self):
        if not messagebox.askyesno("Reset Templates", "Reset templates to defaults?", parent=self):
            return
        self.templates = dict(DEFAULT_GROUP_INVITE_TEMPLATES)
        save_group_invite_templates(self.templates)
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
        self._update_preview()

    def _update_preview(self, event=None):
        if not hasattr(self, "preview_box"):
            return
        template = self._get_template_text()
        ctx = self._get_global_ctx()
        
        sample_name = "منى"
        sample_phone = "+201012345678"
        if getattr(self, "last_previewed_msg", None):
            sample_name = self.last_previewed_msg.get("name", sample_name)
            sample_phone = self.last_previewed_msg.get("phone", sample_phone)
        elif getattr(self, "rows", None) and len(self.rows) > 0:
            sample_student = self.rows[0]
            sample_name = sample_student.get("name", sample_name)
            sample_phone = sample_student.get("phone", sample_phone)

        for k in ("name", "Arabic Name", "arabic_name", "arabic name", "first_name", "الاسم", "الاسم بالعربي"):
            ctx[k] = sample_name
        for k in ("phone", "Phone", "mobile", "Mobile", "raw_phone"):
            ctx[k] = sample_phone

        try:
            from template_utils import render_template
            text = render_template(template, ctx)
        except Exception:
            text = template

        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", text if text.strip() else "Type a template message...")
        self.preview_box.configure(state="disabled")

    def _get_global_ctx(self):
        return {
            "coordinator_name": self.f_coordinator.get().strip(),
            "program":          self.f_program.get().strip(),
            "stage":            self.f_stage.get().strip(),
            "start_day":        self.f_start_day.get().strip(),
            "start_date":       self.f_start_date.get().strip(),
            "start_time":       self.f_start_time.get().strip(),
            "end_time":         self.f_end_time.get().strip(),
            "duration_weeks":   self.f_duration.get().strip(),
            "link":             self.f_link.get().strip(),
        }

    def _upload_files(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Select Student Sheet",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return
        rows, errors = parse_sheet(path)
        if errors:
            messagebox.showerror("Sheet Error", "\n".join(errors), parent=self)
            return
        self.rows = rows
        self.sheet_file = os.path.basename(path)
        self._file_label.configure(text=f"{self.sheet_file} ({len(rows)} students)", text_color=C["text_primary"])
        self._regenerate_messages()

    def _regenerate_messages(self):
        if not self.rows:
            messagebox.showwarning("No Sheet", "Please load a student sheet first.", parent=self)
            return
        template = self._get_template_text()
        if not template:
            messagebox.showwarning("No Template", "Template is empty.", parent=self)
            return
        
        self._save_current_template()
        ctx = self._get_global_ctx()
        self.messages = build_all_messages(self.rows, template, ctx)
        self._render_table()
        self.status_var.set(f"Preview built for {len(self.messages)} messages.")
        self._update_stats()
        if self.messages:
            self._preview_message(self.messages[0])

    def _filter_table(self):
        self._render_table()

    def _render_table(self):
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        search_query = self.search_var.get().casefold()
        progress = load_progress()
        self.row_widgets.clear()

        for i, msg in enumerate(self.messages):
            name = msg["name"]
            phone = msg["phone"]
            if search_query and search_query not in name.casefold() and search_query not in phone.casefold():
                continue

            p_info = progress.get(phone, {})
            status = p_info.get("status", "Pending")

            row_f = ctk.CTkFrame(self.table_frame, fg_color=C["bg_card"] if i % 2 == 0 else C["bg_input"], corner_radius=4)
            row_f.pack(fill="x", padx=4, pady=2)
            
            ctk.CTkLabel(row_f, text=str(i + 1), width=40, anchor="w", text_color=C["text_muted"]).pack(side="left", padx=3)
            ctk.CTkLabel(row_f, text=name, width=190, anchor="w", text_color=C["text_primary"], font=ctk.CTkFont(size=12)).pack(side="left", padx=3)
            ctk.CTkLabel(row_f, text=phone or "(invalid)", width=150, anchor="w", text_color=C["text_secondary"]).pack(side="left", padx=3)
            
            status_lbl = ctk.CTkLabel(row_f, text=status.capitalize(), width=100, anchor="w", text_color=STATUS_COLORS.get(status.capitalize(), C["text_muted"]), font=ctk.CTkFont(size=12, weight="bold"))
            status_lbl.pack(side="left", padx=3)
            self.row_widgets[i] = status_lbl

            actions_f = ctk.CTkFrame(row_f, fg_color="transparent", width=240)
            actions_f.pack(side="left", padx=3, fill="x", expand=True)
            StyledButton(actions_f, text="Preview", width=55, color=C["bg_input"], hover=C["hover"], command=lambda m=msg: self._preview_message(m)).pack(side="left", padx=2)
            StyledButton(actions_f, text="Edit", width=50, color=C["accent_purple"], hover=C["accent_blue"], command=lambda m=msg: self._edit_message(m)).pack(side="left", padx=2)
            StyledButton(actions_f, text="Send", width=50, color=C["accent_cyan"], command=lambda m=msg: self._start_sending([m])).pack(side="left", padx=2)
            StyledButton(actions_f, text="Delete", width=55, color=C["danger"], hover="#991b1b", command=lambda m=msg: self._delete_row(m)).pack(side="left", padx=2)

    def _delete_row(self, msg):
        if msg in self.messages:
            self.messages.remove(msg)
        for r in self.rows:
            if r.get("phone") == msg.get("phone") and r.get("name") == msg.get("name"):
                self.rows.remove(r)
                break
        self._render_table()
        self._update_stats()

    def _edit_message(self, msg):
        dlg = ctk.CTkToplevel(self)
        dlg.title(f"Edit Message - {msg['name']}")
        dlg.geometry("500x400")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        dlg.grab_set()

        ctk.CTkLabel(dlg, text=f"Editing Message for {msg['name']}", font=ctk.CTkFont(size=16, weight="bold"), text_color=C["text_primary"]).pack(pady=10)

        textbox = StyledTextbox(dlg, fg_color=C["bg_input"], wrap="word")
        textbox.pack(fill="both", expand=True, padx=20, pady=10)
        textbox.insert("1.0", msg["message"])

        def save_edit():
            msg["message"] = textbox.get("1.0", "end").strip()
            self._preview_message(msg)
            dlg.destroy()

        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)
        StyledButton(btn_frame, text="Cancel", width=100, color=C["bg_input"], hover=C["hover"], command=dlg.destroy).pack(side="left")
        StyledButton(btn_frame, text="Save", width=100, color=C["accent_cyan"], command=save_edit).pack(side="right")

    def _preview_message(self, msg):
        self.last_previewed_msg = msg
        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", msg["message"])
        self.preview_box.configure(state="disabled")

    def _update_stats(self):
        progress = load_progress()
        sent = 0
        failed = 0
        for msg in self.messages:
            p_info = progress.get(msg["phone"], {})
            st = p_info.get("status", "")
            if st == "sent":
                sent += 1
            elif st == "failed":
                failed += 1
        total = len(self.messages)
        self.stats_label.configure(text=f"Recipients: {total} | Sent: {sent} | Failed: {failed}")

    def _show_progress_panel(self):
        self.progress_panel.pack(fill="x", padx=20, pady=(0, 10), before=self.tabview)
        self.progress_bar.set(0)
        self.progress_current_label.configure(text="Starting...")
        self.progress_counters_label.configure(text="0 / 0")
        self.progress_countdown_label.configure(text="")

    def _hide_progress_panel(self):
        self.progress_panel.pack_forget()

    def _send_all(self):
        if not self.messages:
            messagebox.showwarning("No Messages", "Build a preview first.", parent=self)
            return
        if self._sending:
            return
        self._start_sending(self.messages)

    def _start_sending(self, messages_to_send):
        try:
            min_d = int(self.delay_min_var.get())
            max_d = int(self.delay_max_var.get())
            limit = int(self.daily_limit_var.get())
        except ValueError:
            messagebox.showerror("Invalid Settings", "Delays and limit must be integers.", parent=self)
            return

        self._sending = True
        self._stop_event.clear()
        self._show_progress_panel()
        self.status_var.set("Sending in progress...")
        self.send_all_btn.configure(state="disabled")

        total = len(messages_to_send)
        self.progress_counters_label.configure(text=f"0 / {total}")

        def _progress(idx, name, phone, ok, detail):
            self.after(0, lambda: self._on_progress(idx, name, phone, ok, detail, total))

        def _delay_tick(remaining, tot):
            self.after(0, lambda: self.progress_countdown_label.configure(text=f"Next in {remaining}s" if remaining > 0 else "Sending..."))

        def _complete():
            self.after(0, self._on_complete)

        send_batch(
            messages_to_send,
            min_delay=min_d,
            max_delay=max_d,
            daily_limit=limit,
            on_progress=_progress,
            on_delay=_delay_tick,
            on_complete=_complete,
            stop_event=self._stop_event,
        )

    def _on_progress(self, idx, name, phone, ok, detail, total):
        self.progress_bar.set((idx + 1) / total)
        self.progress_counters_label.configure(text=f"{idx + 1} / {total}")
        
        status_str = "Skipped" if detail == "skipped_already_sent" else ("Sent" if ok else "Failed")
        
        if detail == "daily_limit_reached":
            self.progress_current_label.configure(text=f"Daily limit reached.")
            status_str = "Skipped"
        else:
            self.progress_current_label.configure(text=f"{name} - {status_str}")

        if idx in self.row_widgets:
            self.row_widgets[idx].configure(text=status_str, text_color=STATUS_COLORS.get(status_str, C["text_muted"]))
        self._update_stats()
        
        # log 
        if ok is True or ok is False:
            self.log_rows.append({
                "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Name": name,
                "Phone": phone,
                "Status": "sent" if ok else "failed",
                "Error": "" if ok else detail
            })

    def _on_complete(self):
        self._sending = False
        self.send_all_btn.configure(state="normal")
        self.progress_countdown_label.configure(text="")
        self.status_var.set("Batch complete.")
        play_sound("success")
        self._hide_progress_panel()
        self._update_stats()

    def _stop_sending(self):
        if not self._sending: return
        self._stop_event.set()
        self._sending = False
        self.send_all_btn.configure(state="normal")
        self.progress_countdown_label.configure(text="Stopping...")
        self.status_var.set("Stopped by user.")

    def _clear_progress(self):
        if messagebox.askyesno("Clear Progress", "Clear all saved send progress? (Next send will re-send everyone.)", parent=self):
            clear_progress()
            self._render_table()
            self._update_stats()

    def _export_log_quick(self):
        if not self.log_rows:
            messagebox.showinfo("No Log", "No log data yet. Send some messages first.", parent=self)
            return
        path = save_log(self.log_rows)
        messagebox.showinfo("Log Saved", f"Log saved to:\n{path}", parent=self)
