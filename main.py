"""DECI Automation — EYouth coordination WhatsApp Messenger."""
import customtkinter as ctk
import openpyxl, os, random, datetime, threading, tkinter as tk, winsound
from tkinter import filedialog, messagebox
from app_config import COLORS as C, load_templates, save_templates
from ui_components import StatusDot, StyledButton, SidebarButton
from engine import MessageEngine
from schedule_manager import ScheduleMixin
from coordinator_ui import CoordinatorApp, choose_recipient_type

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

STATUS_NOT_SENT = "Not Sent"
STATUS_SENDING  = "Sending…"
STATUS_SENT     = "Sent"
STATUS_FAILED   = "Failed"
STATUS_ACK      = "Acknowledged"
STATUS_CONF     = "Confirmed"

STATUS_COLORS = {
    STATUS_NOT_SENT: C["status_red"],
    STATUS_SENDING:  C["status_yellow"],
    STATUS_SENT:     C["status_yellow"],
    STATUS_FAILED:   C["danger"],
    STATUS_ACK:      C["accent_cyan"],
    STATUS_CONF:     C["status_green"],
}

def normalize_phone(phone_str):
    if not phone_str:
        return ""
    p = str(phone_str).strip()
    
    # Remove common Excel formatting
    if p.endswith(".00"): p = p[:-3]
    elif p.endswith(".0"): p = p[:-2]
        
    # Keep only digits and '+'
    p = "".join(c for c in p if c.isdigit() or c == '+')
    
    if not p: return ""
    if p.startswith("+"): return p
    
    # Handle Egyptian numbers (10 or 11 digits starting with 01 or 1)
    if len(p) == 11 and p.startswith("0"):
        return "+20" + p[1:]
    if len(p) == 10 and p.startswith(("10", "11", "12", "15")):
        return "+20" + p
    
    # Handle numbers starting with 20 (country code without +)
    if p.startswith("20") and len(p) >= 11:
        return "+" + p
        
    # Handle 00 prefix
    if p.startswith("00"):
        return "+" + p[2:]
        
    # Default: Add +20 for Egypt if it looks like a local number
    if len(p) in (10, 11) and not p.startswith("+"):
        clean = p[1:] if p.startswith("0") else p
        return "+20" + clean

    # Generic fallback
    return "+" + p if not p.startswith("+") else p


class TourGuide(ctk.CTkToplevel):
    def __init__(self, parent, steps):
        super().__init__(parent)
        self.parent = parent
        self.steps = steps
        self.current_step = 0
        self._pulse_state = False
        self._pulse_job = None

        self.title("Assistant Guide")
        self.geometry("400x260")
        self.attributes("-topmost", True)
        self.configure(fg_color=C["bg_card"])
        self.overrideredirect(True)

        self.title_lbl = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=18, weight="bold"), text_color=C["accent_cyan"])
        self.title_lbl.pack(pady=(20, 5))

        self.msg_lbl = ctk.CTkLabel(self, text="", wraplength=350, font=ctk.CTkFont(size=13))
        self.msg_lbl.pack(pady=10, padx=20)

        self.btn_f = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_f.pack(side="bottom", fill="x", pady=20)

        self.skip_btn = StyledButton(self.btn_f, text="Skip", color=C["bg_input"],
            text_color=C["text_secondary"], hover=C["hover"], command=self.close_tour, width=70)
        self.skip_btn.pack(side="left", padx=20)

        self.step_lbl = ctk.CTkLabel(self.btn_f, text="", font=ctk.CTkFont(size=11), text_color=C["text_muted"])
        self.step_lbl.pack(side="left", expand=True)

        self.next_btn = StyledButton(self.btn_f, text="Next →", command=self.next_step, width=100)
        self.next_btn.pack(side="right", padx=20)

        self.show_step()

    def show_step(self):
        # Clean up previous step's highlight
        if self.current_step > 0:
            prev = self.steps[self.current_step - 1].get("widget")
            if prev:
                self._reset_highlight(prev)
        self._stop_pulse()

        step = self.steps[self.current_step]
        self.title_lbl.configure(text=step["title"])
        self.msg_lbl.configure(text=step["msg"])
        self.step_lbl.configure(text=f"Step {self.current_step + 1} of {len(self.steps)}")

        if self.current_step == len(self.steps) - 1:
            self.next_btn.configure(text="Finish 🎉", color=C["status_green"], hover="#15803d")

        target = step.get("widget")
        if target:
            targets = target if isinstance(target, list) else [target]
            for t in targets:
                t.configure(border_width=3, border_color=C["accent_cyan"])

            action = step.get("action")
            if action:
                action()

            # Start pulsing highlight
            self._start_pulse(targets)

            # Delay positioning until Tk has rendered the widgets
            self.after(60, lambda: self._position_near(targets[0]))

    def _start_pulse(self, targets):
        """Alternate the border color to draw attention."""
        self._pulse_state = not self._pulse_state
        color = C["accent_cyan"] if self._pulse_state else C["accent_purple"]
        for t in targets:
            if t.winfo_exists():
                t.configure(border_color=color)
        self._pulse_job = self.after(500, lambda: self._start_pulse(targets))

    def _stop_pulse(self):
        if self._pulse_job:
            self.after_cancel(self._pulse_job)
            self._pulse_job = None

    def _position_near(self, widget):
        """Smart positioning: sidebar widgets → place to the right; others → place below."""
        if not widget.winfo_exists():
            return

        wx = widget.winfo_rootx()
        wy = widget.winfo_rooty()
        ww = widget.winfo_width()
        wh = widget.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()

        # If widget is in the sidebar (left ~300px), place tour window to the right
        if wx < 350:
            x = wx + ww + 20
            # If that pushes off-screen, place it in the main panel area
            if x + 400 > sw:
                x = self.parent.winfo_x() + self.parent.winfo_width() - 440
            y = wy + (wh // 2) - 130
        else:
            # Place below the widget
            x = wx + (ww // 2) - 200
            y = wy + wh + 15

        # Clamp to screen bounds
        x = max(10, min(x, sw - 410))
        y = max(10, min(y, sh - 270))
        self.geometry(f"+{x}+{y}")

    def next_step(self):
        self.current_step += 1
        if self.current_step < len(self.steps):
            self.show_step()
        else:
            final = self.steps[-1].get("final_msg")
            if final:
                messagebox.showinfo("Tour Complete", final)
            self.close_tour()

    def close_tour(self):
        self._stop_pulse()
        idx = min(self.current_step, len(self.steps) - 1)
        target = self.steps[idx].get("widget")
        if target:
            self._reset_highlight(target)
        self.destroy()

    def _reset_highlight(self, widget):
        targets = widget if isinstance(widget, list) else [widget]
        for t in targets:
            if t.winfo_exists():
                t.configure(
                    border_width=1 if isinstance(t, ctk.CTkTextbox) else 0,
                    border_color=C["border"]
                )


class App(ctk.CTk, ScheduleMixin):
    def __init__(self):
        super().__init__()
        self.title("DECI Automation — EYouth coordination")
        self.geometry("1280x780")
        self.minsize(1100, 650)
        self.configure(fg_color=C["bg_dark"])

        self.templates = load_templates()
        self.engine = MessageEngine()
        self.all_data = self._load_all_data()
        
        for slot in self.all_data:
            for row in self.all_data[slot]:
                row["phone"] = normalize_phone(row.get("phone", ""))

        new_all_data = {}
        for slot_name, rows in self.all_data.items():
            for row in rows:
                if "day" not in row:
                    if " - " in slot_name:
                        parts = slot_name.split(" - ", 1)
                        row["day"] = parts[0]
                    else:
                        row["day"] = "Friday"
                row["day"] = row["day"].strip().title()
                
            if " - " in slot_name:
                parts = slot_name.split(" - ", 1)
                new_slot = f"{parts[0].strip().title()} - {parts[1]}"
            else:
                new_slot = slot_name.strip().title()
                if slot_name.lower() == "general":
                    new_slot = "General"
                    
            if new_slot not in new_all_data:
                new_all_data[new_slot] = []
            new_all_data[new_slot].extend(rows)
            
        self.all_data = new_all_data

        for s in ["9:00 AM", "2:00 PM", "5:00 PM", "General"]:
            if s in self.all_data and not self.all_data[s]:
                del self.all_data[s]

        self.day_filter_var = ctk.StringVar(value="All Days")
        self.slot_var = ctk.StringVar(value="All slots")

        self.row_widgets = []
        self.last_selected_index = None # Track selection for preview
        self._is_sending = False
        self._schedules = []
        self._schedule_timers = {}
        self._schedule_countdown_job = None

        self._build_sidebar()
        self._build_main()
        self._update_slot_menu()
        self._render_table()

        self._load_schedules()
        self._start_global_countdown()

    @property
    def current_slot(self):
        day_filter = self.day_filter_var.get()
        time_filter = self.slot_var.get()
        if day_filter == "All Days" and time_filter == "All slots":
            return "All slots"
        elif day_filter == "All Days":
            return f"All Days — {time_filter}"
        elif time_filter == "All slots":
            return f"{day_filter} — All slots"
        else:
            return f"{day_filter} — {time_filter}"

    def _populate_rows(self):
        day_filter = self.day_filter_var.get()
        time_filter = self.slot_var.get()
        
        self.rows = []
        for slot_key, slot_rows in self.all_data.items():
            # Check if this slot matches the day filter
            if day_filter != "All Days":
                if not slot_key.startswith(day_filter):
                    continue
            
            # Check if this slot matches the time filter
            if time_filter != "All slots":
                slot_time = ""
                if " - " in slot_key:
                    slot_time = slot_key.split(" - ", 1)[1]
                else:
                    if slot_key == "General":
                        slot_time = "General"
                
                if slot_time != time_filter:
                    continue
            
            # Add matching rows
            self.rows.extend(slot_rows)

    def _animate_logo(self):
        self._logo_anim_step = getattr(self, "_logo_anim_step", 0) + 1
        i = self._logo_anim_step % 4

        # Bright Neon Palette
        neon_cyan = "#00f6ff"
        neon_purple = "#b026ff"
        neon_pink = "#ff007f"
        neon_blue = "#0078ff"

        icon_colors = [neon_cyan, neon_purple, neon_pink, neon_blue]
        border_colors = [neon_blue, neon_cyan, neon_purple, neon_pink]

        try:
            self.logo_icon.configure(text_color=icon_colors[i])
            self.logo_title.configure(text_color=icon_colors[(i+1)%4])
            
            # Neon glowing border effect
            self.logo_frame.configure(border_color=border_colors[i], border_width=2)

            self.after(800, self._animate_logo)
        except Exception:
            pass

    # ── Sidebar ──────────────────────────────────────────────
    def _build_sidebar(self):
        sb = ctk.CTkScrollableFrame(self, width=280, fg_color=C["bg_sidebar"], corner_radius=0)
        sb.pack(side="left", fill="y")

        # Logo area - reduced padding
        self.logo_frame = ctk.CTkFrame(
            sb,
            fg_color=C["bg_card"],
            corner_radius=18,
            border_width=2,
            border_color=C["border"]
        )
        self.logo_frame.pack(fill="x", padx=14, pady=(10, 10))

        self.logo_inner = ctk.CTkFrame(self.logo_frame, fg_color="transparent")
        self.logo_inner.pack(fill="x", padx=12, pady=10)

        ctk.CTkFrame(self.logo_inner, height=2, fg_color=C["accent_blue"]).pack(fill="x", pady=(0, 8))

        self.logo_icon = ctk.CTkLabel(
            self.logo_inner,
            text="⚡",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color=C["accent_cyan"]
        )
        self.logo_icon.pack(pady=(2, 2))

        self.logo_title = ctk.CTkLabel(
            self.logo_inner,
            text="WhatsApp Automation",
            font=ctk.CTkFont(size=21, weight="bold")
        )
        self.logo_title.pack()

        self.logo_subtitle = ctk.CTkLabel(
            self.logo_inner,
            text="Coordination by Osama Elfeky",
            font=ctk.CTkFont(size=12),
            text_color=C["text_secondary"]
        )
        self.logo_subtitle.pack(pady=(2, 2))



        ctk.CTkFrame(self.logo_inner, height=2, fg_color=C["accent_purple"]).pack(fill="x", pady=(8, 0))

        self._animate_logo()

        ctk.CTkFrame(sb, height=1, fg_color=C["border"]).pack(fill="x", padx=16, pady=4)

        # IMPROVED PLACEHOLDER SECTION
        ctk.CTkLabel(sb, text="TEMPLATE TOOLS", font=ctk.CTkFont(size=10, weight="bold"),
                      text_color=C["text_muted"]).pack(anchor="w", padx=20, pady=(8, 0))

        # Template Selector
        self.tpl_var = ctk.StringVar(value=list(self.templates.keys())[0] if self.templates else "")
        self.tpl_menu = ctk.CTkOptionMenu(sb, variable=self.tpl_var,
            values=list(self.templates.keys()), width=240,
            fg_color=C["bg_input"], button_color=C["accent_blue"],
            button_hover_color=C["accent_purple"],
            command=self._on_template_select)
        self.tpl_menu.pack(padx=10, pady=(5, 5))
        
        # Insert Placeholder - Moved above editor for better flow
        self.ph_var = ctk.StringVar(value="✨ Insert Variable...")
        self.ph_menu = ctk.CTkOptionMenu(sb, variable=self.ph_var, width=240,
            values=["{name}", "{group_code}", "{time}", "{day}"],
            fg_color=C["bg_input"], button_color=C["accent_cyan"], 
            text_color=C["text_primary"], command=self._insert_placeholder)
        self.ph_menu.pack(padx=10, pady=(5, 5))

        self.tpl_editor = ctk.CTkTextbox(sb, height=150, fg_color=C["bg_input"],
            font=ctk.CTkFont(size=13), text_color=C["text_primary"],
            border_color=C["border"], border_width=1, corner_radius=8)
        self.tpl_editor.pack(padx=10, pady=2, fill="x")
        self.tpl_editor.bind("<KeyRelease>", lambda e: self._update_preview())

        # Action Buttons - Re-added Delete Template
        btn_frame = ctk.CTkFrame(sb, fg_color="transparent")
        btn_frame.pack(padx=10, pady=(4, 0), fill="x")
        StyledButton(btn_frame, text="💾 Save", width=70, command=self._save_current_tpl).pack(side="left", padx=2)
        StyledButton(btn_frame, text="➕ New", width=70, color=C["accent_purple"], command=self._new_template).pack(side="left", padx=2)
        StyledButton(btn_frame, text="🗑️", width=40, color=C["danger"], command=self._delete_template).pack(side="left", padx=2)
        StyledButton(btn_frame, text="😀", width=40, color=C["accent_cyan"],
                     hover=C["accent_blue"], command=self._show_emoji_picker).pack(side="left", padx=2)

        # LIVE PREVIEW - Not Italic
        ctk.CTkLabel(sb, text="LIVE PREVIEW", font=ctk.CTkFont(size=10, weight="bold"),
                      text_color=C["text_muted"]).pack(anchor="w", padx=20, pady=(12, 2))
        
        self.preview_box = ctk.CTkTextbox(sb, height=100, fg_color=C["bg_dark"],
            font=ctk.CTkFont(size=11), text_color=C["text_secondary"],
            border_color=C["border"], border_width=1, corner_radius=6)
        self.preview_box.pack(padx=10, pady=2, fill="x")
        self._load_tpl_text()

        ctk.CTkFrame(sb, height=1, fg_color=C["border"]).pack(fill="x", padx=16, pady=12)

        # Actions
        ctk.CTkLabel(sb, text="QUICK ACTIONS", font=ctk.CTkFont(size=10, weight="bold"),
                      text_color=C["text_muted"]).pack(anchor="w", padx=20, pady=(4, 6))
        
        StyledButton(sb, text="🌐 Verify WhatsApp Web", width=240, color=C["accent_purple"],
                     command=self.engine.open_whatsapp_web).pack(padx=10, pady=3)
        StyledButton(sb, text="📂 Load Excel File", width=240,
                     command=self._load_excel).pack(padx=10, pady=3)
        StyledButton(sb, text="📥 Download Template", width=240, color=C["accent_cyan"],
                     hover=C["accent_blue"], command=self._download_template).pack(padx=10, pady=3)

        StyledButton(sb, text="✨ Take a Pro Tour", width=240, color=C["status_green"],
                     hover="#15803d", command=self._start_app_tour).pack(padx=10, pady=(3, 3))

        # Delay & Wait config
        self.config_frame = ctk.CTkFrame(sb, fg_color="transparent")
        self.config_frame.pack(fill="x", padx=16, pady=(12, 0))
        ctk.CTkFrame(self.config_frame, height=1, fg_color=C["border"]).pack(fill="x", pady=12)
        ctk.CTkLabel(self.config_frame, text="AUTOMATION CONFIG", font=ctk.CTkFont(size=10, weight="bold"),
                      text_color=C["text_muted"]).pack(anchor="w", padx=4, pady=(4, 4))
        
        df = ctk.CTkFrame(self.config_frame, fg_color="transparent")
        df.pack(padx=0, fill="x")
        ctk.CTkLabel(df, text="Delay (s):", text_color=C["text_secondary"], font=ctk.CTkFont(size=11)).pack(side="left")
        self.delay_min = ctk.CTkEntry(df, width=40, fg_color=C["bg_input"])
        self.delay_min.insert(0, "20")
        self.delay_min.pack(side="left", padx=(4, 4))
        ctk.CTkLabel(df, text="-", text_color=C["text_secondary"]).pack(side="left")
        self.delay_max = ctk.CTkEntry(df, width=40, fg_color=C["bg_input"])
        self.delay_max.insert(0, "40")
        self.delay_max.pack(side="left", padx=4)

        wf = ctk.CTkFrame(sb, fg_color="transparent")
        wf.pack(padx=10, fill="x", pady=(8, 0))
        ctk.CTkLabel(wf, text="Wait Time:", text_color=C["text_secondary"], font=ctk.CTkFont(size=11)).pack(side="left")
        self.wait_time_entry = ctk.CTkEntry(wf, width=40, fg_color=C["bg_input"])
        self.wait_time_entry.insert(0, "15")
        self.wait_time_entry.pack(side="left", padx=(4, 8))

        # Config Explanation
        explanations = (
            "• Delay: Random pause between sending messages to mimic human behavior and prevent bans.\n"
            "• Wait Time: Initial seconds given to let WhatsApp Web fully load."
        )
        ctk.CTkLabel(sb, text=explanations, text_color=C["text_muted"], font=ctk.CTkFont(size=10),
                     justify="left", wraplength=230).pack(padx=15, pady=(8, 15), anchor="w")

    def _insert_placeholder(self, choice):
        if choice == "Insert Placeholder...": return
        self.tpl_editor.insert("insert", choice)
        self.ph_var.set("Insert Placeholder...") # Reset dropdown visually
        self._update_preview()

    def _update_preview(self, event=None):
        text = self.tpl_editor.get("1.0", "end-1c")
        
        # Logic: Use real data if a row is selected, else use dummy
        if self.last_selected_index is not None and self.rows[self.last_selected_index]["selected"]:
            data = self.rows[self.last_selected_index]
        else:
            data = {"name": "Ahmed Ali", "group_code": "G-102", "time": "02:00 PM", "day": "Friday"}

        full_name = data.get("name", "")
        first_name = full_name.split()[0] if full_name else ""
        
        text = text.replace("{name}", first_name)
        text = text.replace("{group_code}", data.get("group_code", ""))
        text = text.replace("{time}", data.get("time", ""))
        text = text.replace("{day}", data.get("day", ""))
        
        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", text if text.strip() else "Type a message above...")
        self.preview_box.configure(state="disabled")

    # ── Main Panel ───────────────────────────────────────────
    def _build_main(self):
        main = ctk.CTkFrame(self, fg_color=C["bg_dark"], corner_radius=0)
        main.pack(side="right", fill="both", expand=True)

        # Top bar
        top = ctk.CTkFrame(main, height=60, fg_color=C["bg_card"], corner_radius=0)
        top.pack(fill="x")
        top.pack_propagate(False)
        ctk.CTkLabel(top, text="📋  Message Queue", font=ctk.CTkFont(size=18, weight="bold"),
                      text_color=C["text_primary"]).pack(side="left", padx=20)

        self.status_label = ctk.CTkLabel(top, text="No file loaded",
            font=ctk.CTkFont(size=12), text_color=C["text_muted"])
        self.status_label.pack(side="right", padx=20)

        # Action bar — Row 1: Filters & core actions
        ab = ctk.CTkFrame(main, height=50, fg_color=C["bg_card"], corner_radius=0)
        ab.pack(fill="x", pady=(1, 0))
        ab.pack_propagate(False)

        # Slot Selector
        ctk.CTkLabel(ab, text="Time Slot:", text_color=C["accent_cyan"], font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(16, 4))
        self.slot_menu = ctk.CTkOptionMenu(ab, variable=self.slot_var, width=120,
            values=["All slots"],
            fg_color=C["bg_input"], button_color=C["accent_cyan"],
            command=self._on_slot_change)
        self.slot_menu.pack(side="left", padx=4)

        # Day Filter
        ctk.CTkLabel(ab, text="Day Filter:", text_color=C["text_secondary"]).pack(side="left", padx=(16, 4))
        self.day_filter_menu = ctk.CTkOptionMenu(ab, variable=self.day_filter_var, width=120,
            values=["All Days", "Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
            fg_color=C["bg_input"], button_color=C["accent_blue"],
            command=self._on_day_filter_change)
        self.day_filter_menu.pack(side="left", padx=4)

        # Send buttons
        StyledButton(ab, text="🚀 Send Selected", command=self._send_selected
                     ).pack(side="right", padx=4, pady=8)
        StyledButton(ab, text="🗑 Delete Selected", color=C["danger"], hover="#991b1b",
                     command=self._delete_selected).pack(side="right", padx=4, pady=8)
        StyledButton(ab, text="➕ Add Manual", color=C["accent_blue"],
                     command=self._add_manual_row).pack(side="right", padx=4, pady=8)
        self.bulk_paste_btn = StyledButton(ab, text="📋 Bulk Paste", color=C["accent_blue"],
                     command=self._bulk_paste_dialog)
        self.bulk_paste_btn.pack(side="right", padx=4, pady=8)
        StyledButton(ab, text="⛔ Stop", color=C["danger"], hover="#991b1b",
                     command=self._stop_sending).pack(side="right", padx=4, pady=8)

        # Action bar — Row 2: Schedule controls + Search bar
        ab2 = ctk.CTkFrame(main, height=46, fg_color=C["bg_card"], corner_radius=0)
        ab2.pack(fill="x", pady=(1, 0))
        ab2.pack_propagate(False)

        StyledButton(ab2, text="⏰ Schedule Manager", color=C["accent_purple"],
                     hover=C["accent_blue"],
                     command=self._open_schedule_dialog).pack(side="left", padx=(16, 4), pady=6)
        StyledButton(ab2, text="⛔ Cancel All", color=C["text_muted"],
                     hover=C["hover"], command=self._cancel_schedule
                     ).pack(side="left", padx=4, pady=6)

        # Search bar (inline in row 2)
        ctk.CTkLabel(ab2, text="🔍", font=ctk.CTkFont(size=14),
                     text_color=C["text_muted"]).pack(side="left", padx=(20, 4))

        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(
            ab2,
            textvariable=self.search_var,
            placeholder_text="Search by name, phone, or group code…",
            width=280,
            fg_color=C["bg_input"],
            border_color=C["border"],
            border_width=1,
        )
        self.search_entry.pack(side="left", padx=(0, 6), pady=6)
        self.search_var.trace_add("write", self._on_search_change)

        self.search_clear_btn = StyledButton(
            ab2, text="✕ Clear", width=65,
            color=C["bg_input"], text_color=C["text_secondary"], hover=C["hover"],
            command=self._clear_search,
        )
        self.search_clear_btn.pack(side="left", padx=2)

        self.search_result_lbl = ctk.CTkLabel(
            ab2, text="", font=ctk.CTkFont(size=11), text_color=C["text_muted"]
        )
        self.search_result_lbl.pack(side="right", padx=16)

        # Table header
        hdr = ctk.CTkFrame(main, height=36, fg_color=C["bg_sidebar"], corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        
        self.select_all_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(hdr, text="", variable=self.select_all_var, width=40,
            command=self._toggle_select_all, fg_color=C["accent_blue"],
            hover_color=C["accent_purple"]).pack(side="left", padx=(8, 0))

        cols = [("Instructor", 160), ("Phone", 130), ("Group Code", 130),
                ("Day", 90), ("Time", 90), ("Status", 100), ("Actions", 180)]
        for text, w in cols:
            ctk.CTkLabel(hdr, text=text, width=w, font=ctk.CTkFont(size=11, weight="bold"),
                          text_color=C["text_muted"]).pack(side="left", padx=2)

        # Scrollable table body
        self.table_frame = ctk.CTkScrollableFrame(main, fg_color=C["bg_dark"])
        self.table_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # Bottom stats bar
        self.stats_bar = ctk.CTkFrame(main, height=36, fg_color=C["bg_card"], corner_radius=0)
        self.stats_bar.pack(fill="x", side="bottom")
        self.stats_label = ctk.CTkLabel(self.stats_bar, text="",
            font=ctk.CTkFont(size=11), text_color=C["text_secondary"])
        self.stats_label.pack(side="left", padx=16)

    # ── Template helpers ─────────────────────────────────────
    def _load_tpl_text(self):
        self.tpl_editor.delete("1.0", "end")
        name = self.tpl_var.get()
        self.tpl_editor.insert("1.0", self.templates.get(name, ""))
        self._update_preview()

    def _on_template_select(self, _=None):
        self._load_tpl_text()

    def _save_current_tpl(self):
        name = self.tpl_var.get()
        self.templates[name] = self.tpl_editor.get("1.0", "end").strip()
        save_templates(self.templates)
        messagebox.showinfo("Saved", f"Template '{name}' saved.")

    def _new_template(self):
        dlg = ctk.CTkInputDialog(text="Template name:", title="New Template")
        name = dlg.get_input()
        if not name:
            return
        self.templates[name] = "Hello Eng. {name}, your {group_code} group..."
        save_templates(self.templates)
        self.tpl_menu.configure(values=list(self.templates.keys()))
        self.tpl_var.set(name)
        self._load_tpl_text()

    def _delete_template(self):
        name = self.tpl_var.get()
        if not name:
            return
        if messagebox.askyesno("Delete Template", f"Are you sure you want to delete the template '{name}'?"):
            if name in self.templates:
                del self.templates[name]
                if not self.templates:
                    self.templates["New Template"] = ""
                save_templates(self.templates)
                keys = list(self.templates.keys())
                self.tpl_menu.configure(values=keys)
                self.tpl_var.set(keys[0])
                self._load_tpl_text()

    def _reset_templates(self):
        if messagebox.askyesno("Reset Templates", "Are you sure you want to reset all templates to their defaults?\n\nThis will erase any custom templates you have created."):
            from app_config import DEFAULT_TEMPLATES
            self.templates = dict(DEFAULT_TEMPLATES)
            save_templates(self.templates)
            keys = list(self.templates.keys())
            self.tpl_menu.configure(values=keys)
            self.tpl_var.set(keys[0])
            self._load_tpl_text()
            messagebox.showinfo("Reset Successful", "Templates have been reset to defaults.")

    def _show_emoji_picker(self):
        emojis = [
            "😊", "😂", "🙏", "👍", "❤️", "✨", "✅", "❌", "⚠️", "📅", "🕒", "🎓", 
            "📚", "👋", "🎉", "🔥", "📌", "📢", "🌟", "💡", "📝", "📊", "🚀", "💪", 
            "🤝", "🙌", "👏", "😉", "😍", "😎", "🤔", "😅", "💯", "❗", "❓", "✅",
            "🏆", "🎯", "🥇", "👨‍🏫", "👩‍🏫", "⏰", "⌛", "💻", "📱", "📞", "✉️"
        ]
        dlg = ctk.CTkToplevel(self)
        dlg.title("Emojis")
        dlg.geometry("280x250")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        
        self.update_idletasks()
        x = self.winfo_x() + 50
        y = self.winfo_y() + 150
        dlg.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkScrollableFrame(dlg, fg_color="transparent")
        frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        row, col = 0, 0
        for em in emojis:
            btn = ctk.CTkButton(frame, text=em, width=35, height=35, font=ctk.CTkFont(size=16),
                                fg_color=C["bg_input"], hover_color=C["hover"],
                                command=lambda e=em: self._insert_emoji(e))
            btn.grid(row=row, column=col, padx=2, pady=2)
            col += 1
            if col > 4:
                col = 0
                row += 1

    def _insert_emoji(self, emoji):
        self.tpl_editor.insert("insert", emoji)
        self._update_preview()

    def _load_all_data(self):
        from app_config import DATA_FILE
        import json
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if data:
                        return data
            except: pass
        return {}

    def _save_all_data(self):
        from app_config import DATA_FILE
        import json
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(self.all_data, f, ensure_ascii=False, indent=2)
        self._auto_save_excel()

    def _auto_save_excel(self):
        try:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            from openpyxl.styles import Font, PatternFill, Alignment
            
            for slot, rows in self.all_data.items():
                if not rows: continue
                title = str(slot)[:31].replace("/", "_").replace("\\", "_")
                ws = wb.create_sheet(title=title)
                
                headers = ["Instructor", "Phone", "Group Code", "Day", "Time", "Status"]
                for c, h in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=c, value=h)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(fill_type="solid", start_color="1a73e8")
                    cell.alignment = Alignment(horizontal="center")
                    
                for ri, row in enumerate(rows, 2):
                    for ci, key in enumerate(["name","phone","group_code","day","time","status"], 1):
                        cell = ws.cell(row=ri, column=ci, value=row.get(key, ""))
                        cell.alignment = Alignment(horizontal="center")
                        if key == "status":
                            cmap = {STATUS_CONF: "22c55e", STATUS_ACK: "06b6d4",
                                    STATUS_SENT: "f59e0b", STATUS_NOT_SENT: "ef4444"}
                            c_hex = cmap.get(row["status"], "6b7280")
                            cell.font = Font(bold=True, color=c_hex)
                for c in range(1, len(headers)+1):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 18
            
            if not wb.sheetnames:
                wb.create_sheet("Empty")
                
            wb.save("Coordination_AutoSave.xlsx")
        except Exception:
            pass

    def _update_slot_menu(self):
        day_filter = self.day_filter_var.get()
        all_slots = sorted([s for s, rows in self.all_data.items() if rows])
        
        if day_filter == "All Days":
            filtered_slots = all_slots
        else:
            filtered_slots = [s for s in all_slots if s.startswith(day_filter)]
            
        times = set()
        for s in filtered_slots:
            if " - " in s:
                times.add(s.split(" - ", 1)[1])
            elif s == "General":
                times.add("General")
                
        sorted_times = sorted(list(times))
        menu_values = ["All slots"] + sorted_times
        
        self.slot_menu.configure(values=menu_values)
        
        if self.slot_var.get() not in menu_values:
            self.slot_var.set("All slots")
            
        self._populate_rows()

    def _on_day_filter_change(self, _):
        self._update_slot_menu()
        self._render_table()

    def _on_slot_change(self, slot=None):
        self._populate_rows()
        self.search_var.set("")
        self._render_table()
        self.status_label.configure(text=f"Switched to {self.current_slot} schedule")

    # ── Excel loading ────────────────────────────────────────
    def _load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            new_rows = []
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r[0]:
                    continue
                new_rows.append({
                    "name": str(r[0] or ""),
                    "phone": normalize_phone(r[1]),
                    "group_code": str(r[2] or ""),
                    "day": str(r[3] or "Friday").strip().title(),
                    "time": str(r[4] or ""),
                    "status": STATUS_NOT_SENT,
                    "selected": False,
                    "visible": True,
                })
            
            self._import_with_dedup(new_rows, source_label="Excel Import")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ── Table rendering ──────────────────────────────────────
    def _render_table(self):
        for w in self.table_frame.winfo_children():
            w.destroy()
        self.row_widgets.clear()

        for i, row in enumerate(self.rows):
            if not row["visible"]:
                self.row_widgets.append(None)
                continue
            bg = C["bg_card"] if i % 2 == 0 else C["bg_dark"]
            rf = ctk.CTkFrame(self.table_frame, height=44, fg_color=bg, corner_radius=0)
            rf.pack(fill="x", pady=1)
            rf.pack_propagate(False)

            var = ctk.BooleanVar(value=row["selected"])
            cb = ctk.CTkCheckBox(rf, text="", variable=var, width=40,
                fg_color=C["accent_blue"], hover_color=C["accent_purple"],
                command=lambda idx=i, v=var: self._toggle_row(idx, v))
            cb.pack(side="left", padx=(8, 0))

            name_disp = row["name"] if len(row["name"]) <= 22 else row["name"][:20] + "…"
            ctk.CTkLabel(rf, text=name_disp, width=160, anchor="w",
                text_color=C["text_primary"], font=ctk.CTkFont(size=12)).pack(side="left", padx=2)
            ctk.CTkLabel(rf, text=row["phone"], width=130, anchor="w",
                text_color=C["text_secondary"], font=ctk.CTkFont(size=12)).pack(side="left", padx=2)
            
            gc_disp = row["group_code"] if len(row["group_code"]) <= 18 else row["group_code"][:16] + "…"
            ctk.CTkLabel(rf, text=gc_disp, width=130, anchor="w",
                text_color=C["text_secondary"], font=ctk.CTkFont(size=11)).pack(side="left", padx=2)
            
            ctk.CTkLabel(rf, text=row.get("day", ""), width=90, anchor="w",
                text_color=C["text_secondary"], font=ctk.CTkFont(size=11)).pack(side="left", padx=2)
            ctk.CTkLabel(rf, text=row["time"], width=90, anchor="w",
                text_color=C["text_secondary"], font=ctk.CTkFont(size=11)).pack(side="left", padx=2)

            sf = ctk.CTkFrame(rf, fg_color="transparent", width=100)
            sf.pack(side="left", padx=2)
            sf.pack_propagate(False)
            dot = StatusDot(sf, color=STATUS_COLORS.get(row["status"], C["status_red"]))
            dot.pack(side="left", padx=(4, 6), pady=14)
            sl = ctk.CTkLabel(sf, text=row["status"], font=ctk.CTkFont(size=11),
                text_color=STATUS_COLORS.get(row["status"], C["text_secondary"]))
            sl.pack(side="left")
            sl.bind("<Button-1>", lambda e, idx=i: self._show_error_detail(idx))

            af = ctk.CTkFrame(rf, fg_color="transparent", width=180)
            af.pack(side="left", padx=4)
            StyledButton(af, text="✎ Edit", width=60, color=C["accent_purple"],
                hover=C["accent_blue"],
                command=lambda idx=i: self._edit_row_dialog(idx)).pack(side="left", padx=2, pady=6)
            StyledButton(af, text="🔄 Reset", width=65, color=C["text_muted"],
                hover=C["hover"],
                command=lambda idx=i: self._set_status(idx, STATUS_NOT_SENT)).pack(side="left", padx=2, pady=6)
            StyledButton(af, text="🗑", width=36, color=C["danger"],
                hover="#991b1b",
                command=lambda idx=i: self._delete_row(idx)).pack(side="left", padx=2, pady=6)

            self.row_widgets.append({"frame": rf, "cb_var": var, "dot": dot, "status_lbl": sl})

        self._update_stats()
        self._save_all_data()

    def _toggle_row(self, idx, var):
        self.rows[idx]["selected"] = var.get()
        if var.get():
            self.last_selected_index = idx
            self._update_preview() # Update preview with real data

    def _toggle_select_all(self):
        val = self.select_all_var.get()
        for i, row in enumerate(self.rows):
            if row["visible"]:
                row["selected"] = val
                w = self.row_widgets[i]
                if w:
                    w["cb_var"].set(val)

    def _set_status(self, row, status, detail=None):
        if isinstance(row, int):
            row_dict = self.rows[row]
            idx = row
        else:
            row_dict = row
            try:
                idx = self.rows.index(row_dict)
            except ValueError:
                idx = None

        row_dict["status"] = status
        if detail:
            row_dict["error_detail"] = detail
        
        if idx is not None:
            w = self.row_widgets[idx]
            if w:
                color = STATUS_COLORS.get(status, C["text_muted"])
                w["dot"].set_color(color)
                
                display_text = status
                if status == STATUS_FAILED and detail:
                    display_text = f"{STATUS_FAILED} (!)"
                
                w["status_lbl"].configure(text=display_text, text_color=color)
                
                if status == STATUS_FAILED:
                    err_msg = detail or "Unknown error"
                    self.status_label.configure(text=f"Error: {err_msg}", text_color=C["status_red"])
                    messagebox.showerror("Send Failed", 
                        f"Failed to send message to:\n{row_dict['name']} ({row_dict['phone']})\n\nError:\n{err_msg}")
        
        self._update_stats()
        self._save_all_data()

    def _show_error_detail(self, idx):
        row = self.rows[idx]
        if row["status"] == STATUS_FAILED:
            detail = row.get("error_detail", "No detailed error message available.")
            messagebox.showerror("Send Failure", 
                f"Instructor: {row['name']}\nPhone: {row['phone']}\n\nError:\n{detail}")

    def _delete_row(self, idx):
        row_to_delete = self.rows[idx]
        name = row_to_delete["name"]
        if messagebox.askyesno("Confirm Delete", f"Remove '{name}' from {self.current_slot}?"):
            for slot_key, slot_rows in self.all_data.items():
                if row_to_delete in slot_rows:
                    slot_rows.remove(row_to_delete)
                    break
            self._save_all_data()
            self._update_slot_menu()
            self._render_table()

    def _delete_selected(self):
        count = sum(1 for r in self.rows if r["selected"] and r["visible"])
        if count == 0:
            return
        if messagebox.askyesno("Confirm Delete", f"Delete {count} selected items from {self.current_slot}?"):
            to_delete = [r for r in self.rows if r["selected"] and r["visible"]]
            for r in to_delete:
                for slot_key, slot_rows in self.all_data.items():
                    if r in slot_rows:
                        slot_rows.remove(r)
                        break
            self._save_all_data()
            self._update_slot_menu()
            self._render_table()

    def _add_manual_row(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Add Manual Instructor")
        dlg.geometry("450x550")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        dlg.grab_set()
        
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 225
        y = self.winfo_y() + (self.winfo_height() // 2) - 275
        dlg.geometry(f"+{x}+{y}")

        content = ctk.CTkFrame(dlg, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=30, pady=30)

        ctk.CTkLabel(content, text="Manual Entry", font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=C["accent_blue"]).pack(pady=(0, 20))

        fields = {}
        items = [("Name", "name"), ("Phone", "phone"), ("Group Code", "group_code"), 
                 ("Day", "day"), ("Time", "time")]
        for label, key in items:
            ctk.CTkLabel(content, text=label, text_color=C["text_secondary"], font=ctk.CTkFont(size=12)).pack(anchor="w")
            entry = ctk.CTkEntry(content, width=350, fg_color=C["bg_input"], border_color=C["border"])
            if key == "day":
                entry.insert(0, "Friday")
            entry.pack(pady=(2, 12))
            fields[key] = entry

        def save():
            day_val = (fields["day"].get().strip() or "Friday").title()
            time_val = fields["time"].get().strip()
            slot_key = f"{day_val} - {time_val}" if time_val else day_val

            new_row = {
                "name": fields["name"].get().strip(),
                "phone": normalize_phone(fields["phone"].get().strip()),
                "group_code": fields["group_code"].get().strip(),
                "day": day_val,
                "time": time_val,
                "status": STATUS_NOT_SENT,
                "selected": False,
                "visible": True,
            }
            if not new_row["name"] or not new_row["phone"]:
                messagebox.showerror("Error", "Name and Phone are required.")
                return
            
            if slot_key not in self.all_data:
                self.all_data[slot_key] = []
                
            self.all_data[slot_key].append(new_row)
            self._save_all_data()
            
            day_filter = self.day_filter_var.get()
            time_filter = self.slot_var.get()
            matches_day = (day_filter == "All Days" or day_val == day_filter)
            matches_time = (time_filter == "All slots" or time_val == time_filter)
            
            if matches_day and matches_time:
                self._update_slot_menu()
                self._render_table()
            else:
                if messagebox.askyesno("Slot Changed", f"Instructor added to slot '{slot_key}'. Switch to that slot?"):
                    self.day_filter_var.set(day_val)
                    self.slot_var.set(time_val if time_val else "All slots")
                    self._update_slot_menu()
                    self._render_table()
            dlg.destroy()

        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))
        
        StyledButton(btn_frame, text="Cancel", color=C["bg_input"], text_color=C["text_primary"],
                     hover=C["hover"], command=dlg.destroy).pack(side="left", padx=5, expand=True)
        StyledButton(btn_frame, text="Add to Queue", command=save).pack(side="left", padx=5, expand=True)

    def _edit_row_dialog(self, idx):
        row = self.rows[idx]
        dlg = ctk.CTkToplevel(self)
        dlg.title(f"Edit: {row['name']}")
        dlg.geometry("450x550")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        dlg.grab_set()
        
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 225
        y = self.winfo_y() + (self.winfo_height() // 2) - 275
        dlg.geometry(f"+{x}+{y}")

        content = ctk.CTkFrame(dlg, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=30, pady=30)

        ctk.CTkLabel(content, text="Edit Entry", font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=C["accent_purple"]).pack(pady=(0, 20))

        fields = {}
        items = [("Name", "name"), ("Phone", "phone"), ("Group Code", "group_code"), 
                 ("Day", "day"), ("Time", "time")]
        
        for label, key in items:
            ctk.CTkLabel(content, text=label, text_color=C["text_secondary"], font=ctk.CTkFont(size=12)).pack(anchor="w")
            entry = ctk.CTkEntry(content, width=350, fg_color=C["bg_input"], border_color=C["border"])
            entry.insert(0, row.get(key, ""))
            entry.pack(pady=(2, 12))
            fields[key] = entry

        def save_changes():
            original_slot = None
            for slot_k, slot_rows in self.all_data.items():
                if row in slot_rows:
                    original_slot = slot_k
                    break
                    
            new_day = (fields["day"].get().strip() or "Friday").title()
            new_time = fields["time"].get().strip()
            new_slot = f"{new_day} - {new_time}" if new_time else new_day
                
            updated_data = {
                "name": fields["name"].get().strip(),
                "phone": normalize_phone(fields["phone"].get().strip()),
                "group_code": fields["group_code"].get().strip(),
                "day": new_day,
                "time": new_time,
            }
            
            if original_slot != new_slot:
                if new_slot not in self.all_data:
                    self.all_data[new_slot] = []
                
                row.update(updated_data)
                if original_slot and row in self.all_data[original_slot]:
                    self.all_data[original_slot].remove(row)
                self.all_data[new_slot].append(row)
                
                self._update_slot_menu()
                self._save_all_data()
                
                day_filter = self.day_filter_var.get()
                time_filter = self.slot_var.get()
                matches_day = (day_filter == "All Days" or new_day == day_filter)
                matches_time = (time_filter == "All slots" or new_time == time_filter)
                
                if matches_day and matches_time:
                    self._render_table()
                else:
                    if messagebox.askyesno("Slot Changed", f"Moved to slot '{new_slot}'. Switch to that slot?"):
                        self.day_filter_var.set(new_day)
                        self.slot_var.set(new_time if new_time else "All slots")
                        self._update_slot_menu()
                        self._render_table()
            else:
                row.update(updated_data)
                self._save_all_data()
                self._render_table()
            dlg.destroy()

        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))
        
        StyledButton(btn_frame, text="Cancel", color=C["bg_input"], text_color=C["text_primary"],
                     command=dlg.destroy).pack(side="left", padx=5, expand=True)
        StyledButton(btn_frame, text="Save Changes", color=C["accent_purple"], command=save_changes).pack(side="left", padx=5, expand=True)

    def _bulk_paste_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Grid Import — Paste & Edit")
        dlg.geometry("1000x700")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)
        dlg.grab_set()

        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 500
        y = self.winfo_y() + (self.winfo_height() // 2) - 350
        dlg.geometry(f"+{x}+{y}")

        content = ctk.CTkFrame(dlg, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)

        top_f = ctk.CTkFrame(content, fg_color="transparent")
        top_f.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(top_f, text="Grid Data Import", font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=C["accent_blue"]).pack(side="left")
        
        # --- NEW: Visually guide the user on formatting ---
        guide_frame = ctk.CTkFrame(content, fg_color=C["bg_sidebar"], corner_radius=6, border_width=1, border_color=C["accent_blue"])
        guide_frame.pack(fill="x", pady=(0, 15))
        guide_text = "💡 Excel/Sheets Format Required:   Column 1: Name   |   Column 2: Phone   |   Column 3: Group Code   |   Column 4: Day   |   Column 5: Time"
        ctk.CTkLabel(guide_frame, text=guide_text, font=ctk.CTkFont(size=12, weight="bold"), text_color=C["accent_cyan"]).pack(pady=8, padx=10)
        
        hdr = ctk.CTkFrame(content, height=30, fg_color=C["bg_sidebar"])
        hdr.pack(fill="x", pady=(0, 2))
        cols_cfg = [("Name*", 170), ("Phone*", 130), ("Group Code*", 130), ("Day", 90), ("Time", 90)]
        for t, w in cols_cfg:
            ctk.CTkLabel(hdr, text=t, width=w, font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=C["text_muted"]).pack(side="left", padx=5)

        grid_frame = ctk.CTkScrollableFrame(content, fg_color=C["bg_input"], corner_radius=8)
        grid_frame.pack(fill="both", expand=True)
        
        entry_rows = [] 

        def add_row():
            row_f = ctk.CTkFrame(grid_frame, fg_color="transparent")
            row_f.pack(fill="x", pady=2)
            row_data = {}
            for key, w in [("name", 170), ("phone", 130), ("group_code", 130), ("day", 90), ("time", 90)]:
                e = ctk.CTkEntry(row_f, width=w, fg_color=C["bg_card"], border_width=1)
                e.pack(side="left", padx=5)
                row_data[key] = e
            entry_rows.append(row_data)

        def paste_excel():
            try:
                raw = self.clipboard_get()
            except tk.TclError:
                messagebox.showerror("Error", "Clipboard is empty or contains non-text data.")
                return
            
            lines = [l for l in raw.strip().split('\n') if l.strip()]
            if not lines: return

            # Clear existing inputs to prevent mixing old and new data
            for r in entry_rows:
                for val in r.values():
                    val.delete(0, "end")

            while len(entry_rows) < len(lines):
                add_row()
            
            for i, line in enumerate(lines):
                parts = line.split('\t')
                row = entry_rows[i]
                fields = ["name", "phone", "group_code", "day", "time"]
                for pi, val in enumerate(parts):
                    if pi < len(fields):
                        row[fields[pi]].insert(0, val.strip())

        def validate_import():
            final = []
            errors = []
            for i, row in enumerate(entry_rows):
                vals = {k: e.get().strip() for k, e in row.items()}
                if not any(vals.values()): continue # Skip entirely blank rows
                
                missing = []
                if not vals["name"]: missing.append("Name")
                if not vals["phone"]: missing.append("Phone")
                if not vals["group_code"]: missing.append("Group Code")
                
                if missing:
                    errors.append(f"Row {i+1}: Missing {', '.join(missing)}")
                else:
                    vals["phone"] = normalize_phone(vals["phone"])
                    final.append({**vals, "status": STATUS_NOT_SENT, "selected": False, "visible": True})

            if errors:
                messagebox.showerror("Missing Data", "Fix errors before importing:\n\n" + "\n".join(errors[:15]))
                return
            
            if not final: return
            
            self._import_with_dedup(final, source_label="Bulk Paste")
            dlg.destroy()

        StyledButton(top_f, text="📋 Paste from Excel", width=160, color=C["accent_purple"],
                     command=paste_excel).pack(side="right", padx=5)
        StyledButton(top_f, text="➕ Add Blank Row", width=120, color=C["bg_input"], text_color=C["text_primary"],
                     command=add_row).pack(side="right", padx=5)

        for _ in range(10): add_row()

        btn_f = ctk.CTkFrame(content, fg_color="transparent")
        btn_f.pack(fill="x", pady=(20, 0))
        StyledButton(btn_f, text="Cancel", color=C["bg_input"], text_color=C["text_primary"],
                     command=dlg.destroy).pack(side="left", padx=5, expand=True)
        StyledButton(btn_f, text="Confirm & Import", color=C["status_green"], hover="#15803d",
                     command=validate_import).pack(side="right", padx=5, expand=True)

    def _update_stats(self):
        total = len(self.rows)
        sent = sum(1 for r in self.rows if r["status"] in (STATUS_SENT, STATUS_ACK, STATUS_CONF))
        conf = sum(1 for r in self.rows if r["status"] == STATUS_CONF)
        pending = total - sent
        self.stats_label.configure(
            text=f"Total: {total}  |  Sent: {sent}  |  Confirmed: {conf}  |  Pending: {pending}")

    # ── Sending ──────────────────────────────────────────────
    def _get_template_text(self):
        return self.tpl_editor.get("1.0", "end").strip()

    def _send_selected(self):
        selected = [{"row_index": i, **r} for i, r in enumerate(self.rows)
                     if r["selected"] and r["visible"] and r["status"] not in (STATUS_SENT, STATUS_ACK, STATUS_CONF)]
        if not selected:
            messagebox.showwarning("Nothing selected", "Select rows to send.")
            return
        self._do_send(selected)

    def _send_all(self):
        targets = [{"row_index": i, **r} for i, r in enumerate(self.rows)
                    if r["visible"] and r["status"] not in (STATUS_SENT, STATUS_ACK, STATUS_CONF)]
        if not targets:
            messagebox.showinfo("Done", "All visible messages already sent.")
            return
        self._do_send(targets)

    def _do_send(self, targets, tpl=None, slot_name=None, schedule_id=None):
        self._is_sending = True
        if tpl is None:
            tpl = self._get_template_text()
        if slot_name is None:
            slot_name = self.current_slot

        # CHECK IF ALREADY SENT
        already_sent = []
        for t in targets:
            row_dict = t["row_index"] if not isinstance(t["row_index"], int) else self.rows[t["row_index"]]
            if row_dict["status"] in (STATUS_SENT, STATUS_ACK, STATUS_CONF):
                already_sent.append(row_dict["name"])
        
        if already_sent:
            msg = f"The following instructors are already marked as Sent/Confirmed:\n\n" + \
                  "\n".join(already_sent[:5]) + ("\n..." if len(already_sent) > 5 else "") + \
                  "\n\nPlease reset their status to 'Not Sent' before sending again."
            messagebox.showwarning("Already Sent", msg)
            self._is_sending = False
            return

        try:
            mn, mx = int(self.delay_min.get()), int(self.delay_max.get())
            wt = int(self.wait_time_entry.get())
            self.engine.min_delay, self.engine.max_delay = mn, mx
            self.engine.wait_time = wt
        except ValueError:
            pass

        for t in targets:
            t["phone"] = normalize_phone(t.get("phone", ""))
            self._set_status(t["row_index"], STATUS_SENDING)

        progress_dlg = ctk.CTkToplevel(self)
        progress_dlg.title("Sending Messages")
        progress_dlg.geometry("380x280")
        progress_dlg.configure(fg_color=C["bg_dark"])
        progress_dlg.transient(self)
        progress_dlg.grab_set()

        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 190
        y = self.winfo_y() + (self.winfo_height() // 2) - 140
        progress_dlg.geometry(f"+{x}+{y}")

        self._send_title_lbl = ctk.CTkLabel(progress_dlg, text="📡 Sending Batch...", font=ctk.CTkFont(size=16, weight="bold"), text_color=C["accent_blue"])
        self._send_title_lbl.pack(pady=(20, 6))
        
        progress_bar = ctk.CTkProgressBar(progress_dlg, width=300, progress_color=C["accent_blue"])
        progress_bar.pack(pady=8)
        progress_bar.set(0)

        progress_label = ctk.CTkLabel(progress_dlg, text=f"0 / {len(targets)} sent", text_color=C["text_secondary"], font=ctk.CTkFont(size=12))
        progress_label.pack()

        # Countdown timer label — shows the delay between messages
        countdown_label = ctk.CTkLabel(progress_dlg, text="", text_color=C["accent_cyan"], font=ctk.CTkFont(size=13, weight="bold"))
        countdown_label.pack(pady=(6, 0))

        def stop_sending():
            self._stop_sending()
            if progress_dlg.winfo_exists():
                progress_dlg.destroy()

        self._send_stop_btn = StyledButton(progress_dlg, text="⛔ Stop", color=C["danger"], hover="#991b1b", command=stop_sending, width=100)
        self._send_stop_btn.pack(pady=(12, 0))

        progress_state = {"count": 0, "total": len(targets)}

        def on_prog(idx, ok, detail):
            self.after(0, lambda: self._set_status(idx, STATUS_SENT if ok else STATUS_FAILED, detail))
            def update_ui():
                progress_state["count"] += 1
                if progress_dlg.winfo_exists():
                    progress_bar.set(progress_state["count"] / progress_state["total"])
                    progress_label.configure(text=f"{progress_state['count']} / {progress_state['total']} sent")
                    # Clear countdown when a message just finished sending
                    countdown_label.configure(text="")
            self.after(0, update_ui)

        def on_delay(remaining, total_delay):
            def update_timer():
                if progress_dlg.winfo_exists():
                    if remaining > 0:
                        countdown_label.configure(
                            text=f"⏳ Next message in {remaining}s …"
                        )
                    else:
                        countdown_label.configure(text="Sending next message…")
            self.after(0, update_timer)

        def on_done():
            self._is_sending = False
            
            def get_row_dict(target):
                ri = target["row_index"]
                if isinstance(ri, int):
                    return self.rows[ri]
                return ri
                
            sent_ok = sum(
                1 for t in targets
                if get_row_dict(t)["status"] in (STATUS_SENT, STATUS_ACK, STATUS_CONF)
            )
            failed = len(targets) - sent_ok

            if schedule_id:
                for s in self._schedules:
                    if s["id"] == schedule_id:
                        s["status"] = "Completed" if failed == 0 else "Completed with failures"
                        self._save_schedules()
                        if hasattr(self, "_schedule_manager_dlg") and self._schedule_manager_dlg.winfo_exists():
                            self._refresh_schedule_manager_list()
                        break

            def show_completion():
                self.status_label.configure(text="Batch complete.")
                self._notify_batch_complete(
                    total=len(targets), sent_ok=sent_ok, slot=slot_name
                )

                # Transform progress dialog into a success/complete popup
                if progress_dlg.winfo_exists():
                    # Clear existing widgets
                    for child in progress_dlg.winfo_children():
                        child.destroy()

                    progress_dlg.geometry("400x300")
                    # Re-center
                    self.update_idletasks()
                    cx = self.winfo_x() + (self.winfo_width() // 2) - 200
                    cy = self.winfo_y() + (self.winfo_height() // 2) - 150
                    progress_dlg.geometry(f"+{cx}+{cy}")

                    if failed == 0:
                        icon = "✅"
                        title = "Batch Complete!"
                        detail = f"All {sent_ok} messages were sent successfully."
                        title_color = C["status_green"]
                    else:
                        icon = "⚠️"
                        title = "Batch Finished"
                        detail = f"{sent_ok} of {len(targets)} sent successfully.\n{failed} message(s) failed."
                        title_color = C["status_yellow"]

                    ctk.CTkLabel(progress_dlg, text=icon, font=ctk.CTkFont(size=48)).pack(pady=(25, 5))
                    ctk.CTkLabel(progress_dlg, text=title, font=ctk.CTkFont(size=20, weight="bold"), text_color=title_color).pack(pady=(0, 8))
                    ctk.CTkLabel(progress_dlg, text=detail, font=ctk.CTkFont(size=13), text_color=C["text_secondary"], wraplength=340).pack(pady=(0, 15))
                    ctk.CTkLabel(progress_dlg, text=f"Slot: {slot_name}", font=ctk.CTkFont(size=11), text_color=C["text_muted"]).pack()

                    StyledButton(progress_dlg, text="OK", width=120, color=C["accent_blue"],
                                 command=progress_dlg.destroy).pack(pady=(15, 0))

                    # Play notification sound
                    try:
                        winsound.MessageBeep(winsound.MB_ICONASTERISK)
                    except Exception:
                        pass

            self.after(0, show_completion)

        self.engine.send_batch(targets, tpl, on_progress=on_prog, on_complete=on_done, on_delay=on_delay)
        self.status_label.configure(text=f"Sending {len(targets)} messages...")

    def _stop_sending(self):
        self.engine.stop()
        self._is_sending = False
        self.status_label.configure(text="Stopped.")

    # ── Excel export ─────────────────────────────────────────
    def _download_template(self):
        fname = "Instructor_Template.xlsx"
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
            initialfile=fname, filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Template"
        headers = ["Instructor Name", "Phone Number", "Group Code", "Day", "Time"]
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        header_fill = PatternFill(start_color="1a73e8", end_color="1a73e8", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=12)
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
            
        sample_data = ["Mohammed Ahmed", "+966501234567", "GRP-A", "Friday", "2:00 PM"]
        for col, val in enumerate(sample_data, 1):
            cell = ws.cell(row=2, column=col, value=val)
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
            
        notes = [
            "⚠️ IMPORTANT FORMATTING NOTES ⚠️",
            "- Day: Single days (e.g., 'Friday', 'Saturday') or combined days (e.g., 'Friday/Saturday').",
            "- Time: Standard 12-hour or 24-hour formats (e.g., '09:00 AM', '03:00 PM', '15:00').",
            "- Phone Number: MUST include country code (e.g., Egyptian: '+201001234567' or '00201001234567')."
        ]

        note_start_row = 4
        for i, note in enumerate(notes):
            row_idx = note_start_row + i
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=len(headers))
            cell = ws.cell(row=row_idx, column=1, value=note)
            if i == 0:
                cell.font = Font(bold=True, color="FFFFFF", size=12)
                cell.fill = PatternFill(start_color="EF4444", end_color="EF4444", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                ws.row_dimensions[row_idx].height = 25
            else:
                cell.font = Font(bold=True, color="1E3A8A", size=11)
                cell.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
                cell.alignment = Alignment(horizontal='left', vertical='center')
                ws.row_dimensions[row_idx].height = 22
            
        for c in range(1, len(headers)+1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 20
            
        try:
            wb.save(path)
            messagebox.showinfo("Saved", f"Template saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save template:\n{str(e)}")


    def _on_search_change(self, *_):
        query = self.search_var.get().strip().lower()
        visible_count = 0
        for row in self.rows:
            if not query:
                row["visible"] = True
                visible_count += 1
            else:
                match = (
                    query in row.get("name", "").lower()
                    or query in row.get("phone", "").lower()
                    or query in row.get("group_code", "").lower()
                )
                row["visible"] = match
                if match:
                    visible_count += 1

        total = len(self.rows)
        if query:
            self.search_result_lbl.configure(
                text=f"{visible_count} of {total} shown",
                text_color=C["accent_cyan"] if visible_count else C["danger"],
            )
        else:
            self.search_result_lbl.configure(text="")

        self._render_table()

    def _clear_search(self):
        self.search_var.set("")
        self.search_entry.focus()

    # ── Feature 2: Duplicate Detection on Import ──────────
    def _get_all_existing_phones(self):
        """Return a set of every phone number currently stored across all slots."""
        phones = set()
        for rows in self.all_data.values():
            for row in rows:
                p = row.get("phone", "").strip()
                if p:
                    phones.add(p)
        return phones

    def _import_with_dedup(self, new_rows: list, source_label: str = "Import"):
        """
        Deduplicate new_rows against existing data, show a report, then import.
        """
        existing_phones = self._get_all_existing_phones()

        clean_rows = []
        duplicate_rows = []

        for row in new_rows:
            if row["phone"] in existing_phones:
                duplicate_rows.append(row)
            else:
                clean_rows.append(row)
                existing_phones.add(row["phone"])

        if duplicate_rows:
            lines = [f"  • {r['name']}  ({r['phone']})" for r in duplicate_rows[:10]]
            if len(duplicate_rows) > 10:
                lines.append(f"  … and {len(duplicate_rows) - 10} more")
            summary = "\n".join(lines)

            msg = (
                f"{source_label} found {len(duplicate_rows)} duplicate phone number(s) "
                f"already in your data:\n\n{summary}\n\n"
                f"Click YES to skip duplicates and import {len(clean_rows)} new rows.\n"
                f"Click NO to import all {len(new_rows)} rows anyway."
            )
            skip_dupes = messagebox.askyesno("Duplicates Found", msg)
            rows_to_import = clean_rows if skip_dupes else new_rows
        else:
            rows_to_import = clean_rows

        if not rows_to_import:
            messagebox.showinfo(source_label, "No new rows to import.")
            return

        for r in rows_to_import:
            d = r.get("day", "Friday").strip().title()
            r["day"] = d
            t = r.get("time", "").strip()
            slot_key = f"{d} - {t}" if t else d
            if slot_key not in self.all_data:
                self.all_data[slot_key] = []
            self.all_data[slot_key].append(r)

        self._update_slot_menu()
        self._save_all_data()
        self._render_table()
        self.status_label.configure(
            text=f"Imported {len(rows_to_import)} rows"
            + (f" ({len(duplicate_rows)} duplicates skipped)" if duplicate_rows else "")
        )


    # ── Feature 4: Desktop Notification on Batch Complete ─
    def _notify_batch_complete(self, total: int, sent_ok: int, slot: str):
        """Fire a desktop notification summarising the completed batch."""
        failed = total - sent_ok
        title = "DECI Automation — Batch Complete ✅"

        if failed == 0:
            msg = f"All {total} messages sent successfully.\nSlot: {slot}"
        else:
            msg = (
                f"{sent_ok}/{total} messages sent.\n"
                f"{failed} failed — check the queue for details.\n"
                f"Slot: {slot}"
            )

        try:
            from plyer import notification
            notification.notify(
                title=title,
                message=msg,
                app_name="DECI Automation",
                timeout=8,
            )
        except Exception:
            print(f"[DECI] Batch done: {msg}")

    def _start_app_tour(self):
        def select_dummy_for_preview():
            if self.rows:
                self.rows[0]["selected"] = True
                if self.row_widgets and self.row_widgets[0]:
                    self.row_widgets[0]["cb_var"].set(True)
                self.last_selected_index = 0
                self._update_preview()

        tour_steps = [
            {
                "title": "1. Mission Selector 📝",
                "msg": "Choose your mission: Switch between reminders for the day-before, pre-session readiness, or attendance checks. Each one is pre-loaded with professional emojis.",
                "widget": self.tpl_menu 
            },
            {
                "title": "2. Placeholder Menu 🏷️",
                "msg": "Don't type names manually. Click here to insert tags like {name} or {group_code}. This ensures every instructor gets a personalized message without you typing 50 different names.",
                "widget": self.ph_menu
            },
            {
                "title": "3. Live Preview ✨",
                "msg": "The 'What You See Is What They Get' (WYSIWYG) box. This shows the final message. It pulls data from the row you currently have selected.",
                "widget": self.preview_box,
                "action": select_dummy_for_preview
            },
            {
                "title": "4. Filters 🕒",
                "msg": "Organize your chaos. Managing 100 sessions? Filter by 'Friday' or '2:00 PM' to only see the instructors relevant to right now.",
                "widget": [self.slot_menu, self.day_filter_menu]
            },
            {
                "title": "5. Bulk Paste 📋",
                "msg": "The shortcut to efficiency. Copy directly from Excel and paste it into the grid. Just make sure your columns match: Name | Phone | Group | Day | Time.",
                "widget": self.bulk_paste_btn
            },
            {
                "title": "6. Account Bodyguard 🛡️",
                "msg": "The 'Delay' (20-40s) is vital. It prevents WhatsApp from flagging you as a spam bot. Never set this to 0!",
                "widget": self.config_frame,
                "final_msg": "You're all set! 🚀\n\nThe system is now configured. Remember to check your WhatsApp Web connection before starting a batch."
            }
        ]
        TourGuide(self, tour_steps)


def main():
    recipient_type = choose_recipient_type()
    if recipient_type == "instructors":
        app = App()
        app.mainloop()
    elif recipient_type == "coordinators":
        app = CoordinatorApp()
        app.mainloop()


if __name__ == "__main__":
    main()
