import datetime
import json
import os
import random
import threading
import time as _time

import customtkinter as ctk
from tkinter import messagebox


COL_WIDTH  = 148   # px per trigger column
ROW_HEIGHT = 80    # px per session row
HEADER_H   = 46


class ScheduleMixin:
    """All scheduling logic – mixed into App at class declaration time."""

    def _get_trigger_defs(self):
        from app_config import TRIGGERS_FILE
        import json
        import datetime
        import os
        
        if not os.path.exists(TRIGGERS_FILE):
            default_triggers = [
                {"key": "day_before", "offset_minutes": -1440, "template": "1. Day Before", "label": "📅 Day Before"},
                {"key": "pre_session", "offset_minutes": -15, "template": "2. Pre Session", "label": "🚨 -15 min"},
                {"key": "post_thanks", "offset_minutes": 120, "template": "6. Post Thanks", "label": "🌟 +2 h"}
            ]
            with open(TRIGGERS_FILE, "w", encoding="utf-8") as f:
                json.dump(default_triggers, f, ensure_ascii=False, indent=2)
                
        with open(TRIGGERS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            
        import time, random
        return [(d.get("key", f"t_{int(time.time())}_{random.randint(1000,9999)}"), datetime.timedelta(minutes=d["offset_minutes"]), d.get("template", d.get("prefix", "")), d["label"]) for d in data]

    def _open_trigger_settings(self):
        C = self._C()
        from app_config import TRIGGERS_FILE
        import json
        import os
        import customtkinter as ctk
        from tkinter import messagebox
        
        dlg = ctk.CTkToplevel(self._schedule_manager_dlg)
        dlg.title("Edit Triggers")
        dlg.geometry("620x500")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self._schedule_manager_dlg)
        dlg.grab_set()
        
        self._schedule_manager_dlg.update_idletasks()
        x = self._schedule_manager_dlg.winfo_x() + 100
        y = self._schedule_manager_dlg.winfo_y() + 100
        dlg.geometry(f"+{x}+{y}")
        
        ctk.CTkLabel(dlg, text="Trigger Configurations", font=ctk.CTkFont(size=16, weight="bold"), text_color=C["accent_cyan"]).pack(pady=10)
        
        scroll = ctk.CTkScrollableFrame(dlg, fg_color=C["bg_card"])
        scroll.pack(fill="both", expand=True, padx=20, pady=10)
        
        try:
            with open(TRIGGERS_FILE, "r", encoding="utf-8") as f:
                triggers = json.load(f)
        except:
            triggers = []
            
        entries = []
        
        def render_rows():
            for w in scroll.winfo_children():
                w.destroy()
            entries.clear()
            
            # Helper to get templates
            tpl_keys = list(getattr(self, "templates", {}).keys())
            
            for i, t in enumerate(triggers):
                row = ctk.CTkFrame(scroll, fg_color="transparent")
                row.pack(fill="x", pady=5)
                
                ctk.CTkLabel(row, text="Label:", width=40).pack(side="left")
                lbl_e = ctk.CTkEntry(row, width=120)
                lbl_e.insert(0, t.get("label", ""))
                lbl_e.pack(side="left", padx=2)
                
                ctk.CTkLabel(row, text="Offset(m):", width=60).pack(side="left")
                off_e = ctk.CTkEntry(row, width=60)
                off_e.insert(0, str(t.get("offset_minutes", 0)))
                off_e.pack(side="left", padx=2)
                
                ctk.CTkLabel(row, text="Template:", width=60).pack(side="left")
                
                # Resolve existing prefix to a full template name for display if necessary
                initial_tpl = t.get("template", t.get("prefix", ""))
                if initial_tpl not in tpl_keys:
                    found = self._find_template_by_prefix(initial_tpl)
                    if found:
                        initial_tpl = found
                if not initial_tpl or (initial_tpl not in tpl_keys and tpl_keys):
                    initial_tpl = tpl_keys[0] if tpl_keys else "No templates"
                
                tpl_var = ctk.StringVar(value=initial_tpl)
                tpl_menu = ctk.CTkOptionMenu(row, variable=tpl_var, values=tpl_keys or ["No templates"], width=160, fg_color=C["bg_input"], button_color=C["accent_blue"])
                tpl_menu.pack(side="left", padx=2)
                
                import time, random
                key_val = t.get("key", f"t_{int(time.time())}_{random.randint(1000,9999)}")
                
                def del_row(idx=i):
                    triggers.pop(idx)
                    render_rows()
                    
                ctk.CTkButton(row, text="X", width=30, fg_color=C["danger"], command=del_row).pack(side="right", padx=6)
                
                entries.append((lbl_e, off_e, tpl_var, key_val))
                
        render_rows()
        
        def add_row():
            import time, random
            triggers.append({"key": f"t_{int(time.time())}_{random.randint(1000,9999)}", "offset_minutes": 0, "template": "", "label": "New Trigger"})
            render_rows()
            
        ctk.CTkButton(dlg, text="Add Trigger", command=add_row, fg_color=C["accent_purple"]).pack(pady=5)
        
        def save():
            new_triggers = []
            for lbl_e, off_e, tpl_var, key_val in entries:
                try:
                    new_triggers.append({
                        "key": key_val,
                        "offset_minutes": int(off_e.get().strip()),
                        "template": tpl_var.get() if tpl_var.get() != "No templates" else "",
                        "label": lbl_e.get().strip()
                    })
                except ValueError:
                    messagebox.showerror("Error", "Offset must be an integer.", parent=dlg)
                    return
            with open(TRIGGERS_FILE, "w", encoding="utf-8") as f:
                json.dump(new_triggers, f, ensure_ascii=False, indent=2)
            self._rebuild_grid_rows(C)
            
            # Update header (since header depends on triggers)
            # Re-opening dialog to refresh header
            dlg.destroy()
            self._schedule_manager_dlg.destroy()
            self._open_schedule_dialog()
            
        bot = ctk.CTkFrame(dlg, fg_color="transparent")
        bot.pack(pady=10)
        ctk.CTkButton(bot, text="Cancel", command=dlg.destroy, fg_color=C["bg_input"]).pack(side="left", padx=10)
        ctk.CTkButton(bot, text="Save & Reload", command=save, fg_color=C["success"]).pack(side="left", padx=10)


    # ═════════════════════════════════════════════════════════════════
    #  MAIN ENTRY – opens the Schedule Manager window
    # ═════════════════════════════════════════════════════════════════

    def _open_schedule_dialog(self):
        C = self._C()

        # bring existing window to front instead of opening a second one
        if hasattr(self, "_schedule_manager_dlg") and \
                self._schedule_manager_dlg.winfo_exists():
            self._schedule_manager_dlg.lift()
            return

        if not hasattr(self, "_mgr_slot_var"):
            self._mgr_slot_var = ctk.StringVar(value="All slots")

        self._schedule_manager_dlg = ctk.CTkToplevel(self)
        dlg = self._schedule_manager_dlg
        dlg.title("Schedule Manager — Session x Trigger Grid")
        dlg.geometry("1180x700")
        dlg.minsize(900, 500)
        dlg.configure(fg_color=C["bg_dark"])
        dlg.transient(self)

        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width()  // 2) - 590
        y = self.winfo_y() + (self.winfo_height() // 2) - 350
        dlg.geometry(f"+{x}+{y}")

        self._build_grid_dialog(dlg, C)
        self._grid_tick(dlg)

    # ═════════════════════════════════════════════════════════════════
    #  GRID DIALOG LAYOUT
    # ═════════════════════════════════════════════════════════════════

    def _build_grid_dialog(self, dlg, C):
        from ui_components import StyledButton

        # ── top bar ──────────────────────────────────────────────────
        top = ctk.CTkFrame(dlg, fg_color=C["bg_card"],
                           corner_radius=0, height=52)
        top.pack(fill="x")
        top.pack_propagate(False)

        ctk.CTkLabel(top,
                     text="Session x Trigger Grid",
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=C["accent_cyan"]
                     ).pack(side="left", padx=20)

        all_slots = sorted([s for s, rows in getattr(self, "all_data", {}).items() if rows])
        ctk.CTkLabel(top, text="Filter Slot:", text_color=C["text_secondary"], 
                     font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(10, 4))
        self._mgr_slot_menu = ctk.CTkOptionMenu(top, variable=self._mgr_slot_var, width=160,
            values=["All slots"] + all_slots,
            fg_color=C["bg_input"], button_color=C["accent_cyan"],
            command=lambda _: self._rebuild_grid_rows(C))
        self._mgr_slot_menu.pack(side="left", padx=4)


        StyledButton(top, text="Edit Triggers",
                     color=C["bg_input"], hover=C["hover"], width=120,
                     command=self._open_trigger_settings
                     ).pack(side="right", padx=8, pady=8)
        StyledButton(top, text="Cancel All Pending",
                     color=C["danger"], hover="#991b1b", width=160,
                     command=self._cancel_all_schedules_from_mgr
                     ).pack(side="right", padx=8, pady=8)

        StyledButton(top, text="Schedule ALL Sessions",
                     color=C["accent_purple"], hover=C["accent_blue"], width=180,
                     command=self._schedule_all_sessions
                     ).pack(side="right", padx=4, pady=8)

        self._next_lbl = ctk.CTkLabel(
            top, text="",
            font=ctk.CTkFont(size=11),
            text_color=C["accent_cyan"])
        self._next_lbl.pack(side="right", padx=16)

        # ── column header ─────────────────────────────────────────────
        hdr = ctk.CTkFrame(dlg, fg_color=C["bg_sidebar"],
                           height=HEADER_H, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        ctk.CTkLabel(hdr, text="Session / Slot",
                     width=210, anchor="w",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=C["text_muted"]
                     ).pack(side="left", padx=(16, 0))

        for _, _, _, short in self._get_trigger_defs():
            ctk.CTkLabel(hdr, text=short,
                         width=COL_WIDTH, anchor="center",
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=C["text_muted"]
                         ).pack(side="left", padx=2)

        ctk.CTkLabel(hdr, text="", width=136).pack(side="left")  # actions spacer

        # ── scrollable body ───────────────────────────────────────────
        self._grid_body = ctk.CTkScrollableFrame(dlg, fg_color=C["bg_dark"])
        self._grid_body.pack(fill="both", expand=True)

        self._rebuild_grid_rows(C)

        # ── bottom status bar ─────────────────────────────────────────
        bot = ctk.CTkFrame(dlg, fg_color=C["bg_card"],
                           corner_radius=0, height=34)
        bot.pack(fill="x", side="bottom")
        bot.pack_propagate(False)
        self._grid_status_lbl = ctk.CTkLabel(
            bot, text="",
            font=ctk.CTkFont(size=11),
            text_color=C["text_secondary"])
        self._grid_status_lbl.pack(side="left", padx=16)

    # ─────────────────────────────────────────────────────────────────
    #  Build / rebuild all session rows
    # ─────────────────────────────────────────────────────────────────

    def _rebuild_grid_rows(self, C=None):
        if C is None:
            C = self._C()
        from ui_components import StyledButton

        if not hasattr(self, "_grid_body") or not self._grid_body.winfo_exists():
            return

        for w in self._grid_body.winfo_children():
            w.destroy()

        if not hasattr(self, "_grid_cells"):
            self._grid_cells = {}

        sorted_slots = sorted(
            [s for s, rows in self.all_data.items() if rows])

        if hasattr(self, "_mgr_slot_var"):
            selected_slot = self._mgr_slot_var.get()
            if selected_slot != "All slots" and selected_slot in sorted_slots:
                sorted_slots = [selected_slot]

        if not sorted_slots:
            ctk.CTkLabel(
                self._grid_body,
                text="No sessions loaded yet. Use 'Load Excel File' to add data.",
                font=ctk.CTkFont(size=14),
                text_color=C["text_muted"]
            ).pack(pady=60)
            return

        for i, slot_key in enumerate(sorted_slots):
            bg = C["bg_card"] if i % 2 == 0 else C["bg_dark"]

            # ── initialise cell state on first render ─────────────────
            if slot_key not in self._grid_cells:
                self._grid_cells[slot_key] = {}
                day, time_str = self._parse_slot_day_time(slot_key)
                now  = datetime.datetime.now()
                date = self._next_weekday(now, day)
                try:
                    t_obj = datetime.datetime.strptime(time_str, "%H:%M")
                    session_start = date.replace(
                        hour=t_obj.hour, minute=t_obj.minute,
                        second=0, microsecond=0)
                except ValueError:
                    session_start = date.replace(
                        hour=14, minute=0, second=0, microsecond=0)

                for key, offset, tpl_val, _ in self._get_trigger_defs():
                    trigger_dt = session_start + offset
                    tpl_name = tpl_val
                    if tpl_name not in getattr(self, "templates", {}):
                        # Attempt to resolve by prefix for backward compat
                        found = self._find_template_by_prefix(tpl_val)
                        if found:
                            tpl_name = found
                    if not tpl_name:
                        tpl_name = ""
                    self._grid_cells[slot_key][key] = {
                        "enabled":  ctk.BooleanVar(value=True),
                        "time_str": trigger_dt.strftime("%d/%m/%Y %H:%M"),
                        "tpl_name": tpl_name,
                        "sched_id": None,
                    }

            self._sync_cell_sched_ids(slot_key)

            # ── row frame ─────────────────────────────────────────────
            row_f = ctk.CTkFrame(self._grid_body, fg_color=bg,
                                 height=ROW_HEIGHT, corner_radius=0)
            row_f.pack(fill="x", pady=1)
            row_f.pack_propagate(False)

            # session name + recipient count
            lbl_f = ctk.CTkFrame(row_f, fg_color="transparent", width=210)
            lbl_f.pack(side="left", padx=(10, 0))
            lbl_f.pack_propagate(False)

            n = len(self.all_data.get(slot_key, []))
            ctk.CTkLabel(lbl_f, text=slot_key,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=C["text_primary"],
                         wraplength=198, anchor="w"
                         ).pack(anchor="w", pady=(12, 0))
            ctk.CTkLabel(lbl_f,
                         text=f"{n} recipient{'s' if n != 1 else ''}",
                         font=ctk.CTkFont(size=10),
                         text_color=C["text_muted"], anchor="w"
                         ).pack(anchor="w")

            # trigger cells
            for key, _, _, _ in self._get_trigger_defs():
                state  = self._grid_cells[slot_key][key]
                cell_f = ctk.CTkFrame(row_f, fg_color="transparent",
                                      width=COL_WIDTH)
                cell_f.pack(side="left", padx=2, pady=4)
                cell_f.pack_propagate(False)
                self._render_cell(cell_f, slot_key, key, state, C)

            # row action buttons
            act_f = ctk.CTkFrame(row_f, fg_color="transparent", width=130)
            act_f.pack(side="left", padx=6)
            act_f.pack_propagate(False)

            StyledButton(act_f, text="Schedule Row",
                         color=C["accent_purple"], hover=C["accent_blue"],
                         width=118, height=28,
                         command=lambda sk=slot_key: self._schedule_row(sk)
                         ).pack(pady=(6, 2))
            StyledButton(act_f, text="Cancel Row",
                         color=C["danger"], hover="#991b1b",
                         width=118, height=26,
                         command=lambda sk=slot_key: self._cancel_row(sk)
                         ).pack(pady=2)

    # ─────────────────────────────────────────────────────────────────
    #  Render one trigger cell
    # ─────────────────────────────────────────────────────────────────

    def _render_cell(self, parent, slot_key, trig_key, state, C):
        for w in parent.winfo_children():
            w.destroy()

        sched_id = state.get("sched_id")
        sched    = self._find_sched(sched_id) if sched_id else None

        # checkbox row
        cb_row = ctk.CTkFrame(parent, fg_color="transparent")
        cb_row.pack(fill="x", padx=2, pady=(4, 0))

        ctk.CTkCheckBox(cb_row, text="",
                        variable=state["enabled"], width=20,
                        fg_color=C["accent_purple"],
                        hover_color=C["accent_blue"]
                        ).pack(side="left")

        if sched:
            status = sched.get("status", "Pending")
            badge_bg, badge_txt = self._status_badge(status, C)
            ctk.CTkLabel(cb_row, text=badge_txt,
                         fg_color=badge_bg, corner_radius=4,
                         font=ctk.CTkFont(size=9, weight="bold"),
                         text_color=C["text_primary"],
                         width=52, height=16
                         ).pack(side="left", padx=3)

            # time + countdown
            try:
                target_dt = datetime.datetime.fromisoformat(sched["target_dt"])
                diff = (target_dt - datetime.datetime.now()).total_seconds()
                t_label = sched["target_dt"][11:16]   # HH:MM
                if diff > 0 and status == "Pending":
                    mins, secs = divmod(int(diff), 60)
                    hrs,  mins = divmod(mins, 60)
                    cd = f"{hrs}h{mins}m" if hrs else f"{mins}m{secs}s"
                    body = f"{t_label}\n{cd}"
                else:
                    body = f"{t_label}\n{status[:8]}"
            except Exception:
                body = status
        else:
            short_t = state["time_str"][11:16] if len(state["time_str"]) > 10 \
                else state["time_str"]
            ctk.CTkLabel(cb_row, text=short_t,
                         font=ctk.CTkFont(size=10),
                         text_color=C["text_muted"]
                         ).pack(side="left", padx=2)
            body = None

        if body:
            ctk.CTkLabel(parent, text=body,
                         font=ctk.CTkFont(size=10),
                         text_color=C["text_secondary"],
                         justify="center"
                         ).pack()

        # edit button
        ctk.CTkButton(parent, text="Edit",
                      width=80, height=22,
                      fg_color=C["bg_input"],
                      hover_color=C["hover"],
                      text_color=C["text_secondary"],
                      font=ctk.CTkFont(size=10),
                      command=lambda: self._open_cell_editor(
                          parent, slot_key, trig_key, state, C)
                      ).pack(pady=(2, 0))

    # ─────────────────────────────────────────────────────────────────
    #  Cell editor popup
    # ─────────────────────────────────────────────────────────────────

    def _open_cell_editor(self, cell_parent, slot_key, trig_key, state, C):
        from ui_components import StyledButton

        ed = ctk.CTkToplevel(self._schedule_manager_dlg)
        ed.title("Edit Trigger")
        ed.geometry("360x320")
        ed.configure(fg_color=C["bg_dark"])
        ed.transient(self._schedule_manager_dlg)
        ed.grab_set()
        ed.attributes("-topmost", True)

        self._schedule_manager_dlg.update_idletasks()
        px = self._schedule_manager_dlg.winfo_x()
        py = self._schedule_manager_dlg.winfo_y()
        pw = self._schedule_manager_dlg.winfo_width()
        ph = self._schedule_manager_dlg.winfo_height()
        ed.geometry(f"+{px + pw//2 - 180}+{py + ph//2 - 160}")

        short = next(
            (d[3] for d in self._get_trigger_defs() if d[0] == trig_key), trig_key)

        ctk.CTkLabel(ed,
                     text=f"Edit: {short}",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=C["accent_cyan"]
                     ).pack(pady=(18, 2), padx=16)
        ctk.CTkLabel(ed,
                     text=slot_key,
                     font=ctk.CTkFont(size=11),
                     text_color=C["text_muted"]
                     ).pack(pady=(0, 10))

        ctk.CTkLabel(ed, text="Date & Time  (DD/MM/YYYY HH:MM):",
                     text_color=C["text_secondary"],
                     font=ctk.CTkFont(size=11)
                     ).pack(anchor="w", padx=20)
        dt_entry = ctk.CTkEntry(ed, fg_color=C["bg_input"], width=300)
        dt_entry.insert(0, state["time_str"])
        dt_entry.pack(padx=20, pady=(2, 12))

        ctk.CTkLabel(ed, text="Message Template:",
                     text_color=C["text_secondary"],
                     font=ctk.CTkFont(size=11)
                     ).pack(anchor="w", padx=20)
        tpl_keys = list(self.templates.keys())
        cur_tpl  = state["tpl_name"] if state["tpl_name"] in tpl_keys \
            else (tpl_keys[0] if tpl_keys else "")
        tpl_var  = ctk.StringVar(value=cur_tpl)
        ctk.CTkOptionMenu(ed,
                          variable=tpl_var,
                          values=tpl_keys or ["No templates"],
                          fg_color=C["bg_input"],
                          button_color=C["accent_blue"],
                          button_hover_color=C["accent_purple"],
                          width=300
                          ).pack(padx=20, pady=(2, 12))

        err_lbl = ctk.CTkLabel(ed, text="",
                               text_color=C["danger"],
                               font=ctk.CTkFont(size=11))
        err_lbl.pack()

        def save():
            raw = dt_entry.get().strip()
            valid = False
            for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S"):
                try:
                    datetime.datetime.strptime(raw, fmt)
                    valid = True
                    break
                except ValueError:
                    pass
            if not valid:
                err_lbl.configure(text="Use DD/MM/YYYY HH:MM")
                return

            # cancel old live schedule if any
            old_id = state.get("sched_id")
            if old_id:
                self._cancel_schedule_item(old_id)
                state["sched_id"] = None

            state["time_str"] = raw
            state["tpl_name"] = tpl_var.get()
            self._render_cell(cell_parent, slot_key, trig_key, state, C)
            ed.destroy()

        btn_f = ctk.CTkFrame(ed, fg_color="transparent")
        btn_f.pack(fill="x", padx=20, pady=(4, 16))
        StyledButton(btn_f, text="Cancel",
                     color=C["bg_input"], text_color=C["text_primary"],
                     hover=C["hover"], command=ed.destroy
                     ).pack(side="left", expand=True, padx=4)
        StyledButton(btn_f, text="Save",
                     color=C["accent_purple"], hover=C["accent_blue"],
                     command=save
                     ).pack(side="left", expand=True, padx=4)

    # ═════════════════════════════════════════════════════════════════
    #  SCHEDULING ACTIONS
    # ═════════════════════════════════════════════════════════════════

    def _schedule_row(self, slot_key):
        C = self._C()
        cells   = self._grid_cells.get(slot_key, {})
        now     = datetime.datetime.now()
        created = 0
        skipped = []

        for key, _, _, short in self._get_trigger_defs():
            state = cells.get(key)
            if not state or not state["enabled"].get():
                continue

            # skip if already pending
            if state.get("sched_id"):
                s = self._find_sched(state["sched_id"])
                if s and s.get("status") == "Pending":
                    skipped.append(f"{short} already scheduled")
                    continue

            try:
                target_dt = datetime.datetime.strptime(
                    state["time_str"], "%d/%m/%Y %H:%M")
            except ValueError:
                skipped.append(f"{short} — bad time '{state['time_str']}'")
                continue

            if target_dt <= now:
                skipped.append(f"{short} — time is in the past")
                continue

            tpl_name = state["tpl_name"]
            tpl_text = self.templates.get(tpl_name, "")
            if not tpl_text:
                skipped.append(f"{short} — template '{tpl_name}' is empty/missing")
                continue

            ok, res = self._add_schedule(target_dt, slot_key, tpl_name, tpl_text)
            if ok:
                state["sched_id"] = res
                created += 1
            else:
                skipped.append(f"{short} — {res}")

        self._rebuild_grid_rows(C)

        if created:
            msg = f"Scheduled {created} trigger(s) for '{slot_key}'."
            if skipped:
                msg += "\n\nSkipped:\n" + "\n".join(f"  {s}" for s in skipped)
                messagebox.showwarning("Partial Schedule", msg)
            else:
                self._grid_status_lbl.configure(
                    text=f"  {created} trigger(s) scheduled for {slot_key}",
                    text_color=C["status_green"])
            self._beep()
        else:
            err = "\n".join(f"  {s}" for s in skipped) if skipped \
                else "No enabled triggers with valid future times."
            messagebox.showerror("Nothing Scheduled", err)

    def _cancel_row(self, slot_key):
        cells = self._grid_cells.get(slot_key, {})
        count = 0
        for state in cells.values():
            sid = state.get("sched_id")
            if sid:
                s = self._find_sched(sid)
                if s and s.get("status") == "Pending":
                    self._cancel_schedule_item(sid)
                    count += 1
                state["sched_id"] = None
        self._rebuild_grid_rows()
        C = self._C()
        self._grid_status_lbl.configure(
            text=f"Cancelled {count} trigger(s) for {slot_key}",
            text_color=C["text_muted"])

    def _schedule_all_sessions(self):
        if not hasattr(self, "_grid_cells"):
            return
        C   = self._C()
        now = datetime.datetime.now()
        total    = 0
        all_skip = []

        for slot_key, cells in self._grid_cells.items():
            for key, _, _, short in self._get_trigger_defs():
                state = cells.get(key)
                if not state or not state["enabled"].get():
                    continue
                if state.get("sched_id"):
                    s = self._find_sched(state["sched_id"])
                    if s and s.get("status") == "Pending":
                        continue
                try:
                    target_dt = datetime.datetime.strptime(
                        state["time_str"], "%d/%m/%Y %H:%M")
                except ValueError:
                    all_skip.append(f"[{slot_key}] {short} bad time")
                    continue
                if target_dt <= now:
                    all_skip.append(f"[{slot_key}] {short} in the past")
                    continue
                tpl_name = state["tpl_name"]
                tpl_text = self.templates.get(tpl_name, "")
                if not tpl_text:
                    all_skip.append(f"[{slot_key}] {short} template empty")
                    continue
                ok, res = self._add_schedule(
                    target_dt, slot_key, tpl_name, tpl_text)
                if ok:
                    state["sched_id"] = res
                    total += 1
                else:
                    all_skip.append(f"[{slot_key}] {short} — {res}")

        self._rebuild_grid_rows(C)
        if total:
            msg = f"Scheduled {total} trigger(s) across all sessions."
            if all_skip:
                msg += "\n\nSkipped:\n" + "\n".join(
                    f"  {s}" for s in all_skip[:20])
            self._grid_status_lbl.configure(
                text=f"  {total} triggers scheduled",
                text_color=C["status_green"])
            messagebox.showinfo("All Sessions Scheduled", msg)
            self._beep()
        else:
            note = ("No new triggers created.\n"
                    "Check that times are in the future and "
                    "templates are non-empty.")
            if all_skip:
                note += "\n\nSkipped:\n" + "\n".join(
                    f"  {s}" for s in all_skip[:15])
            messagebox.showwarning("Nothing Scheduled", note)

    def _cancel_all_schedules_from_mgr(self):
        if messagebox.askyesno("Cancel All", "Cancel ALL pending schedules?"):
            self._cancel_all_schedules()
            if hasattr(self, "_grid_cells"):
                for cells in self._grid_cells.values():
                    for state in cells.values():
                        state["sched_id"] = None
            self._rebuild_grid_rows()

    # ─────────────────────────────────────────────────────────────────
    #  Live countdown tick (runs every second while dialog is open)
    # ─────────────────────────────────────────────────────────────────

    def _grid_tick(self, dlg):
        if not dlg.winfo_exists():
            return

        now      = datetime.datetime.now()
        min_diff = None
        upcoming = None
        for s in self._schedules:
            if s.get("status") == "Pending":
                try:
                    dt   = datetime.datetime.fromisoformat(s["target_dt"])
                    diff = (dt - now).total_seconds()
                    if diff > 0 and (min_diff is None or diff < min_diff):
                        min_diff = diff
                        upcoming = s
                except Exception:
                    pass

        if upcoming and min_diff is not None:
            mins, secs = divmod(int(min_diff), 60)
            hrs,  mins = divmod(mins, 60)
            sn = upcoming.get("template_name", "")
            if "." in sn:
                sn = sn.split(".", 1)[1].strip()
            txt = (f"Next: {hrs}h {mins}m {secs}s — "
                   f"{upcoming['slot']} ({sn})") if hrs else \
                  (f"Next: {mins}m {secs}s — "
                   f"{upcoming['slot']} ({sn})")
            if hasattr(self, "_next_lbl") and self._next_lbl.winfo_exists():
                self._next_lbl.configure(text=txt)
        else:
            if hasattr(self, "_next_lbl") and self._next_lbl.winfo_exists():
                self._next_lbl.configure(text="No pending triggers")

        # refresh cell badge countdowns every 5 s
        if not hasattr(self, "_grid_tick_n"):
            self._grid_tick_n = 0
        self._grid_tick_n += 1
        if self._grid_tick_n % 5 == 0:
            self._rebuild_grid_rows()

        dlg.after(1000, lambda: self._grid_tick(dlg))

    # ═════════════════════════════════════════════════════════════════
    #  CORE SCHEDULE PRIMITIVES
    # ═════════════════════════════════════════════════════════════════

    def _add_schedule(self, target_dt, slot, template_name, template_text):
        """Create a timer and register the schedule. Returns (True, sched_id)."""
        now   = datetime.datetime.now()
        delay = (target_dt - now).total_seconds()
        if delay <= 0:
            return False, "Chosen time is in the past."

        sched_id = f"sched_{int(_time.time())}_{random.randint(1000, 9999)}"
        sched = {
            "id":            sched_id,
            "target_dt":     target_dt.isoformat(),
            "slot":          slot,
            "template_name": template_name,
            "template_text": template_text,
            "status":        "Pending",
        }

        def _fire(sid=sched_id):
            self.after(0, lambda: self._fire_schedule(sid))

        t = threading.Timer(delay, _fire)
        t.daemon = True
        t.start()

        self._schedule_timers[sched_id] = t
        self._schedules.append(sched)
        self._save_schedules()
        return True, sched_id

    def _load_schedules(self):
        from app_config import TEMPLATES_FILE
        sf = os.path.join(os.path.dirname(TEMPLATES_FILE), "schedules.json")

        self._schedules = []
        if os.path.exists(sf):
            try:
                with open(sf, "r", encoding="utf-8") as f:
                    self._schedules = json.load(f)
            except Exception:
                pass

        if hasattr(self, "_schedule_timers"):
            for t in self._schedule_timers.values():
                try:
                    t.cancel()
                except Exception:
                    pass
        self._schedule_timers = {}

        now     = datetime.datetime.now()
        updated = False
        for s in self._schedules:
            if s.get("status") == "Pending":
                try:
                    target_dt = datetime.datetime.fromisoformat(s["target_dt"])
                    delay     = (target_dt - now).total_seconds()
                    if delay > 0:
                        sid = s["id"]

                        def _fire(sid=sid):
                            self.after(0, lambda: self._fire_schedule(sid))

                        t = threading.Timer(delay, _fire)
                        t.daemon = True
                        t.start()
                        self._schedule_timers[sid] = t
                    else:
                        s["status"] = "Expired"
                        updated = True
                except Exception:
                    s["status"] = "Error"
                    updated = True

        if updated:
            self._save_schedules()

    def _save_schedules(self):
        from app_config import TEMPLATES_FILE
        sf = os.path.join(os.path.dirname(TEMPLATES_FILE), "schedules.json")
        clean = [{
            "id":            s["id"],
            "target_dt":     s["target_dt"],
            "slot":          s["slot"],
            "template_name": s["template_name"],
            "template_text": s["template_text"],
            "status":        s["status"],
        } for s in self._schedules]
        try:
            with open(sf, "w", encoding="utf-8") as f:
                json.dump(clean, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[DECI] save_schedules error: {e}")

    def _cancel_schedule_item(self, sched_id):
        if sched_id in self._schedule_timers:
            try:
                self._schedule_timers[sched_id].cancel()
            except Exception:
                pass
            del self._schedule_timers[sched_id]
        for s in self._schedules:
            if s["id"] == sched_id:
                s["status"] = "Cancelled"
                break
        self._save_schedules()

    def _cancel_all_schedules(self, silent=False):
        for t in self._schedule_timers.values():
            try:
                t.cancel()
            except Exception:
                pass
        self._schedule_timers = {}
        for s in self._schedules:
            if s.get("status") == "Pending":
                s["status"] = "Cancelled"
        self._save_schedules()
        if not silent:
            self.status_label.configure(
                text="All active schedules cancelled.",
                text_color="#9ca3af")
            messagebox.showinfo("Cancelled",
                                "All pending schedules have been cancelled.")

    def _cancel_schedule(self, silent=False):
        self._cancel_all_schedules(silent=silent)

    def _fire_schedule(self, sched_id):
        C     = self._C()
        sched = self._find_sched(sched_id)
        if not sched or sched.get("status") != "Pending":
            return

        # postpone if another batch is already running
        if self._is_sending:
            now        = datetime.datetime.now()
            new_target = now + datetime.timedelta(seconds=60)
            sched["target_dt"] = new_target.isoformat()
            self._save_schedules()

            def _fire(sid=sched_id):
                self.after(0, lambda: self._fire_schedule(sid))

            t = threading.Timer(60, _fire)
            t.daemon = True
            t.start()
            self._schedule_timers[sched_id] = t
            self.status_label.configure(
                text=f"Schedule '{sched['slot']}' postponed — another batch running.",
                text_color=C["status_yellow"])
            return

        slot = sched["slot"]

        # resolve status constants at runtime to avoid circular imports
        import sys
        main_mod = sys.modules.get("__main__")
        S_SENT = getattr(main_mod, "STATUS_SENT", "Sent")
        S_ACK  = getattr(main_mod, "STATUS_ACK",  "Acknowledged")
        S_CONF = getattr(main_mod, "STATUS_CONF", "Confirmed")

        targets   = []
        slot_keys = list(self.all_data.keys()) if slot == "All slots" else [slot]
        for sk in slot_keys:
            for r in self.all_data.get(sk, []):
                if r.get("status") not in (S_SENT, S_ACK, S_CONF):
                    targets.append({
                        "name":       r["name"],
                        "phone":      r["phone"],
                        "group_code": r["group_code"],
                        "time":       r.get("time", ""),
                        "day":        r.get("day", ""),
                        "row_index":  r,
                    })

        if not targets:
            sched["status"] = "Completed (No Recipients)"
            self._save_schedules()
            self.status_label.configure(
                text=f"'{slot}' fired — no unsent recipients found.",
                text_color=C["text_secondary"])
            return

        sched["status"] = "Running"
        self._save_schedules()
        self.status_label.configure(
            text=f"Firing scheduled send for '{slot}'…",
            text_color=C["accent_cyan"])

        self._do_send(targets,
                      tpl=sched["template_text"],
                      slot_name=slot,
                      schedule_id=sched_id)

    # ═════════════════════════════════════════════════════════════════
    #  STATUS-BAR COUNTDOWN  (runs even when the dialog is closed)
    # ═════════════════════════════════════════════════════════════════

    def _start_global_countdown(self):
        def tick():
            now      = datetime.datetime.now()
            min_diff = None
            upcoming = None
            for s in self._schedules:
                if s.get("status") == "Pending":
                    try:
                        dt   = datetime.datetime.fromisoformat(s["target_dt"])
                        diff = (dt - now).total_seconds()
                        if diff > 0 and (min_diff is None or diff < min_diff):
                            min_diff = diff
                            upcoming = s
                    except Exception:
                        pass

            if upcoming and min_diff is not None:
                mins, secs = divmod(int(min_diff), 60)
                hrs,  mins = divmod(mins, 60)
                sn = upcoming.get("template_name", "")
                if "." in sn:
                    sn = sn.split(".", 1)[1].strip()
                lbl = (f"Next: {hrs}h {mins}m {secs}s — "
                       f"{upcoming['slot']} ({sn})") if hrs else \
                      (f"Next: {mins}m {secs}s — "
                       f"{upcoming['slot']} ({sn})")
                self.status_label.configure(text=lbl, text_color="#06b6d4")
            else:
                curr = self.status_label.cget("text")
                if curr and "Next:" in curr:
                    self.status_label.configure(
                        text="No active schedules.", text_color="#9ca3af")

            self._schedule_countdown_job = self.after(1000, tick)

        tick()

    # ═════════════════════════════════════════════════════════════════
    #  HELPERS
    # ═════════════════════════════════════════════════════════════════

    def _C(self):
        from app_config import COLORS
        return COLORS

    def _find_sched(self, sched_id):
        return next(
            (s for s in self._schedules if s["id"] == sched_id), None)

    def _sync_cell_sched_ids(self, slot_key):
        """Back-fill sched_id into grid cells from self._schedules."""
        if not hasattr(self, "_grid_cells"):
            return
        cells = self._grid_cells.get(slot_key, {})
        for key, _, _, _ in self._get_trigger_defs():
            state = cells.get(key)
            if state is None or state.get("sched_id"):
                continue
            tpl_name = state.get("tpl_name", "")
            for s in self._schedules:
                if (s.get("slot") == slot_key
                        and s.get("template_name") == tpl_name
                        and s.get("status") in ("Pending", "Running")):
                    state["sched_id"] = s["id"]
                    break

    def _find_template_by_prefix(self, prefix):
        for name in self.templates:
            if name.strip().startswith(prefix):
                return name
        return None

    @staticmethod
    def _status_badge(status, C):
        mapping = {
            "Pending":   (C["accent_purple"], "PENDING"),
            "Running":   (C["status_yellow"], "RUNNING"),
            "Cancelled": (C["text_muted"],    "CANCELLED"),
            "Expired":   (C["text_muted"],    "EXPIRED"),
        }
        if "Completed" in status:
            return C["status_green"], "DONE"
        return mapping.get(status, (C["accent_blue"], status[:8].upper()))

    @staticmethod
    def _next_weekday(d, weekday_name):
        days = {"monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
                "friday": 4, "saturday": 5, "sunday": 6}
        target = days.get(weekday_name.lower())
        if target is None:
            return d
        ahead = target - d.weekday()
        if ahead <= 0:
            ahead += 7
        return d + datetime.timedelta(days=ahead)

    @staticmethod
    def _parse_slot_day_time(slot_key):
        day, time_str = "Friday", "14:00"
        if " - " in slot_key:
            parts  = slot_key.split(" - ", 1)
            day    = parts[0].strip()
            t_part = parts[1].strip()
            for fmt in ("%I:%M %p", "%H:%M"):
                try:
                    time_str = datetime.datetime.strptime(
                        t_part, fmt).strftime("%H:%M")
                    break
                except ValueError:
                    pass
        else:
            day = slot_key.strip()
        return day, time_str

    # ── legacy stubs so nothing breaks if old call-sites remain ──────
    def _cancel_schedule_from_mgr(self, sched_id):
        self._cancel_schedule_item(sched_id)

    def _refresh_schedule_manager_list(self):
        pass  # replaced by _rebuild_grid_rows / _grid_tick

    def _add_single_schedule_from_mgr(self, *_):
        pass

    def _add_suite_schedules_from_mgr(self, *_):
        pass

    def _on_suite_slot_change(self, *_):
        pass

    @staticmethod
    def _beep():
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass
