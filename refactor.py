import re
import os

with open('schedule_manager.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Remove global TRIGGER_DEFS
content = re.sub(r'# ── trigger column definitions.*?\]\n', '', content, flags=re.DOTALL)

# 2. Add load/save methods and trigger editor inside ScheduleMixin
new_methods = '''
    def _get_trigger_defs(self):
        from app_config import TRIGGERS_FILE
        import json
        import datetime
        import os
        
        if not os.path.exists(TRIGGERS_FILE):
            default_triggers = [
                {"key": "day_before", "offset_minutes": -1440, "prefix": "1.", "label": "📅 Day Before"},
                {"key": "pre_session", "offset_minutes": -15, "prefix": "2.", "label": "🚨 -15 min"},
                {"key": "post_thanks", "offset_minutes": 120, "prefix": "6.", "label": "🌟 +2 h"}
            ]
            with open(TRIGGERS_FILE, "w", encoding="utf-8") as f:
                json.dump(default_triggers, f, ensure_ascii=False, indent=2)
                
        with open(TRIGGERS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            
        return [(d["key"], datetime.timedelta(minutes=d["offset_minutes"]), d["prefix"], d["label"]) for d in data]

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
                
                ctk.CTkLabel(row, text="Prefix:", width=40).pack(side="left")
                pfx_e = ctk.CTkEntry(row, width=40)
                pfx_e.insert(0, t.get("prefix", ""))
                pfx_e.pack(side="left", padx=2)
                
                ctk.CTkLabel(row, text="Key:", width=30).pack(side="left")
                key_e = ctk.CTkEntry(row, width=80)
                key_e.insert(0, t.get("key", ""))
                key_e.pack(side="left", padx=2)
                
                def del_row(idx=i):
                    triggers.pop(idx)
                    render_rows()
                    
                ctk.CTkButton(row, text="X", width=30, fg_color=C["danger"], command=del_row).pack(side="right", padx=2)
                
                entries.append((lbl_e, off_e, pfx_e, key_e))
                
        render_rows()
        
        def add_row():
            triggers.append({"key": f"trigger_{len(triggers)}", "offset_minutes": 0, "prefix": "", "label": "New Trigger"})
            render_rows()
            
        ctk.CTkButton(dlg, text="Add Trigger", command=add_row, fg_color=C["accent_purple"]).pack(pady=5)
        
        def save():
            new_triggers = []
            for lbl_e, off_e, pfx_e, key_e in entries:
                try:
                    new_triggers.append({
                        "key": key_e.get().strip(),
                        "offset_minutes": int(off_e.get().strip()),
                        "prefix": pfx_e.get().strip(),
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
'''

content = content.replace('class ScheduleMixin:\n    """All scheduling logic – mixed into App at class declaration time."""', 
                          'class ScheduleMixin:\n    """All scheduling logic – mixed into App at class declaration time."""\n' + new_methods)

# 3. Replace TRIGGER_DEFS with self._get_trigger_defs()
content = content.replace('TRIGGER_DEFS', 'self._get_trigger_defs()')

# 4. Add the 'Edit Triggers' button in _build_grid_dialog
edit_btn = '''
        StyledButton(top, text="Edit Triggers",
                     color=C["bg_input"], hover=C["hover"], width=120,
                     command=self._open_trigger_settings
                     ).pack(side="right", padx=8, pady=8)
'''
content = content.replace('        StyledButton(top, text="Cancel All Pending",', edit_btn + '        StyledButton(top, text="Cancel All Pending",')

with open('schedule_manager.py', 'w', encoding='utf-8') as f:
    f.write(content)
print('Refactored schedule_manager.py')
