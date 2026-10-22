"""centers_ui.py - CentersApp window for PM Automation."""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

import threading
import time
import random

from app_config import COLORS as C
from ui_components import (
    CardFrame, SidebarButton, StyledButton, StyledEntry,
    StyledTextbox, StatusDot, play_sound
)
from db_manager import (
    init_db, get_all_centers, insert_centers_bulk,
    update_center_details, update_center_confirmation,
    delete_center, clear_all_centers,
)
from centers_logic import load_sessions_for_centers, generate_center_messages, save_outputs, OUTPUT_LOG_FILE
from engine import MessageEngine

# Column aliases for Excel/CSV auto-mapping
_COL_ALIASES = {
    "center_name":  {"center name", "centername", "center", "name"},
    "governorate":  {"governorate", "gov", "governorate name"},
    "city":         {"city", "town"},
    "area":         {"area", "district"},
    "address":      {"address", "location"},
    "poc":          {"poc", "point of contact", "contact person", "contact"},
    "telephone":    {"telephone", "phone", "tel", "mobile"},
}


def _normalize_header(h: str) -> str:
    return h.strip().lower().replace(" ", "").replace("_", "").replace("-", "")


def _detect_column(header: str):
    norm = _normalize_header(header)
    for key, aliases in _COL_ALIASES.items():
        if norm in {_normalize_header(a) for a in aliases}:
            return key
    return None


def _load_file(path: str):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows_raw = list(ws.iter_rows(values_only=True))
        if not rows_raw:
            return [], []
        headers = [str(c).strip() if c is not None else "" for c in rows_raw[0]]
        data = [
            [str(c).strip() if c is not None else "" for c in row]
            for row in rows_raw[1:]
            if any(c not in (None, "") for c in row)
        ]
    else:
        import csv
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            all_rows = list(reader)
        if not all_rows:
            return [], []
        headers = [h.strip() for h in all_rows[0]]
        data = [[c.strip() for c in r] for r in all_rows[1:] if any(c.strip() for c in r)]
    return headers, data


# ===========================================================================
#  Edit / Confirm Dialog
# ===========================================================================

class _EditCenterDialog(ctk.CTkToplevel):
    def __init__(self, parent, center: dict, on_save):
        super().__init__(parent)
        self.title("Edit Center")
        self.geometry("560x700")
        self.resizable(False, True)
        self.configure(fg_color=C["bg_dark"])
        self.grab_set()
        self.lift()
        self.focus_force()

        self._center = center
        self._on_save = on_save

        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"560x700+{(sw - 560) // 2}+{(sh - 700) // 2}")

        self._build()

    def _build(self):
        c = self._center

        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=24, pady=(24, 0))
        scroll.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(scroll, text="Edit Center",
                     font=ctk.CTkFont(family="Inter", size=20, weight="bold"),
                     text_color=C["text_primary"]).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))

        self._section_label(scroll, "Basic Information", row=1)

        fields_basic = [
            ("center_name", "Center Name *"),
            ("governorate",  "Governorate"),
            ("city",         "City"),
            ("area",         "Area"),
            ("address",      "Address"),
            ("poc",          "POC"),
            ("telephone",    "Telephone"),
        ]
        self._entries: dict = {}
        for i, (key, label) in enumerate(fields_basic, start=2):
            ctk.CTkLabel(scroll, text=label,
                         font=ctk.CTkFont(family="Inter", size=12),
                         text_color=C["text_secondary"]).grid(
                row=i, column=0, sticky="w", padx=(0, 12), pady=4)
            ent = StyledEntry(scroll)
            ent.insert(0, str(c.get(key, "") or ""))
            ent.grid(row=i, column=1, sticky="ew", pady=4)
            self._entries[key] = ent

        row_off = len(fields_basic) + 2
        self._section_label(scroll, "POC Confirmation", row=row_off)
        row_off += 1

        ctk.CTkLabel(scroll, text="Groups Count",
                     font=ctk.CTkFont(family="Inter", size=12),
                     text_color=C["text_secondary"]).grid(
            row=row_off, column=0, sticky="w", padx=(0, 12), pady=4)
        self._groups_var = tk.StringVar(value=str(c.get("groups_count") or ""))
        StyledEntry(scroll, textvariable=self._groups_var).grid(
            row=row_off, column=1, sticky="ew", pady=4)
        row_off += 1

        ctk.CTkLabel(scroll, text="Rooms Count",
                     font=ctk.CTkFont(family="Inter", size=12),
                     text_color=C["text_secondary"]).grid(
            row=row_off, column=0, sticky="w", padx=(0, 12), pady=4)
        self._rooms_var = tk.StringVar(value=str(c.get("rooms_count") or ""))
        StyledEntry(scroll, textvariable=self._rooms_var).grid(
            row=row_off, column=1, sticky="ew", pady=4)
        row_off += 1

        self._confirmed_var = tk.BooleanVar(value=bool(c.get("confirmed", False)))
        sw_row = ctk.CTkFrame(scroll, fg_color="transparent")
        sw_row.grid(row=row_off, column=0, columnspan=2, sticky="w", pady=10)
        ctk.CTkSwitch(sw_row, text="Mark as Confirmed",
                      variable=self._confirmed_var,
                      onvalue=True, offvalue=False,
                      font=ctk.CTkFont(family="Inter", size=13),
                      text_color=C["text_primary"],
                      fg_color=C["accent_blue"],
                      progress_color=C["status_green"]).pack(side="left")
        row_off += 1

        ctk.CTkLabel(scroll, text="Notes",
                     font=ctk.CTkFont(family="Inter", size=12),
                     text_color=C["text_secondary"]).grid(
            row=row_off, column=0, sticky="nw", padx=(0, 12), pady=(8, 0))
        self._notes_box = StyledTextbox(scroll, height=80)
        self._notes_box.grid(row=row_off, column=1, sticky="ew", pady=(8, 0))
        if c.get("notes"):
            self._notes_box.insert("1.0", c["notes"])

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=16)
        btn_frame.grid_columnconfigure((0, 1), weight=1)

        StyledButton(btn_frame, text="Cancel",
                     color=C["bg_card"], hover=C["hover"],
                     command=self.destroy).grid(row=0, column=0, padx=(0, 8), sticky="ew")
        StyledButton(btn_frame, text="Save",
                     command=self._save).grid(row=0, column=1, padx=(8, 0), sticky="ew")

    def _section_label(self, parent, text, row):
        ctk.CTkLabel(parent, text=text,
                     font=ctk.CTkFont(family="Inter", size=13, weight="bold"),
                     text_color=C["accent_blue"]).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(16, 4))

    def _save(self):
        name = self._entries["center_name"].get().strip()
        if not name:
            messagebox.showerror("Validation", "Center Name is required.", parent=self)
            return

        def _int_or_none(s):
            s = str(s).strip()
            return int(s) if s.isdigit() else None

        groups    = _int_or_none(self._groups_var.get())
        rooms     = _int_or_none(self._rooms_var.get())
        confirmed = self._confirmed_var.get()
        notes     = self._notes_box.get("1.0", "end").strip()

        updated = {
            **self._center,
            "center_name": name,
            "governorate": self._entries["governorate"].get().strip(),
            "city":        self._entries["city"].get().strip(),
            "area":        self._entries["area"].get().strip(),
            "address":     self._entries["address"].get().strip(),
            "poc":         self._entries["poc"].get().strip(),
            "telephone":   self._entries["telephone"].get().strip(),
            "groups_count": groups,
            "rooms_count":  rooms,
            "confirmed":    confirmed,
            "notes":        notes,
        }

        update_center_details(
            updated["id"],
            updated["center_name"], updated["governorate"], updated["city"],
            updated["area"], updated["address"], updated["poc"], updated["telephone"],
        )
        update_center_confirmation(updated["id"], groups, rooms, confirmed, notes)

        self._on_save(updated)
        self.destroy()


# ===========================================================================
#  Upload Preview Dialog
# ===========================================================================

class _UploadPreviewDialog(ctk.CTkToplevel):
    def __init__(self, parent, centers_list: list, on_import):
        super().__init__(parent)
        self.title("Preview Import")
        w, h = 860, 560
        self.geometry(f"{w}x{h}")
        self.resizable(True, True)
        self.configure(fg_color=C["bg_dark"])
        self.grab_set()
        self.lift()
        self.focus_force()

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

        self._centers = centers_list
        self._on_import = on_import
        self._build()

    def _build(self):
        ctk.CTkLabel(self, text=f"Preview  -  {len(self._centers)} centers found",
                     font=ctk.CTkFont(family="Inter", size=16, weight="bold"),
                     text_color=C["text_primary"]).pack(pady=(20, 4), padx=24, anchor="w")

        ctk.CTkLabel(
            self,
            text="Review before importing. Duplicates (same name + telephone) will be skipped automatically.",
            font=ctk.CTkFont(family="Inter", size=12),
            text_color=C["text_secondary"],
        ).pack(padx=24, anchor="w", pady=(0, 12))

        scroll = ctk.CTkScrollableFrame(self, fg_color=C["bg_card"], corner_radius=10)
        scroll.pack(fill="both", expand=True, padx=24, pady=(0, 8))

        cols   = ["center_name", "governorate", "city", "area", "poc", "telephone"]
        heads  = ["Center Name",  "Governorate",  "City",  "Area", "POC", "Telephone"]
        widths = [200, 110, 100, 100, 140, 130]

        for ci, (h, w) in enumerate(zip(heads, widths)):
            ctk.CTkLabel(scroll, text=h,
                         font=ctk.CTkFont(family="Inter", size=12, weight="bold"),
                         text_color=C["accent_blue"],
                         width=w, anchor="w").grid(
                row=0, column=ci, padx=6, pady=(10, 4), sticky="w")

        for ri, c in enumerate(self._centers, start=1):
            for ci, (col, w) in enumerate(zip(cols, widths)):
                ctk.CTkLabel(scroll,
                             text=str(c.get(col, "") or ""),
                             font=ctk.CTkFont(family="Inter", size=12),
                             text_color=C["text_primary"],
                             width=w, anchor="w",
                             wraplength=w - 6).grid(
                    row=ri, column=ci, padx=6, pady=2, sticky="w")

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=24, pady=12)
        btn_frame.grid_columnconfigure((0, 1), weight=1)

        StyledButton(btn_frame, text="Cancel",
                     color=C["bg_card"], hover=C["hover"],
                     command=self.destroy).grid(row=0, column=0, padx=(0, 8), sticky="ew")
        StyledButton(btn_frame,
                     text=f"Import  {len(self._centers)}  Centers",
                     command=self._do_import).grid(row=0, column=1, padx=(8, 0), sticky="ew")

    def _do_import(self):
        self._on_import(self._centers)
        self.destroy()


# ===========================================================================
#  CentersApp - Main Window
# ===========================================================================

class CentersApp(ctk.CTk):
    """Full window for Centers Management."""

    _SIDEBAR_W = 210

    def __init__(self):
        super().__init__()
        init_db()
        self.title("PM Automation - Centers Management")
        self.geometry("1150x700")
        self.minsize(900, 580)
        self.configure(fg_color=C["bg_dark"])

        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"1150x700+{(sw - 1150) // 2}+{(sh - 700) // 2}")

        self._msg_data = []
        self._msg_statuses = {}
        self._msg_is_sending = False
        self._msg_thread = None
        self._engine = None

        self._centers: list = []
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", self._on_search_change)

        self._build_layout()
        self._load_centers()

    # Layout

    def _build_layout(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self._sidebar = ctk.CTkFrame(
            self, width=self._SIDEBAR_W,
            fg_color=C["bg_sidebar"], corner_radius=0,
        )
        self._sidebar.grid(row=0, column=0, sticky="nsew")
        self._sidebar.grid_propagate(False)
        self._build_sidebar()

        self._content = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self._content.grid(row=0, column=1, sticky="nsew")
        self._content.grid_rowconfigure(0, weight=1)
        self._content.grid_columnconfigure(0, weight=1)

        self._frames: dict = {}
        self._build_centers_list_view()
        self._build_upload_view()
        self._build_send_messages_view()
        self._show_view("centers_list")

    def _build_sidebar(self):
        sb = self._sidebar

        logo = ctk.CTkFrame(sb, fg_color="transparent")
        logo.pack(fill="x", padx=16, pady=(28, 12))
        ctk.CTkLabel(logo, text="🏢", font=ctk.CTkFont(size=36)).pack()
        ctk.CTkLabel(logo, text="Centers",
                     font=ctk.CTkFont(family="Inter", size=16, weight="bold"),
                     text_color=C["text_primary"]).pack()
        ctk.CTkLabel(logo, text="Management",
                     font=ctk.CTkFont(family="Inter", size=11),
                     text_color=C["text_muted"]).pack()

        ctk.CTkFrame(sb, height=1, fg_color=C["border"]).pack(fill="x", padx=16, pady=16)

        self._nav_btns: dict = {}
        for view, label in [("centers_list", "   Centers List"),
                             ("upload",       "   Upload Centers Sheet"),
                             ("send_messages", "📨 Send Messages")]:
            btn = SidebarButton(sb, text=label,
                                command=lambda v=view: self._show_view(v))
            btn.pack(fill="x", padx=8, pady=2)
            self._nav_btns[view] = btn

        ctk.CTkFrame(sb, fg_color="transparent").pack(fill="y", expand=True)
        ctk.CTkFrame(sb, height=1, fg_color=C["border"]).pack(fill="x", padx=16, pady=8)
        StyledButton(sb, text="<- Back to Launcher",
                     color=C["bg_card"], hover=C["hover"],
                     command=self._go_back).pack(fill="x", padx=16, pady=(0, 20))

    def _show_view(self, view: str):
        for name, frame in self._frames.items():
            if name == view:
                frame.grid(row=0, column=0, sticky="nsew")
            else:
                frame.grid_remove()
        for name, btn in self._nav_btns.items():
            btn.set_active(name == view)

    # Centers List View

    def _build_centers_list_view(self):
        frame = ctk.CTkFrame(self._content, fg_color="transparent")
        self._frames["centers_list"] = frame
        frame.grid_rowconfigure(2, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=28, pady=(24, 0))
        header.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(header, text="Training Centers",
                     font=ctk.CTkFont(family="Inter", size=22, weight="bold"),
                     text_color=C["text_primary"]).grid(row=0, column=0, sticky="w")

        self._stats_label = ctk.CTkLabel(
            header, text="",
            font=ctk.CTkFont(family="Inter", size=12),
            text_color=C["text_secondary"])
        self._stats_label.grid(row=1, column=0, sticky="w")

        right_bar = ctk.CTkFrame(header, fg_color="transparent")
        right_bar.grid(row=0, column=1, rowspan=2, sticky="e")

        StyledEntry(right_bar, textvariable=self._search_var,
                    placeholder_text="Search center, city, governorate...",
                    width=300).pack(side="left", padx=(0, 8))

        StyledButton(right_bar, text="Refresh",
                     color=C["bg_card"], hover=C["hover"],
                     width=72, command=self._load_centers).pack(side="left", padx=(0, 8))

        StyledButton(right_bar, text="Clear All",
                     color="#3A1C1C", hover=C["danger"],
                     command=self._clear_all_centers).pack(side="left")

        ctk.CTkFrame(frame, height=1, fg_color=C["border"]).grid(
            row=1, column=0, sticky="ew", padx=28, pady=16)

        self._table_scroll = ctk.CTkScrollableFrame(
            frame, fg_color=C["bg_card"], corner_radius=10)
        self._table_scroll.grid(row=2, column=0, sticky="nsew", padx=28, pady=(0, 28))

        col_weights = [0, 2, 1, 1, 1, 1, 0, 0, 0, 0]
        for ci, w in enumerate(col_weights):
            self._table_scroll.grid_columnconfigure(ci, weight=w)

        for ci, h in enumerate(["#", "Center Name", "Governorate", "City",
                                 "POC", "Telephone", "Groups", "Rooms", "Status", ""]):
            ctk.CTkLabel(self._table_scroll, text=h,
                         font=ctk.CTkFont(family="Inter", size=12, weight="bold"),
                         text_color=C["accent_blue"]).grid(
                row=0, column=ci, padx=8, pady=(10, 6), sticky="w")

        self._render_table()

    def _render_table(self):
        for widget in list(self._table_scroll.winfo_children()):
            info = widget.grid_info()
            if info and int(info.get("row", 0)) > 0:
                widget.destroy()

        search = self._search_var.get().lower().strip()
        visible = [
            c for c in self._centers
            if not search
            or search in c["center_name"].lower()
            or search in c["city"].lower()
            or search in c["governorate"].lower()
            or search in (c["poc"] or "").lower()
        ]

        if not visible:
            ctk.CTkLabel(
                self._table_scroll,
                text="No centers found. Upload a sheet to get started!",
                font=ctk.CTkFont(family="Inter", size=14),
                text_color=C["text_muted"],
            ).grid(row=1, column=0, columnspan=10, pady=60)
            return

        for ri, center in enumerate(visible, start=1):
            self._draw_row(ri, center)

    def _draw_row(self, row_idx: int, center: dict):
        ts = self._table_scroll
        confirmed = center.get("confirmed", False)
        status_color = C["status_green"] if confirmed else C["status_yellow"]

        def cell(text, col, width=None, anchor="w"):
            kw = dict(
                text=str(text) if text is not None else "-",
                font=ctk.CTkFont(family="Inter", size=12),
                text_color=C["text_primary"],
                anchor=anchor,
            )
            if width:
                kw["width"] = width
            ctk.CTkLabel(ts, **kw).grid(row=row_idx, column=col, padx=8, pady=5, sticky="w")

        cell(row_idx,                              col=0, width=32)
        cell(center["center_name"],                col=1)
        cell(center["governorate"],                col=2)
        cell(center["city"],                       col=3)
        cell(center["poc"],                        col=4)
        cell(center["telephone"],                  col=5)
        gval = center["groups_count"] if center["groups_count"] is not None else "-"
        rval = center["rooms_count"]  if center["rooms_count"]  is not None else "-"
        cell(gval, col=6, width=58)
        cell(rval, col=7, width=52)

        sf = ctk.CTkFrame(ts, fg_color="transparent")
        sf.grid(row=row_idx, column=8, padx=8, pady=5, sticky="w")
        StatusDot(sf, color=status_color, size=10).pack(side="left", padx=(0, 5))
        ctk.CTkLabel(sf,
                     text="Confirmed" if confirmed else "Pending",
                     font=ctk.CTkFont(family="Inter", size=11),
                     text_color=status_color).pack(side="left")

        af = ctk.CTkFrame(ts, fg_color="transparent")
        af.grid(row=row_idx, column=9, padx=(4, 12), pady=5, sticky="e")
        StyledButton(af, text="Edit", width=56,
                     command=lambda c=center: self._open_edit(c)).pack(side="left", padx=(0, 6))
        StyledButton(af, text="X", width=32,
                     color="#3A1C1C", hover=C["danger"],
                     command=lambda cid=center["id"]: self._delete_center(cid)).pack(side="left")


    # Upload View

    def _build_upload_view(self):
        frame = ctk.CTkFrame(self._content, fg_color="transparent")
        self._frames["upload"] = frame
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        hdr = ctk.CTkFrame(frame, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=28, pady=(24, 0))
        ctk.CTkLabel(hdr, text="Upload Centers Sheet",
                     font=ctk.CTkFont(family="Inter", size=22, weight="bold"),
                     text_color=C["text_primary"]).pack(anchor="w")
        ctk.CTkLabel(hdr,
                     text="Upload an Excel or CSV with columns: Center Name, Governorate, City, Area, Address, POC, Telephone",
                     font=ctk.CTkFont(family="Inter", size=12),
                     text_color=C["text_secondary"],
                     wraplength=700, justify="left").pack(anchor="w", pady=(4, 0))

        area = ctk.CTkScrollableFrame(frame, fg_color="transparent")
        area.grid(row=1, column=0, sticky="nsew", padx=28, pady=24)
        area.grid_columnconfigure(0, weight=1)

        # Drop-zone
        drop = CardFrame(area)
        drop.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        drop.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(drop, text="📁", font=ctk.CTkFont(size=52)).grid(row=0, column=0, pady=(32, 8))
        ctk.CTkLabel(drop, text="Click to browse for your Excel or CSV file",
                     font=ctk.CTkFont(family="Inter", size=14),
                     text_color=C["text_primary"]).grid(row=1, column=0, pady=(0, 4))
        ctk.CTkLabel(drop, text="Supported: .xlsx  .xlsm  .csv",
                     font=ctk.CTkFont(family="Inter", size=12),
                     text_color=C["text_muted"]).grid(row=2, column=0, pady=(0, 24))
        StyledButton(drop, text="Browse File",
                     command=self._browse_upload).grid(
            row=3, column=0, padx=100, pady=(0, 32), sticky="ew")

        # Column mapping info
        info = CardFrame(area)
        info.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        info.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(info, text="Column Auto-Mapping",
                     font=ctk.CTkFont(family="Inter", size=14, weight="bold"),
                     text_color=C["text_primary"]).grid(row=0, column=0, padx=20, pady=(16, 8), sticky="w")

        mapping_text = (
            "The app auto-detects your column headers. Accepted variants:\n\n"
            "  Center Name  ->  'Center Name', 'Center', 'centername', 'name'\n"
            "  Governorate  ->  'Governorate', 'Gov'\n"
            "  City         ->  'City', 'Town'\n"
            "  Area         ->  'Area', 'District'\n"
            "  Address      ->  'Address', 'Location'\n"
            "  POC          ->  'POC', 'Point of Contact', 'Contact Person', 'Contact'\n"
            "  Telephone    ->  'Telephone', 'Phone', 'Tel', 'Mobile'"
        )
        ctk.CTkLabel(info, text=mapping_text,
                     font=ctk.CTkFont(family="Inter", size=12),
                     text_color=C["text_secondary"],
                     justify="left", wraplength=680).grid(
            row=1, column=0, padx=20, pady=(0, 20), sticky="w")

        # Danger zone
        danger = CardFrame(area, fg_color="#1C1010",
                           border_color=C["danger"], border_width=1)
        danger.grid(row=2, column=0, sticky="ew")
        danger.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(danger, text="Danger Zone",
                     font=ctk.CTkFont(family="Inter", size=13, weight="bold"),
                     text_color=C["danger"]).grid(row=0, column=0, padx=20, pady=(16, 4), sticky="w")
        ctk.CTkLabel(danger,
                     text="Deletes ALL existing centers then imports the new file.",
                     font=ctk.CTkFont(family="Inter", size=12),
                     text_color=C["text_secondary"],
                     wraplength=680, justify="left").grid(
            row=1, column=0, padx=20, pady=(0, 10), sticky="w")
        StyledButton(danger, text="Clear All & Re-Import",
                     color="#3A1C1C", hover=C["danger"],
                     command=self._clear_and_reimport).grid(
            row=2, column=0, padx=20, pady=(0, 20), sticky="w")

    # Actions

    def _load_centers(self):
        self._centers = get_all_centers()
        self._update_stats()
        self._render_table()

    def _update_stats(self):
        total     = len(self._centers)
        confirmed = sum(1 for c in self._centers if c["confirmed"])
        pending   = total - confirmed
        self._stats_label.configure(
            text=f"{total} centers  |  {confirmed} confirmed  |  {pending} pending"
        )

    def _on_search_change(self, *_):
        self._render_table()

    def _open_edit(self, center: dict):
        def on_save(updated: dict):
            for i, c in enumerate(self._centers):
                if c["id"] == updated["id"]:
                    self._centers[i] = updated
                    break
            self._update_stats()
            self._render_table()

        _EditCenterDialog(self, center, on_save)

    def _delete_center(self, center_id: int):
        name = next((c["center_name"] for c in self._centers if c["id"] == center_id), "")
        if not messagebox.askyesno(
            "Delete Center",
            f"Delete '{name}'? This cannot be undone.",
            parent=self,
        ):
            return
        delete_center(center_id)
        self._centers = [c for c in self._centers if c["id"] != center_id]
        self._update_stats()
        self._render_table()

    def _clear_all_centers(self):
        if not self._centers:
            messagebox.showinfo("Nothing to Clear", "No centers in the database.", parent=self)
            return
        if not messagebox.askyesno(
            "Clear All Centers",
            f"Delete ALL {len(self._centers)} centers? This cannot be undone.",
            parent=self,
        ):
            return
        clear_all_centers()
        self._centers = []
        self._update_stats()
        self._render_table()

    def _browse_upload(self, clear_first: bool = False):
        path = filedialog.askopenfilename(
            parent=self,
            title="Select Centers Sheet",
            filetypes=[
                ("Excel or CSV", "*.xlsx *.xlsm *.csv"),
                ("Excel", "*.xlsx *.xlsm"),
                ("CSV", "*.csv"),
            ],
        )
        if not path:
            return
        self._process_upload(path, clear_first=clear_first)

    def _clear_and_reimport(self):
        if not messagebox.askyesno(
            "Clear & Re-Import",
            "This will DELETE all existing centers, then import from the new file.\n\nContinue?",
            parent=self,
        ):
            return
        self._browse_upload(clear_first=True)

    def _process_upload(self, path: str, clear_first: bool = False):
        try:
            headers, rows = _load_file(path)
        except Exception as exc:
            messagebox.showerror("File Error", f"Could not read the file:\n{exc}", parent=self)
            return

        if not headers:
            messagebox.showerror("Empty File", "The file appears to be empty.", parent=self)
            return

        col_map: dict = {}
        for ci, h in enumerate(headers):
            key = _detect_column(h)
            if key:
                col_map[ci] = key

        if "center_name" not in col_map.values():
            messagebox.showerror(
                "Missing Column",
                f"Could not detect a 'Center Name' column.\n\nFound: {', '.join(headers)}",
                parent=self,
            )
            return

        centers_preview: list = []
        for row in rows:
            entry: dict = {}
            for ci, key in col_map.items():
                entry[key] = row[ci] if ci < len(row) else ""
            if entry.get("center_name", "").strip():
                centers_preview.append(entry)

        if not centers_preview:
            messagebox.showwarning("No Data",
                                   "No valid center rows found in the file.", parent=self)
            return

        def on_import(centers_list: list):
            if clear_first:
                clear_all_centers()
                self._centers = []
            inserted, skipped = insert_centers_bulk(centers_list)
            self._load_centers()
            self._show_view("centers_list")
            msg = f"Imported {inserted} center(s)."
            if skipped:
                msg += f"\nSkipped {skipped} duplicate(s)."
            messagebox.showinfo("Import Complete", msg, parent=self)

        _UploadPreviewDialog(self, centers_preview, on_import)

    def _go_back(self):
        """Return to the LauncherApp."""
        from main import LauncherApp
        self.destroy()
        LauncherApp().mainloop()


    # ===========================================================================
    #  Send Messages View
    # ===========================================================================

    def _build_send_messages_view(self):
        frame = ctk.CTkFrame(self._content, fg_color="transparent")
        self._frames["send_messages"] = frame
        frame.grid_rowconfigure(2, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=28, pady=(24, 0))
        header.grid_columnconfigure(1, weight=1)

        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(title_frame, text="Send WhatsApp Messages",
                     font=ctk.CTkFont(family="Inter", size=22, weight="bold"),
                     text_color=C["text_primary"]).pack(anchor="w")
        
        self._msg_stats_label = ctk.CTkLabel(
            title_frame, text="Load a sessions sheet to generate messages for Centers.",
            font=ctk.CTkFont(family="Inter", size=12),
            text_color=C["text_secondary"])
        self._msg_stats_label.pack(anchor="w", pady=(4, 0))

        right_bar = ctk.CTkFrame(header, fg_color="transparent")
        right_bar.grid(row=0, column=1, sticky="e")

        StyledButton(right_bar, text="Export to Excel",
                     color=C["accent_blue"], hover=C["accent_purple"],
                     command=self._export_messages).pack(side="left", padx=(0, 8))

        StyledButton(right_bar, text="Load Sessions Sheet",
                     command=self._load_sessions_sheet).pack(side="left", padx=(0, 8))
        
        self._btn_send_all = StyledButton(right_bar, text="Start Sending",
                                          color=C["status_green"], hover="#258256",
                                          command=self._start_sending)
        self._btn_send_all.pack(side="left", padx=(0, 8))
        
        self._btn_stop = StyledButton(right_bar, text="Stop",
                                      color=C["danger"], hover="#C3363B",
                                      command=self._stop_sending)
        
        ctk.CTkFrame(frame, height=1, fg_color=C["border"]).grid(
            row=1, column=0, sticky="ew", padx=28, pady=16)

        self._msg_table_scroll = ctk.CTkScrollableFrame(
            frame, fg_color=C["bg_card"], corner_radius=10)
        self._msg_table_scroll.grid(row=2, column=0, sticky="nsew", padx=28, pady=(0, 28))

        col_weights = [0, 2, 2, 1, 1, 1, 0, 0]
        for ci, w in enumerate(col_weights):
            self._msg_table_scroll.grid_columnconfigure(ci, weight=w)

        for ci, h in enumerate(["#", "Center Name", "POC Name", "Phone", "Slots", "Max Rooms", "Status", ""]):
            ctk.CTkLabel(self._msg_table_scroll, text=h,
                         font=ctk.CTkFont(family="Inter", size=12, weight="bold"),
                         text_color=C["accent_blue"]).grid(
                row=0, column=ci, padx=8, pady=(10, 6), sticky="w")

    def _export_messages(self):
        if not self._msg_data:
            messagebox.showinfo("Export", "No messages to export.", parent=self)
            return
        
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Export Messages to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="Center_Messages.xlsx"
        )
        if not path:
            return
        
        try:
            import pandas as pd
            export_data = []
            for item in self._msg_data:
                clean_item = {k: v for k, v in item.items() if not k.startswith("_")}
                export_data.append(clean_item)
            
            df = pd.DataFrame(export_data)
            df.to_excel(path, index=False)
            messagebox.showinfo("Export Complete", f"Successfully exported to:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("Export Failed", str(exc), parent=self)

    def _load_sessions_sheet(self):
        path = filedialog.askopenfilename(
            parent=self, title="Select Sessions Sheet",
            filetypes=[("Excel", "*.xlsx *.xlsm")]
        )
        if not path:
            return
        
        try:
            sessions = load_sessions_for_centers(path)
            self._msg_data, skipped_logs = generate_center_messages(sessions)
            if skipped_logs:
                msg = "\n".join(f"- {log['Name']}: {log['Error']}" for log in skipped_logs[:5])
                if len(skipped_logs) > 5:
                    msg += f"\n...and {len(skipped_logs) - 5} more."
                messagebox.showwarning("Some Centers Skipped", msg, parent=self)
                
            self._msg_statuses = {idx: "Not Sent" for idx in range(len(self._msg_data))}
            self._render_msg_table()
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to load sessions:\n{exc}", parent=self)

    def _render_msg_table(self):
        ts = self._msg_table_scroll
        for widget in list(ts.winfo_children()):
            info = widget.grid_info()
            if info and int(info.get("row", 0)) > 0:
                widget.destroy()

        if not self._msg_data:
            self._msg_stats_label.configure(text="No messages generated.")
            return

        self._msg_stats_label.configure(text=f"Total Messages: {len(self._msg_data)}")

        for ri, msg_item in enumerate(self._msg_data, start=1):
            idx = ri - 1
            def cell(text, col, width=None):
                kw = dict(text=str(text), font=ctk.CTkFont(family="Inter", size=12), text_color=C["text_primary"], anchor="w")
                if width: kw["width"] = width
                ctk.CTkLabel(ts, **kw).grid(row=ri, column=col, padx=8, pady=5, sticky="w")
            
            cell(ri, 0, 32)
            cell(msg_item["Center"], 1)
            cell(msg_item["POC"], 2)
            cell(msg_item["Phone"], 3)
            cell(msg_item["Total Slots"], 4)
            cell(msg_item["Max Rooms Needed"], 5)

            status = self._msg_statuses.get(idx, "Not Sent")
            status_color = C["status_green"] if status == "Sent" else (C["danger"] if status == "Failed" else (C["status_yellow"] if status == "Sending..." else C["text_secondary"]))

            sf = ctk.CTkFrame(ts, fg_color="transparent")
            sf.grid(row=ri, column=6, padx=8, pady=5, sticky="w")
            StatusDot(sf, color=status_color, size=10).pack(side="left", padx=(0, 5))
            lbl = ctk.CTkLabel(sf, text=status, font=ctk.CTkFont(family="Inter", size=11), text_color=status_color)
            lbl.pack(side="left")
            
            # Store label ref so we can update it in real-time
            msg_item["_status_lbl"] = lbl
            msg_item["_status_dot"] = sf.winfo_children()[0]

            af = ctk.CTkFrame(ts, fg_color="transparent")
            af.grid(row=ri, column=7, padx=(4, 12), pady=5, sticky="e")
            StyledButton(af, text="Edit Message", width=90,
                         command=lambda i=idx: self._edit_center_message(i)).pack(side="left")

    def _edit_center_message(self, idx: int):
        msg_item = self._msg_data[idx]
        dlg = ctk.CTkToplevel(self)
        dlg.title(f"Edit Message - {msg_item['Center']}")
        dlg.geometry("500x600")
        dlg.configure(fg_color=C["bg_dark"])
        dlg.grab_set()
        dlg.lift()

        ctk.CTkLabel(dlg, text=f"To: {msg_item['POC']} ({msg_item['Phone']})", font=ctk.CTkFont(weight="bold")).pack(pady=10)
        textbox = StyledTextbox(dlg)
        textbox.pack(fill="both", expand=True, padx=20, pady=10)
        textbox.insert("1.0", msg_item["Message"])

        def _save():
            msg_item["Message"] = textbox.get("1.0", "end").strip()
            dlg.destroy()

        StyledButton(dlg, text="Save", command=_save).pack(pady=20)

    def _start_sending(self):
        if not self._msg_data: return
        self._msg_is_sending = True
        self._btn_send_all.pack_forget()
        self._btn_stop.pack(side="left", padx=(0, 8))
        self._engine = MessageEngine(min_delay=15, max_delay=30)
        
        self._msg_thread = threading.Thread(target=self._send_worker, daemon=True)
        self._msg_thread.start()

    def _stop_sending(self):
        self._msg_is_sending = False
        if self._engine:
            self._engine.close()

    def _update_msg_status(self, idx, status):
        self._msg_statuses[idx] = status
        msg_item = self._msg_data[idx]
        color = C["status_green"] if status == "Sent" else (C["danger"] if status == "Failed" else (C["status_yellow"] if status == "Sending..." else C["text_secondary"]))
        if "_status_lbl" in msg_item:
            msg_item["_status_lbl"].configure(text=status, text_color=color)
        if "_status_dot" in msg_item:
            msg_item["_status_dot"].configure(fg_color=color)

    def _send_worker(self):
        logs = []
        for idx, msg_item in enumerate(self._msg_data):
            if not self._msg_is_sending: break
            if self._msg_statuses.get(idx) == "Sent": continue

            self.after(0, self._update_msg_status, idx, "Sending...")
            
            try:
                ok, detail = self._engine.send_message(msg_item["Phone"], msg_item["Message"])
                status = "Sent" if ok else "Failed"
                logs.append({
                    "Timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Name": msg_item["Center"],
                    "Phone": msg_item["Phone"],
                    "Status": status.lower(),
                    "Error": "" if ok else detail
                })
            except Exception as exc:
                status = "Failed"
                logs.append({
                    "Timestamp": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Name": msg_item["Center"],
                    "Phone": msg_item["Phone"],
                    "Status": "failed",
                    "Error": str(exc)
                })

            self.after(0, self._update_msg_status, idx, status)

            if status == "Sent" and self._msg_is_sending and idx < len(self._msg_data) - 1:
                time.sleep(random.randint(15, 30))

        if self._engine:
            self._engine.close()
            self._engine = None
        
        self.after(0, self._finish_sending, logs)

    def _finish_sending(self, logs):
        self._msg_is_sending = False
        self._btn_stop.pack_forget()
        self._btn_send_all.pack(side="left", padx=(0, 8))
        if logs:
            save_outputs([], logs) # Just save logs for now
            play_sound()
            messagebox.showinfo("Sending Complete", f"Finished sending messages.\nLogs saved to {OUTPUT_LOG_FILE}.", parent=self)
