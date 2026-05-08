"""Reusable UI widgets for DECI app."""
import customtkinter as ctk
from app_config import COLORS

C = COLORS

class StatusDot(ctk.CTkFrame):
    """Colored dot indicator."""
    def __init__(self, master, color=C["status_red"], size=14, **kw):
        super().__init__(master, width=size, height=size, corner_radius=size//2,
                         fg_color=color, **kw)
        self._size = size
        self.configure(width=size, height=size)

    def set_color(self, color):
        self.configure(fg_color=color)


class StyledButton(ctk.CTkButton):
    """App-styled button."""
    def __init__(self, master, text="", color=C["accent_blue"], hover=None, **kw):
        hover = hover or C["accent_purple"]
        super().__init__(master, text=text, fg_color=color, hover_color=hover,
                         corner_radius=8, height=36, font=ctk.CTkFont(size=13, weight="bold"),
                         **kw)


class SidebarButton(ctk.CTkButton):
    """Sidebar nav button."""
    def __init__(self, master, text="", **kw):
        super().__init__(master, text=text, fg_color="transparent",
                         hover_color=C["hover"], anchor="w", height=40,
                         corner_radius=8, font=ctk.CTkFont(size=13), 
                         text_color=C["text_secondary"], **kw)

    def set_active(self, active=True):
        if active:
            self.configure(fg_color=C["accent_blue"], text_color=C["text_primary"])
        else:
            self.configure(fg_color="transparent", text_color=C["text_secondary"])
