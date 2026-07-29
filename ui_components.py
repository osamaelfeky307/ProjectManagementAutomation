"""Reusable UI widgets for DECI app."""
import customtkinter as ctk
from app_config import COLORS

C = COLORS

class StatusDot(ctk.CTkFrame):
    """Colored dot indicator."""
    def __init__(self, master, color=C["status_red"], size=10, **kw):
        # Smaller default size for a refined look, explicitly setting bg_color to transparent
        super().__init__(master, width=size, height=size, corner_radius=size//2,
                         fg_color=color, bg_color="transparent", border_width=0, **kw)
        self._size = size
        self.configure(width=size, height=size)

    def set_color(self, color):
        self.configure(fg_color=color)


class StyledButton(ctk.CTkButton):
    """App-styled button."""
    def __init__(self, master, text="", color=C["accent_blue"], hover=None, **kw):
        hover = hover or C["accent_purple"]
        # Reduced corner radius and specified modern font (Inter)
        super().__init__(master, text=text, fg_color=color, hover_color=hover,
                         corner_radius=6, height=36, 
                         font=ctk.CTkFont(family="Inter", size=13, weight="bold"),
                         border_width=0,
                         **kw)


class SidebarButton(ctk.CTkButton):
    """Sidebar nav button."""
    def __init__(self, master, text="", **kw):
        # Modernized sidebar button with adjusted height and border spacing for inner padding
        super().__init__(master, text=text, fg_color="transparent",
                         hover_color=C["hover"], anchor="w", height=38,
                         corner_radius=6, font=ctk.CTkFont(family="Inter", size=13, weight="normal"), 
                         text_color=C["text_secondary"], border_spacing=10, **kw)

    def set_active(self, active=True):
        if active:
            # Active state uses a subtle background hover color and bright text instead of full brand color
            self.configure(fg_color=C["hover"], text_color=C["text_primary"],
                           font=ctk.CTkFont(family="Inter", size=13, weight="bold"))
        else:
            self.configure(fg_color="transparent", text_color=C["text_secondary"],
                           font=ctk.CTkFont(family="Inter", size=13, weight="normal"))


class CardFrame(ctk.CTkFrame):
    """A card-like frame with rounded corners and elevated background."""
    def __init__(self, master, **kw):
        kw.setdefault("fg_color", C["bg_card"])
        kw.setdefault("corner_radius", 10)
        super().__init__(master, **kw)


class StyledEntry(ctk.CTkEntry):
    """A premium text entry input."""
    def __init__(self, master, **kw):
        defaults = {
            "fg_color": C["bg_input"],
            "text_color": C["text_primary"],
            "border_color": C["border"],
            "border_width": 1,
            "corner_radius": 6,
            "font": ctk.CTkFont(family="Inter", size=13)
        }
        for k, v in defaults.items():
            kw.setdefault(k, v)
        super().__init__(master, **kw)


class StyledTextbox(ctk.CTkTextbox):
    """A premium multi-line text area."""
    def __init__(self, master, **kw):
        defaults = {
            "fg_color": C["bg_input"],
            "text_color": C["text_primary"],
            "border_color": C["border"],
            "border_width": 1,
            "corner_radius": 6,
            "font": ctk.CTkFont(family="Inter", size=13)
        }
        for k, v in defaults.items():
            kw.setdefault(k, v)
        super().__init__(master, **kw)


class StyledOptionMenu(ctk.CTkOptionMenu):
    """A premium dropdown menu."""
    def __init__(self, master, **kw):
        defaults = {
            "fg_color": C["bg_input"],
            "button_color": C["bg_input"],
            "button_hover_color": C["hover"],
            "text_color": C["text_primary"],
            "dropdown_fg_color": C["bg_card"],
            "dropdown_hover_color": C["hover"],
            "dropdown_text_color": C["text_primary"],
            "corner_radius": 6,
            "font": ctk.CTkFont(family="Inter", size=13),
            "dropdown_font": ctk.CTkFont(family="Inter", size=13)
        }
        for k, v in defaults.items():
            kw.setdefault(k, v)
        super().__init__(master, **kw)
