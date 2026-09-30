from __future__ import annotations

UI_COLORS = {
    "background": "#F4F7FB",
    "card": "#FFFFFF",
    "card_alt": "#F8FAFC",
    "border": "#E2E8F0",
    "border_focus": "#93C5FD",
    "text": "#0F172A",
    "text_secondary": "#475569",
    "text_muted": "#94A3B8",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_light": "#EFF6FF",
    "success": "#16A34A",
    "success_light": "#F0FDF4",
    "warning": "#D97706",
    "warning_light": "#FFFBEB",
    "danger": "#DC2626",
    "danger_light": "#FEF2F2",
    "danger_border": "#FECACA",
    "log_background": "#0F172A",
    "log_panel": "#111827",
    "log_text": "#D1D5DB",
}


UI_FONT = "Microsoft YaHei UI"


MONO_FONT = "Consolas"


def configure_styles(style):
    style.configure(
        "App.TEntry",
        fieldbackground=UI_COLORS["card_alt"],
        foreground=UI_COLORS["text"],
        bordercolor=UI_COLORS["border"],
        lightcolor=UI_COLORS["border"],
        darkcolor=UI_COLORS["border"],
        insertcolor=UI_COLORS["primary"],
        padding=(10, 9),
        font=(UI_FONT, 10),
    )
    style.map(
        "App.TEntry",
        bordercolor=[("focus", UI_COLORS["border_focus"])],
        lightcolor=[("focus", UI_COLORS["border_focus"])],
        darkcolor=[("focus", UI_COLORS["border_focus"])],
    )
    style.configure(
        "Blue.Horizontal.TProgressbar",
        troughcolor="#E8EEF7",
        background=UI_COLORS["primary"],
        bordercolor="#E8EEF7",
        lightcolor=UI_COLORS["primary"],
        darkcolor=UI_COLORS["primary"],
        thickness=10,
    )
    style.configure(
        "App.Vertical.TScrollbar",
        troughcolor=UI_COLORS["card"],
        background="#CBD5E1",
        bordercolor=UI_COLORS["card"],
        arrowcolor=UI_COLORS["text_secondary"],
        gripcount=0,
    )
    style.map(
        "App.Vertical.TScrollbar",
        background=[("active", "#94A3B8")],
    )
