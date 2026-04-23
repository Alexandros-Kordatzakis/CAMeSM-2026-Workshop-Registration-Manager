"""
CAMeSM Workshop Registration Manager
=====================================
Desktop GUI — CustomTkinter + Matplotlib

Changes v2:
  • First name + Last name columns detected and combined automatically
  • Designed for any number of workshops (tested logic for 11+)
  • Profile badges wrap across rows — no overflow with many workshops
  • Richer filter bar: min-workshops dropdown, not just a checkbox
  • Export: per-workshop buttons wrap gracefully across rows
  • Allocation table wraps "other_workshops" text correctly

Usage:
    python app.py
"""

import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from pathlib import Path
import math
import os

# ── Theme ─────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BLUE    = "#1A6FA3"
BLUE_LT = "#2596D4"
BLUE_VL = "#3CBFEF"
AMBER   = "#F0A500"
GREEN   = "#27AE60"
RED     = "#C0392B"
BG      = "#1a1a2e"
PANEL   = "#16213e"
CARD    = "#0f3460"
TEXT    = "#e0e0e0"
SUB     = "#9ea3b0"

PALETTE = [
    "#1A6FA3","#E05C2A","#27AE60","#F0A500",
    "#9B59B6","#16A085","#C0392B","#2596D4",
    "#8E44AD","#2ECC71","#3CBFEF","#E67E22",
    "#1ABC9C","#D35400","#2980B9","#7F8C8D",
]

YEAR_ORDER = ["1st Year","2nd Year","3rd Year","4th Year",
              "5th Year","6th Year","Other","Unknown"]

# ── Data layer ────────────────────────────────────────────────────────────────

def normalise_email(e: str) -> str:
    return str(e).strip().lower()


def detect_col(cols: list, keywords: list):
    """Return first column whose lowercased name contains any keyword."""
    cl = [c.lower() for c in cols]
    for kw in keywords:
        for i, c in enumerate(cl):
            if kw in c:
                return cols[i]
    return None


def resolve_name(row, cm: dict) -> str:
    """
    Build a full name from a CSV row.
    Priority: first+last → first-only → last-only → full-name column → "".
    """
    fcol = cm.get("first")
    lcol = cm.get("last")
    ncol = cm.get("name")

    first = str(row.get(fcol, "")).strip() if fcol else ""
    last  = str(row.get(lcol, "")).strip() if lcol else ""
    full  = str(row.get(ncol, "")).strip() if ncol else ""

    # Sanitise nan
    first = "" if first in ("nan","None") else first
    last  = "" if last  in ("nan","None") else last
    full  = "" if full  in ("nan","None") else full

    if first and last:   return f"{first} {last}"
    if first:            return first
    if last:             return last
    if full:             return full
    return ""


def build_master(file_map: dict, col_map: dict) -> pd.DataFrame:
    """
    Long-format master table: one row per (email, workshop).
    col_map shape: {ws_name: {email, first, last, name, year}}
    """
    rows = []
    for ws, path in file_map.items():
        cm   = col_map.get(ws, {})
        ecol = cm.get("email")
        ycol = cm.get("year")
        if not ecol:
            continue
        try:
            df = pd.read_csv(path)
        except Exception as exc:
            print(f"[WARN] Could not read {path}: {exc}")
            continue
        for _, row in df.iterrows():
            email = normalise_email(row.get(ecol, ""))
            if not email or email == "nan":
                continue
            name = resolve_name(row, cm)
            year = str(row.get(ycol, "")).strip() if ycol else "Unknown"
            if year in ("nan","None",""): year = "Unknown"
            rows.append({"email": email, "name": name,
                         "year": year, "workshop": ws})
    if not rows:
        return pd.DataFrame(columns=["email","name","year","workshop"])
    return pd.DataFrame(rows)


def build_profiles(master: pd.DataFrame) -> pd.DataFrame:
    if master.empty:
        return pd.DataFrame()
    grp = (
        master.groupby("email", sort=False)
        .agg(
            name     =("name",     lambda x: next((v for v in x if v), "")),
            year     =("year",     lambda x: next(
                           (v for v in x if v not in ("","Unknown","nan")), "Unknown")),
            workshops=("workshop", lambda x: sorted(set(x))),
        )
        .reset_index()
    )
    grp["n_workshops"] = grp["workshops"].apply(len)
    return grp


def build_allocation(profiles: pd.DataFrame, master: pd.DataFrame) -> pd.DataFrame:
    """
    For every attendee in 2+ workshops: suggest the one with the fewest registrants.
    Works for any number of workshops (tested up to 11).
    """
    if master.empty or profiles.empty:
        return pd.DataFrame()
    ws_counts = master.groupby("workshop")["email"].nunique().to_dict()
    multi = profiles[profiles["n_workshops"] > 1].copy()
    if multi.empty:
        return pd.DataFrame()
    out = []
    for _, row in multi.iterrows():
        wss  = row["workshops"]
        # Sort all workshops by current count, ascending
        ranked = sorted(wss, key=lambda w: ws_counts.get(w, 0))
        best   = ranked[0]
        others = ranked[1:]
        out.append({
            "email":            row["email"],
            "name":             row["name"],
            "year":             row["year"],
            "n_registered":     len(wss),
            "confirm_workshop": best,
            "workshop_count":   ws_counts.get(best, 0),
            "other_workshops":  " | ".join(others),
        })
    return pd.DataFrame(out).sort_values(["n_registered","workshop_count"],
                                          ascending=[False, True])


# ── Matplotlib style helper ───────────────────────────────────────────────────

def styled_axes(axes_list):
    for ax in axes_list:
        ax.set_facecolor(PANEL)
        ax.tick_params(colors=SUB, labelsize=8)
        ax.xaxis.label.set_color(SUB)
        ax.yaxis.label.set_color(SUB)
        for sp in ax.spines.values():
            sp.set_edgecolor("#2a2a4a")


# ══════════════════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CAMeSM · Workshop Registration Manager")
        self.geometry("1340x860")
        self.minsize(1100, 720)
        self.configure(fg_color=BG)

        self.file_map    = {}   # {ws_name: filepath}
        self.col_map     = {}   # {ws_name: {email, first, last, name, year}}
        self.master_df   = pd.DataFrame()
        self.profiles_df = pd.DataFrame()
        self.alloc_df    = pd.DataFrame()
        self._col_widgets = {}  # {ws_name: {field: StringVar}}

        self._build_ui()

    # ── Layout skeleton ───────────────────────────────────────────────────────

    def _build_ui(self):
        # Sidebar
        sb = ctk.CTkFrame(self, width=222, corner_radius=0, fg_color=PANEL)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)
        self.sidebar = sb

        ctk.CTkLabel(sb, text="🩺", font=("Arial", 36)).pack(pady=(26, 2))
        ctk.CTkLabel(sb, text="CAMeSM",
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=BLUE_VL).pack()
        ctk.CTkLabel(sb, text="Workshop Manager",
                     font=ctk.CTkFont(size=11), text_color=SUB).pack(pady=(0, 22))

        self._nav_btns = {}
        for label, icon in [("Load CSVs","📂"), ("Dashboard","📊"),
                             ("Profiles","👤"),  ("Allocation","🎯"),
                             ("Export","📤")]:
            btn = ctk.CTkButton(
                sb, text=f"  {icon}  {label}", anchor="w",
                corner_radius=8, height=40, fg_color="transparent",
                hover_color=CARD, text_color=TEXT, font=ctk.CTkFont(size=13),
                command=lambda l=label: self._switch(l),
            )
            btn.pack(fill="x", padx=12, pady=3)
            self._nav_btns[label] = btn

        self.status_lbl = ctk.CTkLabel(
            sb, text="No data loaded",
            font=ctk.CTkFont(size=10), text_color=SUB, wraplength=190)
        self.status_lbl.pack(side="bottom", pady=16, padx=12)

        # Content area
        self.content = ctk.CTkFrame(self, corner_radius=0, fg_color=BG)
        self.content.pack(side="left", fill="both", expand=True)

        self._panels = {}
        self._build_load_panel()
        self._build_dashboard_panel()
        self._build_profiles_panel()
        self._build_alloc_panel()
        self._build_export_panel()
        self._switch("Load CSVs")

    def _switch(self, name: str):
        for p in self._panels.values():
            p.pack_forget()
        self._panels[name].pack(fill="both", expand=True)
        for n, b in self._nav_btns.items():
            b.configure(
                fg_color=CARD if n == name else "transparent",
                text_color=BLUE_VL if n == name else TEXT,
            )
        if name == "Dashboard"  and not self.master_df.empty:   self._refresh_dashboard()
        if name == "Allocation" and not self.alloc_df.empty:    self._refresh_alloc()
        if name == "Export":                                     self._refresh_export()
        if name == "Profiles"   and not self.profiles_df.empty: self._filter_profiles()

    def _header(self, parent, title, subtitle=""):
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.pack(fill="x", padx=28, pady=(22, 8))
        ctk.CTkLabel(hdr, text=title,
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT).pack(anchor="w")
        if subtitle:
            ctk.CTkLabel(hdr, text=subtitle,
                         font=ctk.CTkFont(size=12), text_color=SUB).pack(anchor="w")

    # ══════════════════════════════════════════════════════════════════════════
    # 1. LOAD CSVs
    # ══════════════════════════════════════════════════════════════════════════

    def _build_load_panel(self):
        panel = ctk.CTkFrame(self.content, fg_color=BG)
        self._panels["Load CSVs"] = panel

        self._header(panel, "Load Workshop CSVs",
                     "One CSV per workshop · filename = workshop name · "
                     "first + last name columns detected automatically.")

        zone = ctk.CTkFrame(panel, fg_color=CARD, corner_radius=12, height=88)
        zone.pack(fill="x", padx=28, pady=(0, 14))
        zone.pack_propagate(False)
        inner = ctk.CTkFrame(zone, fg_color="transparent")
        inner.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(inner, text="📂  Select one CSV per workshop",
                     font=ctk.CTkFont(size=13), text_color=SUB).pack(side="left", padx=8)
        ctk.CTkButton(inner, text="Browse files…", width=140, height=36,
                      fg_color=BLUE, hover_color=BLUE_LT,
                      command=self._browse_csvs).pack(side="left", padx=8)

        self.col_scroll = ctk.CTkScrollableFrame(panel, fg_color=PANEL, corner_radius=10)
        self.col_scroll.pack(fill="both", expand=True, padx=28, pady=(0, 12))

        ctk.CTkButton(panel, text="⚙️  Build / Refresh Data", height=44,
                      font=ctk.CTkFont(size=14, weight="bold"),
                      fg_color=BLUE, hover_color=BLUE_LT,
                      command=self._build_data).pack(padx=28, pady=(0, 22))

    def _browse_csvs(self):
        paths = filedialog.askopenfilenames(
            title="Select workshop CSVs",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        for p in paths:
            ws = Path(p).stem
            self.file_map[ws] = p
        self._render_col_mapping()

    def _render_col_mapping(self):
        for w in self.col_scroll.winfo_children():
            w.destroy()
        self._col_widgets = {}

        for ws, path in self.file_map.items():
            try:
                cols = list(pd.read_csv(path, nrows=1).columns)
            except Exception:
                cols = []

            card = ctk.CTkFrame(self.col_scroll, fg_color=CARD, corner_radius=10)
            card.pack(fill="x", pady=5, padx=4)

            top = ctk.CTkFrame(card, fg_color="transparent")
            top.pack(fill="x", padx=14, pady=(10, 4))
            ctk.CTkLabel(top, text=f"📋  {ws}",
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=BLUE_VL).pack(side="left")
            ctk.CTkLabel(top, text=Path(path).name,
                         font=ctk.CTkFont(size=10), text_color=SUB).pack(side="right")

            row = ctk.CTkFrame(card, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=(0, 10))
            blank   = ["— skip —"]
            options = blank + cols
            widgets = {}

            # Email, First Name, Last Name, Full Name (fallback), Year
            fields = [
                ("email", ["email","e-mail","mail"],                "Email *",      True),
                ("first", ["first","given","forename","prénom"],    "First Name",   False),
                ("last",  ["last","surname","family","nom"],        "Last Name",    False),
                ("name",  ["full","name","student"],                "Full Name ↩",  False),
                ("year",  ["year","semester","study","grade","yr"], "Year",         False),
            ]
            for field, kws, label, required in fields:
                col_f = ctk.CTkFrame(row, fg_color="transparent")
                col_f.pack(side="left", padx=(0, 12))
                lbl_text = label + (" (required)" if required else "")
                ctk.CTkLabel(col_f, text=lbl_text,
                             font=ctk.CTkFont(size=10),
                             text_color=SUB).pack(anchor="w")
                detected = detect_col(cols, kws)
                idx = options.index(detected) if detected in options else 0
                var = ctk.StringVar(value=options[idx])
                ctk.CTkComboBox(col_f, values=options, variable=var,
                                width=170, height=28,
                                fg_color=PANEL, border_color=BLUE,
                                button_color=BLUE).pack()
                widgets[field] = var

            self._col_widgets[ws] = widgets

    def _build_data(self):
        if not self.file_map:
            messagebox.showwarning("No files", "Load at least one CSV first.")
            return
        col_map = {
            ws: {k: (None if v.get() == "— skip —" else v.get())
                 for k, v in wids.items()}
            for ws, wids in self._col_widgets.items()
        }
        self.col_map     = col_map
        master           = build_master(self.file_map, col_map)
        profiles         = build_profiles(master)
        alloc            = build_allocation(profiles, master)
        self.master_df   = master
        self.profiles_df = profiles
        self.alloc_df    = alloc

        self._update_prof_filters()

        u = len(profiles)
        w = master["workshop"].nunique() if not master.empty else 0
        m = int((profiles["n_workshops"] > 1).sum()) if not profiles.empty else 0
        self.status_lbl.configure(
            text=f"✅ {u} attendees\n{w} workshops\n{m} multi-registered",
            text_color=GREEN,
        )
        messagebox.showinfo("Done",
            f"Data built successfully.\n\n"
            f"• {u} unique attendees\n"
            f"• {w} workshops\n"
            f"• {m} registered in 2+ workshops")
        self._switch("Dashboard")

    # ══════════════════════════════════════════════════════════════════════════
    # 2. DASHBOARD
    # ══════════════════════════════════════════════════════════════════════════

    def _build_dashboard_panel(self):
        panel = ctk.CTkScrollableFrame(self.content, fg_color=BG)
        self._panels["Dashboard"] = panel
        self._dash_panel = panel

    def _refresh_dashboard(self):
        for w in self._dash_panel.winfo_children():
            w.destroy()

        master   = self.master_df
        profiles = self.profiles_df
        self._header(self._dash_panel, "Dashboard",
                     "Registration overview at a glance.")

        # ── KPIs ──────────────────────────────────────────────────────────────
        ws_counts = master.groupby("workshop")["email"].nunique()
        kpi_row = ctk.CTkFrame(self._dash_panel, fg_color="transparent")
        kpi_row.pack(fill="x", padx=28, pady=(0, 14))
        for title, val, col in [
            ("Unique Attendees",    str(len(profiles)),                          BLUE_VL),
            ("Total Registrations", str(len(master)),                            BLUE_LT),
            ("Multi-Workshop",      str(int((profiles["n_workshops"]>1).sum())), AMBER),
            ("Avg / Workshop",      f"{ws_counts.mean():.1f}",                  GREEN),
            ("Workshops Loaded",    str(master["workshop"].nunique()),           BLUE),
        ]:
            c = ctk.CTkFrame(kpi_row, fg_color=CARD, corner_radius=12)
            c.pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkLabel(c, text=val,
                         font=ctk.CTkFont(size=26, weight="bold"),
                         text_color=col).pack(pady=(12, 2))
            ctk.CTkLabel(c, text=title,
                         font=ctk.CTkFont(size=10), text_color=SUB).pack(pady=(0, 10))

        n_ws = master["workshop"].nunique()
        # Scale figure height for many workshops
        bar_h = max(3.8, n_ws * 0.38)

        # ── Chart 1: Attendance bar ────────────────────────────────────────────
        fig1, ax1 = plt.subplots(figsize=(13, bar_h))
        fig1.patch.set_facecolor(PANEL)
        styled_axes([ax1])
        ws_s = ws_counts.sort_values()
        bars = ax1.barh(
            ws_s.index, ws_s.values,
            color=[PALETTE[i % len(PALETTE)] for i in range(len(ws_s))],
            edgecolor="none", height=0.55,
        )
        for bar, v in zip(bars, ws_s.values):
            ax1.text(v + 0.2, bar.get_y() + bar.get_height()/2,
                     str(v), va="center", color=TEXT, fontsize=9)
        ax1.set_title("Registrations per Workshop", color=TEXT, fontsize=11, pad=8)
        ax1.set_xlim(0, ws_s.max() * 1.18)
        ax1.invert_yaxis()
        fig1.tight_layout(pad=1.5)
        c1 = FigureCanvasTkAgg(fig1, master=self._dash_panel)
        c1.draw(); c1.get_tk_widget().pack(fill="x", padx=28, pady=(0, 6))

        # ── Charts 2: Year pie + workshops-per-attendee bar ────────────────────
        fig2, (ax2, ax3) = plt.subplots(1, 2, figsize=(13, 3.8))
        fig2.patch.set_facecolor(PANEL)
        styled_axes([ax2, ax3])

        yc = profiles["year"].value_counts()
        wedges, texts, autotexts = ax2.pie(
            yc.values, labels=yc.index, autopct="%1.0f%%", startangle=90,
            colors=[PALETTE[i % len(PALETTE)] for i in range(len(yc))],
            textprops={"color": TEXT, "fontsize": 8},
            wedgeprops={"edgecolor": PANEL, "linewidth": 2},
        )
        for at in autotexts:
            at.set_color(BG); at.set_fontsize(8)
        ax2.set_title("Year of Study", color=TEXT, fontsize=11, pad=8)

        nw = profiles["n_workshops"].value_counts().sort_index()
        # Colour: green=1, amber=2-4, red=5+
        bar_colours = [
            GREEN if k == 1 else AMBER if k <= 4 else RED
            for k in nw.index
        ]
        ax3.bar([str(k) for k in nw.index], nw.values,
                color=bar_colours, edgecolor="none", width=0.55)
        for i, (k, v) in enumerate(zip(nw.index, nw.values)):
            ax3.text(i, v + 0.15, str(v), ha="center", color=TEXT, fontsize=9)
        ax3.set_title("Workshops per Attendee", color=TEXT, fontsize=11, pad=8)
        ax3.set_xlabel("Number of Workshops Registered", color=SUB)
        ax3.set_ylabel("Attendees", color=SUB)
        ax3.set_ylim(0, nw.max() * 1.22)
        fig2.tight_layout(pad=2)
        c2 = FigureCanvasTkAgg(fig2, master=self._dash_panel)
        c2.draw(); c2.get_tk_widget().pack(fill="x", padx=28, pady=(0, 6))

        # ── Chart 3: Year × Workshop heatmap ──────────────────────────────────
        pivot = (
            master.groupby(["year","workshop"])["email"]
            .nunique().reset_index()
            .pivot(index="year", columns="workshop", values="email")
            .fillna(0).astype(int)
        )
        hm_h = max(3.2, len(pivot.index) * 0.45)
        fig3, ax4 = plt.subplots(figsize=(13, hm_h))
        fig3.patch.set_facecolor(PANEL)
        styled_axes([ax4])
        data = pivot.values
        ax4.imshow(data, cmap="Blues", aspect="auto")
        ax4.set_xticks(range(len(pivot.columns)))
        ax4.set_xticklabels(pivot.columns, rotation=30, ha="right",
                            color=TEXT, fontsize=8)
        ax4.set_yticks(range(len(pivot.index)))
        ax4.set_yticklabels(pivot.index, color=TEXT, fontsize=8)
        thresh = data.max() * 0.6 if data.max() > 0 else 1
        for i in range(data.shape[0]):
            for j in range(data.shape[1]):
                ax4.text(j, i, str(data[i, j]),
                         ha="center", va="center", fontsize=9, fontweight="bold",
                         color=BG if data[i, j] >= thresh else TEXT)
        ax4.set_title("Year × Workshop Heatmap", color=TEXT, fontsize=11, pad=8)
        fig3.tight_layout(pad=1.5)
        c3 = FigureCanvasTkAgg(fig3, master=self._dash_panel)
        c3.draw(); c3.get_tk_widget().pack(fill="x", padx=28, pady=(0, 24))

    # ══════════════════════════════════════════════════════════════════════════
    # 3. PROFILES
    # ══════════════════════════════════════════════════════════════════════════

    def _build_profiles_panel(self):
        panel = ctk.CTkFrame(self.content, fg_color=BG)
        self._panels["Profiles"] = panel

        self._header(panel, "Attendee Profiles",
                     "Search · filter by workshop, year, or number of workshops registered.")

        # ── Filter bar ────────────────────────────────────────────────────────
        bar = ctk.CTkFrame(panel, fg_color=PANEL, corner_radius=10)
        bar.pack(fill="x", padx=28, pady=(0, 10))

        # Row 1
        r1 = ctk.CTkFrame(bar, fg_color="transparent")
        r1.pack(fill="x", padx=12, pady=(10, 4))

        ctk.CTkLabel(r1, text="🔍", font=("Arial", 14)).pack(side="left", padx=(0, 4))
        self._psearch = ctk.CTkEntry(
            r1, placeholder_text="Search name or email…",
            width=280, height=34, fg_color=CARD, border_color=BLUE)
        self._psearch.pack(side="left", padx=(0, 10))
        self._psearch.bind("<KeyRelease>", lambda e: self._filter_profiles())

        ctk.CTkLabel(r1, text="Workshop:", text_color=SUB,
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 4))
        self._pws_var = ctk.StringVar(value="All")
        self._pws_cb  = ctk.CTkComboBox(
            r1, variable=self._pws_var, values=["All"], width=200, height=34,
            fg_color=CARD, border_color=BLUE, button_color=BLUE,
            command=lambda _: self._filter_profiles())
        self._pws_cb.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(r1, text="Year:", text_color=SUB,
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 4))
        self._pyr_var = ctk.StringVar(value="All")
        ctk.CTkComboBox(
            r1, variable=self._pyr_var, values=["All"] + YEAR_ORDER,
            width=130, height=34, fg_color=CARD, border_color=BLUE,
            button_color=BLUE,
            command=lambda _: self._filter_profiles()).pack(side="left", padx=(0, 10))

        # Row 2
        r2 = ctk.CTkFrame(bar, fg_color="transparent")
        r2.pack(fill="x", padx=12, pady=(0, 10))

        ctk.CTkLabel(r2, text="Min workshops registered:",
                     text_color=SUB, font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
        self._pmin_var = ctk.StringVar(value="Any")
        self._pmin_opts = ["Any", "2+", "3+", "4+", "5+", "6+", "7+", "8+", "9+", "10+"]
        ctk.CTkSegmentedButton(
            r2, values=self._pmin_opts, variable=self._pmin_var,
            font=ctk.CTkFont(size=11),
            command=lambda _: self._filter_profiles()
        ).pack(side="left", padx=(0, 16))

        self._pcount = ctk.CTkLabel(r2, text="", text_color=SUB,
                                     font=ctk.CTkFont(size=11))
        self._pcount.pack(side="right")

        # ── Card scroll area ──────────────────────────────────────────────────
        self._pscroll = ctk.CTkScrollableFrame(panel, fg_color=BG)
        self._pscroll.pack(fill="both", expand=True, padx=28, pady=(0, 14))

    def _update_prof_filters(self):
        if self.master_df.empty:
            return
        wss = sorted(self.master_df["workshop"].unique().tolist())
        self._pws_cb.configure(values=["All"] + wss)
        self._pws_var.set("All")
        self._pyr_var.set("All")
        self._pmin_var.set("Any")
        self._filter_profiles()

    def _filter_profiles(self, *_):
        if self.profiles_df.empty:
            return
        q    = self._psearch.get().strip().lower()
        ws_q = self._pws_var.get()
        yr_q = self._pyr_var.get()
        min_q = self._pmin_var.get()

        view = self.profiles_df.copy()
        if q:
            view = view[
                view["name"].str.lower().str.contains(q, na=False) |
                view["email"].str.contains(q, na=False)
            ]
        if ws_q not in ("All", ""):
            view = view[view["workshops"].apply(lambda ws: ws_q in ws)]
        if yr_q not in ("All", ""):
            view = view[view["year"] == yr_q]
        if min_q != "Any":
            threshold = int(min_q.replace("+", ""))
            view = view[view["n_workshops"] >= threshold]

        self._pcount.configure(
            text=f"{len(view)} of {len(self.profiles_df)} attendees")
        self._render_cards(view)

    def _render_cards(self, df: pd.DataFrame):
        for w in self._pscroll.winfo_children():
            w.destroy()
        if df.empty:
            ctk.CTkLabel(self._pscroll,
                         text="No attendees match the current filters.",
                         text_color=SUB).pack(pady=40)
            return

        COLS = 3
        for chunk_start in range(0, len(df), COLS):
            chunk = df.iloc[chunk_start:chunk_start + COLS]
            row_f = ctk.CTkFrame(self._pscroll, fg_color="transparent")
            row_f.pack(fill="x", pady=4)
            for _, r in chunk.iterrows():
                n      = r["n_workshops"]
                multi  = n > 1
                border = RED if n >= 5 else AMBER if n >= 2 else "#2a2a5a"
                card   = ctk.CTkFrame(row_f, fg_color=CARD, corner_radius=12,
                                      border_width=2 if multi else 1,
                                      border_color=border)
                card.pack(side="left", expand=True, fill="x", padx=5, anchor="n")

                ctk.CTkLabel(card, text=r["name"] or "—",
                             font=ctk.CTkFont(size=13, weight="bold"),
                             text_color=BLUE_VL).pack(anchor="w", padx=12, pady=(10, 0))
                ctk.CTkLabel(card, text=r["email"],
                             font=ctk.CTkFont(size=10), text_color=SUB).pack(anchor="w", padx=12)
                ctk.CTkLabel(card, text=r["year"],
                             font=ctk.CTkFont(size=10), text_color=SUB).pack(anchor="w", padx=12, pady=(0, 4))

                # Tag — colour-coded by severity
                tag_text = f"⚡ {n} workshops" if multi else "✓ 1 workshop"
                tag_col  = RED if n >= 5 else AMBER if n >= 2 else GREEN
                ctk.CTkLabel(card, text=tag_text,
                             font=ctk.CTkFont(size=10, weight="bold"),
                             text_color=tag_col).pack(anchor="w", padx=12, pady=(0, 6))

                # Badges — wrap every 3 per row to handle many workshops
                workshops = r["workshops"]
                BADGES_PER_ROW = 3
                for row_start in range(0, len(workshops), BADGES_PER_ROW):
                    badge_row = ctk.CTkFrame(card, fg_color="transparent")
                    badge_row.pack(anchor="w", padx=10, pady=1)
                    for j, ws in enumerate(workshops[row_start:row_start + BADGES_PER_ROW]):
                        colour_idx = (workshops.index(ws)) % len(PALETTE)
                        ctk.CTkLabel(
                            badge_row,
                            text=f"  {ws[:20]}{'…' if len(ws)>20 else ''}  ",
                            font=ctk.CTkFont(size=8),
                            fg_color=PALETTE[colour_idx],
                            corner_radius=8, text_color="white",
                        ).pack(side="left", padx=2)
                # Bottom padding
                ctk.CTkFrame(card, fg_color="transparent", height=8).pack()

    # ══════════════════════════════════════════════════════════════════════════
    # 4. ALLOCATION HELPER
    # ══════════════════════════════════════════════════════════════════════════

    def _build_alloc_panel(self):
        panel = ctk.CTkFrame(self.content, fg_color=BG)
        self._panels["Allocation"] = panel
        self._alloc_outer = panel

    def _refresh_alloc(self):
        for w in self._alloc_outer.winfo_children():
            w.destroy()
        self._header(self._alloc_outer, "Allocation Helper",
                     "Attendees in 2+ workshops · sorted by most-registered first · "
                     "suggested confirmation = least-subscribed workshop.")

        alloc = self.alloc_df
        if alloc.empty:
            ctk.CTkLabel(self._alloc_outer,
                         text="🎉  No attendee is registered in more than one workshop.",
                         font=ctk.CTkFont(size=14), text_color=GREEN).pack(pady=60)
            return

        # KPIs
        ws_counts = self.master_df.groupby("workshop")["email"].nunique()
        max_registered = alloc["n_registered"].max()
        krow = ctk.CTkFrame(self._alloc_outer, fg_color="transparent")
        krow.pack(fill="x", padx=28, pady=(0, 12))
        for title, val, col in [
            ("Attendees to Decide",     str(len(alloc)),               AMBER),
            ("Most Subscribed WS",      ws_counts.idxmax(),            RED),
            ("Least Subscribed WS",     ws_counts.idxmin(),            GREEN),
            ("Max WS per Attendee",     str(max_registered),           RED if max_registered>=5 else AMBER),
        ]:
            c = ctk.CTkFrame(krow, fg_color=CARD, corner_radius=12)
            c.pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkLabel(c, text=val,
                         font=ctk.CTkFont(size=14, weight="bold"),
                         text_color=col, wraplength=200).pack(pady=(12, 2))
            ctk.CTkLabel(c, text=title,
                         font=ctk.CTkFont(size=10), text_color=SUB).pack(pady=(0, 10))

        # Filter bar
        bar = ctk.CTkFrame(self._alloc_outer, fg_color=PANEL, corner_radius=8)
        bar.pack(fill="x", padx=28, pady=(0, 8))

        ctk.CTkLabel(bar, text="Confirm in:", text_color=SUB,
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=12, pady=8)
        ws_opts = ["All"] + sorted(alloc["confirm_workshop"].unique())
        ws_var  = ctk.StringVar(value="All")

        ctk.CTkLabel(bar, text="Registered in ≥:", text_color=SUB,
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(20, 4), pady=8)
        n_opts = ["Any"] + [f"{i}+" for i in range(2, max_registered + 1)]
        n_var  = ctk.StringVar(value="Any")

        table_frame = ctk.CTkFrame(self._alloc_outer, fg_color="transparent")
        table_frame.pack(fill="both", expand=True, padx=28, pady=(0, 16))
        scroll_holder = [None]

        def refresh_table(*_):
            if scroll_holder[0]:
                scroll_holder[0].destroy()
            flt = alloc.copy()
            if ws_var.get() != "All":
                flt = flt[flt["confirm_workshop"] == ws_var.get()]
            if n_var.get() != "Any":
                n_th = int(n_var.get().replace("+",""))
                flt  = flt[flt["n_registered"] >= n_th]

            sc = ctk.CTkScrollableFrame(table_frame, fg_color=PANEL, corner_radius=10)
            sc.pack(fill="both", expand=True)
            scroll_holder[0] = sc

            # Header
            hdr = ctk.CTkFrame(sc, fg_color="transparent")
            hdr.pack(fill="x", padx=4, pady=(4, 0))
            for col_name, w in [("Email",210), ("Name",160), ("Year",90),
                                 ("#",45), ("✅ Confirm in",190),
                                 ("Other workshops",340)]:
                ctk.CTkLabel(hdr, text=col_name, width=w,
                             font=ctk.CTkFont(size=10, weight="bold"),
                             text_color=SUB, anchor="w").pack(side="left", padx=4)

            # Rows
            for _, r in flt.iterrows():
                n = r["n_registered"]
                row_col = "#1a1a3a" if _ % 2 == 0 else CARD
                rw = ctk.CTkFrame(sc, fg_color=CARD, corner_radius=8)
                rw.pack(fill="x", padx=4, pady=3)
                n_col = RED if n >= 5 else AMBER if n >= 2 else GREEN
                for val, w, tc in [
                    (r["email"],            210, TEXT),
                    (r["name"] or "—",      160, TEXT),
                    (r["year"],              90, TEXT),
                    (str(n),                 45, n_col),
                    (r["confirm_workshop"], 190, GREEN),
                    # Truncate long "other workshops" string if extreme
                    (r["other_workshops"][:60] + ("…" if len(r["other_workshops"])>60 else ""),
                     340, SUB),
                ]:
                    ctk.CTkLabel(rw, text=val, width=w,
                                 font=ctk.CTkFont(size=10),
                                 text_color=tc, anchor="w").pack(side="left", padx=4, pady=6)

        ctk.CTkComboBox(bar, values=ws_opts, variable=ws_var, width=210, height=32,
                        fg_color=CARD, border_color=BLUE, button_color=BLUE,
                        command=refresh_table).pack(side="left", padx=4, pady=6)

        ctk.CTkComboBox(bar, values=n_opts, variable=n_var, width=100, height=32,
                        fg_color=CARD, border_color=BLUE, button_color=BLUE,
                        command=refresh_table).pack(side="left", padx=4, pady=6)

        refresh_table()

    # ══════════════════════════════════════════════════════════════════════════
    # 5. EXPORT
    # ══════════════════════════════════════════════════════════════════════════

    def _build_export_panel(self):
        panel = ctk.CTkScrollableFrame(self.content, fg_color=BG)
        self._panels["Export"] = panel
        self._export_outer = panel

    def _refresh_export(self):
        for w in self._export_outer.winfo_children():
            w.destroy()
        self._header(self._export_outer, "Export",
                     "Save CSVs ready for confirmation emails.")

        box = ctk.CTkFrame(self._export_outer, fg_color="transparent")
        box.pack(fill="x", padx=28)

        def export_card(label, desc, fn):
            card = ctk.CTkFrame(box, fg_color=CARD, corner_radius=12)
            card.pack(fill="x", pady=6)
            ctk.CTkLabel(card, text=label,
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=TEXT).pack(anchor="w", padx=16, pady=(12, 2))
            ctk.CTkLabel(card, text=desc,
                         font=ctk.CTkFont(size=10), text_color=SUB).pack(anchor="w", padx=16, pady=(0, 8))
            ctk.CTkButton(card, text="⬇️  Export", width=130, height=34,
                          fg_color=BLUE, hover_color=BLUE_LT,
                          command=fn).pack(anchor="w", padx=16, pady=(0, 12))

        export_card("All Profiles",
                    "One row per unique attendee — all workshops listed, first+last name combined.",
                    self._export_profiles)
        export_card("Allocation Suggestions",
                    "Multi-registered attendees with their suggested confirmation workshop.",
                    self._export_alloc)
        export_card("Master Registration Log",
                    "Every registration entry in long format (email × workshop).",
                    self._export_master)

        alloc = self.alloc_df
        if not alloc.empty:
            ctk.CTkLabel(box, text="Per-Workshop Confirmation Lists",
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=TEXT).pack(anchor="w", pady=(16, 2))
            ctk.CTkLabel(box,
                         text="One CSV per workshop — the exact emails to confirm for that slot.",
                         font=ctk.CTkFont(size=10), text_color=SUB).pack(anchor="w", pady=(0, 8))

            # Wrap buttons across rows of 4
            ws_list = sorted(alloc["confirm_workshop"].unique())
            BTNS_PER_ROW = 4
            for i in range(0, len(ws_list), BTNS_PER_ROW):
                btn_row = ctk.CTkFrame(box, fg_color="transparent")
                btn_row.pack(fill="x", pady=3)
                for ws in ws_list[i:i + BTNS_PER_ROW]:
                    ctk.CTkButton(
                        btn_row,
                        text=f"⬇️  {ws[:22]}{'…' if len(ws)>22 else ''}",
                        height=34, fg_color=PANEL, hover_color=CARD,
                        border_color=BLUE, border_width=1,
                        command=lambda w=ws: self._export_ws_confirm(w)
                    ).pack(side="left", padx=4)

    def _save_csv(self, df, default_name):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile=default_name,
        )
        if path:
            df.to_csv(path, index=False)
            messagebox.showinfo("Saved", f"Saved to:\n{path}")

    def _export_profiles(self):
        if self.profiles_df.empty:
            messagebox.showwarning("No data", "Build data first.")
            return
        df = self.profiles_df.copy()
        df["workshops"] = df["workshops"].apply(lambda ws: " | ".join(ws))
        self._save_csv(df, "camasm_all_profiles.csv")

    def _export_alloc(self):
        if self.alloc_df.empty:
            messagebox.showinfo("Nothing", "No multi-registered attendees.")
            return
        self._save_csv(self.alloc_df, "camasm_allocation_suggestions.csv")

    def _export_master(self):
        if self.master_df.empty:
            messagebox.showwarning("No data", "Build data first.")
            return
        self._save_csv(self.master_df, "camasm_master_log.csv")

    def _export_ws_confirm(self, ws):
        sub = self.alloc_df[self.alloc_df["confirm_workshop"] == ws][
            ["email", "name", "year", "other_workshops"]
        ]
        self._save_csv(sub, f"confirm_{ws.replace(' ', '_')}.csv")


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
