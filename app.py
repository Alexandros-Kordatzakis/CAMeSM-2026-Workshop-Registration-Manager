"""
CAMeSM Workshop Registration Manager
=====================================
Desktop GUI — CustomTkinter + Matplotlib
One CSV per workshop → email-anchored attendee profiles + smart allocation.

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
import os

# ── Theme ─────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BLUE       = "#1A6FA3"
BLUE_LT    = "#2596D4"
BLUE_VLT   = "#3CBFEF"
AMBER      = "#F0A500"
GREEN      = "#27AE60"
RED        = "#C0392B"
BG         = "#1a1a2e"
PANEL      = "#16213e"
CARD       = "#0f3460"
TEXT       = "#e0e0e0"
SUBTEXT    = "#9ea3b0"

PALETTE = [
    "#1A6FA3","#2596D4","#3CBFEF","#5DD4F8",
    "#F0A500","#E05C2A","#9B59B6","#27AE60",
    "#C0392B","#16A085","#8E44AD","#2ECC71",
]

YEAR_ORDER = ["1st Year","2nd Year","3rd Year","4th Year",
              "5th Year","6th Year","Other","Unknown"]

# ── Data layer ────────────────────────────────────────────────────────────────

def normalise_email(e: str) -> str:
    return str(e).strip().lower()


def detect_col(cols: list, keywords: list):
    cl = [c.lower() for c in cols]
    for kw in keywords:
        for i, c in enumerate(cl):
            if kw in c:
                return cols[i]
    return None


def build_master(file_map: dict, col_map: dict) -> pd.DataFrame:
    rows = []
    for ws, path in file_map.items():
        cm   = col_map.get(ws, {})
        ecol = cm.get("email")
        ncol = cm.get("name")
        ycol = cm.get("year")
        if not ecol:
            continue
        try:
            df = pd.read_csv(path)
        except Exception:
            continue
        for _, row in df.iterrows():
            email = normalise_email(row.get(ecol, ""))
            if not email or email == "nan":
                continue
            name = str(row.get(ncol, "")).strip() if ncol else ""
            year = str(row.get(ycol, "")).strip() if ycol else "Unknown"
            if name in ("nan", ""):  name = ""
            if year in ("nan", ""): year = "Unknown"
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
            year     =("year",     lambda x: next((v for v in x
                                    if v not in ("","Unknown","nan")), "Unknown")),
            workshops=("workshop", lambda x: sorted(set(x))),
        )
        .reset_index()
    )
    grp["n_workshops"] = grp["workshops"].apply(len)
    return grp


def build_allocation(profiles: pd.DataFrame, master: pd.DataFrame) -> pd.DataFrame:
    if master.empty or profiles.empty:
        return pd.DataFrame()
    ws_counts = master.groupby("workshop")["email"].nunique().to_dict()
    multi = profiles[profiles["n_workshops"] > 1].copy()
    if multi.empty:
        return pd.DataFrame()
    out = []
    for _, row in multi.iterrows():
        wss  = row["workshops"]
        best = min(wss, key=lambda w: ws_counts.get(w, 0))
        others = [w for w in wss if w != best]
        out.append({
            "email":            row["email"],
            "name":             row["name"],
            "year":             row["year"],
            "n_registered":     len(wss),
            "confirm_workshop": best,
            "workshop_count":   ws_counts.get(best, 0),
            "other_workshops":  " | ".join(others),
        })
    return pd.DataFrame(out).sort_values("workshop_count")


# ── Matplotlib helper ─────────────────────────────────────────────────────────

def make_fig(rows=1, cols=1, h=4):
    fig, axes = plt.subplots(rows, cols, figsize=(13, h))
    fig.patch.set_facecolor(PANEL)
    if rows == 1 and cols == 1:
        axes = [axes]
    elif rows == 1 or cols == 1:
        axes = list(axes)
    else:
        axes = [ax for r in axes for ax in r]
    for ax in axes:
        ax.set_facecolor(PANEL)
        ax.tick_params(colors=SUBTEXT, labelsize=8)
        ax.xaxis.label.set_color(SUBTEXT)
        ax.yaxis.label.set_color(SUBTEXT)
        for sp in ax.spines.values():
            sp.set_edgecolor("#2a2a4a")
    return fig, axes


# ══════════════════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CAMeSM · Workshop Registration Manager")
        self.geometry("1300x840")
        self.minsize(1100, 720)
        self.configure(fg_color=BG)

        self.file_map    = {}
        self.col_map     = {}
        self.master_df   = pd.DataFrame()
        self.profiles_df = pd.DataFrame()
        self.alloc_df    = pd.DataFrame()
        self._col_widgets = {}

        self._build_ui()

    # ── Layout skeleton ───────────────────────────────────────────────────────

    def _build_ui(self):
        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color=PANEL)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(self.sidebar, text="🩺", font=("Arial",36)).pack(pady=(26,2))
        ctk.CTkLabel(self.sidebar, text="CAMeSM",
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=BLUE_VLT).pack()
        ctk.CTkLabel(self.sidebar, text="Workshop Manager",
                     font=ctk.CTkFont(size=11), text_color=SUBTEXT).pack(pady=(0,22))

        self._nav_btns = {}
        for label, icon in [("Load CSVs","📂"),("Dashboard","📊"),
                             ("Profiles","👤"),("Allocation","🎯"),("Export","📤")]:
            btn = ctk.CTkButton(
                self.sidebar, text=f"  {icon}  {label}", anchor="w",
                corner_radius=8, height=40, fg_color="transparent",
                hover_color=CARD, text_color=TEXT, font=ctk.CTkFont(size=13),
                command=lambda l=label: self._switch(l),
            )
            btn.pack(fill="x", padx=12, pady=3)
            self._nav_btns[label] = btn

        self.status_lbl = ctk.CTkLabel(
            self.sidebar, text="No data loaded",
            font=ctk.CTkFont(size=10), text_color=SUBTEXT, wraplength=180)
        self.status_lbl.pack(side="bottom", pady=16, padx=12)

        # Content
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
                text_color=BLUE_VLT if n == name else TEXT,
            )
        if name == "Dashboard"  and not self.master_df.empty:   self._refresh_dashboard()
        if name == "Allocation" and not self.alloc_df.empty:    self._refresh_alloc()
        if name == "Export":                                     self._refresh_export()

    def _header(self, parent, title, subtitle=""):
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.pack(fill="x", padx=28, pady=(22,8))
        ctk.CTkLabel(hdr, text=title, font=ctk.CTkFont(size=22, weight="bold"),
                     text_color=TEXT).pack(anchor="w")
        if subtitle:
            ctk.CTkLabel(hdr, text=subtitle, font=ctk.CTkFont(size=12),
                         text_color=SUBTEXT).pack(anchor="w")

    # ── 1. LOAD CSVs ──────────────────────────────────────────────────────────

    def _build_load_panel(self):
        panel = ctk.CTkFrame(self.content, fg_color=BG)
        self._panels["Load CSVs"] = panel

        self._header(panel, "Load Workshop CSVs",
                     "Each file = one workshop · filename becomes the workshop name.")

        zone = ctk.CTkFrame(panel, fg_color=CARD, corner_radius=12, height=88)
        zone.pack(fill="x", padx=28, pady=(0,14))
        zone.pack_propagate(False)
        inner = ctk.CTkFrame(zone, fg_color="transparent")
        inner.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(inner, text="📂  Select one CSV per workshop",
                     font=ctk.CTkFont(size=13), text_color=SUBTEXT).pack(side="left", padx=8)
        ctk.CTkButton(inner, text="Browse files…", width=140, height=36,
                      fg_color=BLUE, hover_color=BLUE_LT,
                      command=self._browse_csvs).pack(side="left", padx=8)

        self.col_scroll = ctk.CTkScrollableFrame(panel, fg_color=PANEL, corner_radius=10)
        self.col_scroll.pack(fill="both", expand=True, padx=28, pady=(0,12))

        ctk.CTkButton(panel, text="⚙️  Build / Refresh Data", height=44,
                      font=ctk.CTkFont(size=14, weight="bold"),
                      fg_color=BLUE, hover_color=BLUE_LT,
                      command=self._build_data).pack(padx=28, pady=(0,22))

    def _browse_csvs(self):
        paths = filedialog.askopenfilenames(
            title="Select workshop CSVs",
            filetypes=[("CSV files","*.csv"), ("All files","*.*")],
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
            top.pack(fill="x", padx=14, pady=(10,4))
            ctk.CTkLabel(top, text=f"📋  {ws}",
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=BLUE_VLT).pack(side="left")
            ctk.CTkLabel(top, text=Path(path).name,
                         font=ctk.CTkFont(size=10), text_color=SUBTEXT).pack(side="right")

            row = ctk.CTkFrame(card, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=(0,10))
            blank   = ["— skip —"]
            options = blank + cols
            widgets = {}
            for field, kws, label in [
                ("email", ["email","e-mail","mail"],         "Email column *"),
                ("name",  ["name","full","student","first"], "Name column"),
                ("year",  ["year","semester","study"],       "Year column"),
            ]:
                col_f = ctk.CTkFrame(row, fg_color="transparent")
                col_f.pack(side="left", padx=(0,14))
                ctk.CTkLabel(col_f, text=label, font=ctk.CTkFont(size=11),
                             text_color=SUBTEXT).pack(anchor="w")
                detected = detect_col(cols, kws)
                idx = options.index(detected) if detected in options else 0
                var = ctk.StringVar(value=options[idx])
                ctk.CTkComboBox(col_f, values=options, variable=var,
                                width=180, height=30, fg_color=PANEL,
                                border_color=BLUE, button_color=BLUE).pack()
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
            f"• {u} unique attendees\n• {w} workshops\n• {m} in 2+ workshops")
        self._switch("Dashboard")

    # ── 2. DASHBOARD ─────────────────────────────────────────────────────────

    def _build_dashboard_panel(self):
        panel = ctk.CTkScrollableFrame(self.content, fg_color=BG)
        self._panels["Dashboard"] = panel
        self._dash_panel = panel

    def _refresh_dashboard(self):
        for w in self._dash_panel.winfo_children():
            w.destroy()

        master   = self.master_df
        profiles = self.profiles_df
        self._header(self._dash_panel, "Dashboard", "Registration overview at a glance.")

        # KPI row
        kpi_row = ctk.CTkFrame(self._dash_panel, fg_color="transparent")
        kpi_row.pack(fill="x", padx=28, pady=(0,14))
        ws_counts = master.groupby("workshop")["email"].nunique()
        for title, val, col in [
            ("Unique Attendees",    str(len(profiles)),                          BLUE_VLT),
            ("Total Registrations", str(len(master)),                            BLUE_LT),
            ("Multi-Workshop",      str(int((profiles["n_workshops"]>1).sum())), AMBER),
            ("Avg / Workshop",      f"{ws_counts.mean():.1f}",                  GREEN),
            ("Workshops Loaded",    str(master["workshop"].nunique()),           BLUE),
        ]:
            c = ctk.CTkFrame(kpi_row, fg_color=CARD, corner_radius=12)
            c.pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkLabel(c, text=val, font=ctk.CTkFont(size=26, weight="bold"),
                         text_color=col).pack(pady=(12,2))
            ctk.CTkLabel(c, text=title, font=ctk.CTkFont(size=10),
                         text_color=SUBTEXT).pack(pady=(0,10))

        # Charts row 1: attendance bar + year pie
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 4.0))
        for f in [fig]:
            f.patch.set_facecolor(PANEL)
        for ax in [ax1, ax2]:
            ax.set_facecolor(PANEL)
            ax.tick_params(colors=SUBTEXT, labelsize=8)
            for sp in ax.spines.values(): sp.set_edgecolor("#2a2a4a")

        ws_s = ws_counts.sort_values()
        bars = ax1.barh(ws_s.index, ws_s.values,
                        color=[PALETTE[i % len(PALETTE)] for i in range(len(ws_s))],
                        edgecolor="none", height=0.55)
        for bar, v in zip(bars, ws_s.values):
            ax1.text(v+0.3, bar.get_y()+bar.get_height()/2, str(v),
                     va="center", color=TEXT, fontsize=9)
        ax1.set_title("Registrations per Workshop", color=TEXT, fontsize=11, pad=8)
        ax1.set_xlim(0, ws_s.max()*1.2)
        ax1.xaxis.label.set_color(SUBTEXT)
        ax1.invert_yaxis()

        yc = profiles["year"].value_counts()
        wedges, texts, autotexts = ax2.pie(
            yc.values, labels=yc.index, autopct="%1.0f%%", startangle=90,
            colors=[PALETTE[i % len(PALETTE)] for i in range(len(yc))],
            textprops={"color": TEXT, "fontsize": 8},
            wedgeprops={"edgecolor": PANEL, "linewidth": 2},
        )
        for at in autotexts: at.set_color(BG); at.set_fontsize(8)
        ax2.set_title("Year of Study", color=TEXT, fontsize=11, pad=8)
        fig.tight_layout(pad=2)
        c1 = FigureCanvasTkAgg(fig, master=self._dash_panel)
        c1.draw(); c1.get_tk_widget().pack(fill="x", padx=28, pady=(0,8))

        # Charts row 2: heatmap + multi-ws bar
        import numpy as np
        fig2, (ax3, ax4) = plt.subplots(1, 2, figsize=(13, 3.6))
        for f in [fig2]:
            f.patch.set_facecolor(PANEL)
        for ax in [ax3, ax4]:
            ax.set_facecolor(PANEL)
            ax.tick_params(colors=SUBTEXT, labelsize=8)
            for sp in ax.spines.values(): sp.set_edgecolor("#2a2a4a")

        pivot = (master.groupby(["year","workshop"])["email"]
                 .nunique().reset_index()
                 .pivot(index="year", columns="workshop", values="email")
                 .fillna(0).astype(int))
        data = pivot.values
        ax3.imshow(data, cmap="Blues", aspect="auto")
        ax3.set_xticks(range(len(pivot.columns)))
        ax3.set_xticklabels(pivot.columns, rotation=28, ha="right",
                            color=TEXT, fontsize=8)
        ax3.set_yticks(range(len(pivot.index)))
        ax3.set_yticklabels(pivot.index, color=TEXT, fontsize=8)
        for i in range(data.shape[0]):
            for j in range(data.shape[1]):
                ax3.text(j, i, str(data[i,j]), ha="center", va="center",
                         color=TEXT if data[i,j] < data.max()*0.6 else BG,
                         fontsize=9, fontweight="bold")
        ax3.set_title("Year × Workshop Heatmap", color=TEXT, fontsize=11, pad=8)

        nw = profiles["n_workshops"].value_counts().sort_index()
        ax4.bar([str(k) for k in nw.index], nw.values,
                color=AMBER, edgecolor="none", width=0.5)
        for i,(k,v) in enumerate(zip(nw.index, nw.values)):
            ax4.text(i, v+0.15, str(v), ha="center", color=TEXT, fontsize=9)
        ax4.set_title("Workshops per Attendee", color=TEXT, fontsize=11, pad=8)
        ax4.set_xlabel("Number of Workshops", color=SUBTEXT)
        ax4.set_ylabel("Attendees", color=SUBTEXT)
        ax4.set_ylim(0, nw.max()*1.2)
        fig2.tight_layout(pad=2)
        c2 = FigureCanvasTkAgg(fig2, master=self._dash_panel)
        c2.draw(); c2.get_tk_widget().pack(fill="x", padx=28, pady=(0,24))

    # ── 3. PROFILES ──────────────────────────────────────────────────────────

    def _build_profiles_panel(self):
        panel = ctk.CTkFrame(self.content, fg_color=BG)
        self._panels["Profiles"] = panel

        self._header(panel, "Attendee Profiles",
                     "Search by name, email, workshop, or year.")

        # Filter bar
        bar = ctk.CTkFrame(panel, fg_color=PANEL, corner_radius=10)
        bar.pack(fill="x", padx=28, pady=(0,10))

        ctk.CTkLabel(bar, text="🔍", font=("Arial",15)).pack(side="left", padx=(12,2), pady=10)
        self._psearch = ctk.CTkEntry(bar, placeholder_text="Search name or email…",
                                     width=250, height=34, fg_color=CARD, border_color=BLUE)
        self._psearch.pack(side="left", padx=4)
        self._psearch.bind("<KeyRelease>", lambda e: self._filter_profiles())

        self._pws_var = ctk.StringVar(value="All workshops")
        self._pws_cb  = ctk.CTkComboBox(bar, variable=self._pws_var,
                                         values=["All workshops"], width=200, height=34,
                                         fg_color=CARD, border_color=BLUE, button_color=BLUE,
                                         command=lambda _: self._filter_profiles())
        self._pws_cb.pack(side="left", padx=4)

        self._pyr_var = ctk.StringVar(value="All years")
        ctk.CTkComboBox(bar, variable=self._pyr_var,
                        values=["All years"]+YEAR_ORDER, width=140, height=34,
                        fg_color=CARD, border_color=BLUE, button_color=BLUE,
                        command=lambda _: self._filter_profiles()).pack(side="left", padx=4)

        self._pmulti = ctk.CTkCheckBox(bar, text="Multi-workshop only", text_color=TEXT,
                                        command=self._filter_profiles)
        self._pmulti.pack(side="left", padx=10)

        self._pcount = ctk.CTkLabel(bar, text="", text_color=SUBTEXT,
                                    font=ctk.CTkFont(size=11))
        self._pcount.pack(side="right", padx=14)

        self._pscroll = ctk.CTkScrollableFrame(panel, fg_color=BG)
        self._pscroll.pack(fill="both", expand=True, padx=28, pady=(0,14))

    def _update_prof_filters(self):
        if self.master_df.empty:
            return
        wss = sorted(self.master_df["workshop"].unique().tolist())
        self._pws_cb.configure(values=["All workshops"]+wss)
        self._pws_var.set("All workshops")
        self._filter_profiles()

    def _filter_profiles(self, *_):
        if self.profiles_df.empty:
            return
        q     = self._psearch.get().strip().lower()
        ws_q  = self._pws_var.get()
        yr_q  = self._pyr_var.get()
        multi = self._pmulti.get()
        view  = self.profiles_df.copy()
        if q:
            view = view[view["name"].str.lower().str.contains(q,na=False) |
                        view["email"].str.contains(q,na=False)]
        if ws_q not in ("All workshops",""):
            view = view[view["workshops"].apply(lambda ws: ws_q in ws)]
        if yr_q not in ("All years",""):
            view = view[view["year"] == yr_q]
        if multi:
            view = view[view["n_workshops"] > 1]
        self._pcount.configure(text=f"{len(view)} of {len(self.profiles_df)} attendees")
        self._render_cards(view)

    def _render_cards(self, df: pd.DataFrame):
        for w in self._pscroll.winfo_children():
            w.destroy()
        if df.empty:
            ctk.CTkLabel(self._pscroll, text="No attendees match the current filters.",
                         text_color=SUBTEXT).pack(pady=40)
            return
        COLS = 3
        for chunk_start in range(0, len(df), COLS):
            chunk = df.iloc[chunk_start:chunk_start+COLS]
            row_f = ctk.CTkFrame(self._pscroll, fg_color="transparent")
            row_f.pack(fill="x", pady=4)
            for _, r in chunk.iterrows():
                multi  = r["n_workshops"] > 1
                border = AMBER if multi else "#2a2a5a"
                card   = ctk.CTkFrame(row_f, fg_color=CARD, corner_radius=12,
                                      border_width=2 if multi else 1,
                                      border_color=border)
                card.pack(side="left", expand=True, fill="x", padx=5, anchor="n")

                ctk.CTkLabel(card, text=r["name"] or "—",
                             font=ctk.CTkFont(size=13, weight="bold"),
                             text_color=BLUE_VLT).pack(anchor="w", padx=12, pady=(10,0))
                ctk.CTkLabel(card, text=r["email"],
                             font=ctk.CTkFont(size=10), text_color=SUBTEXT).pack(anchor="w", padx=12)
                ctk.CTkLabel(card, text=r["year"],
                             font=ctk.CTkFont(size=10), text_color=SUBTEXT).pack(anchor="w", padx=12, pady=(0,5))
                tag_text = f"⚡ {r['n_workshops']} workshops" if multi else "✓ 1 workshop"
                tag_col  = AMBER if multi else GREEN
                ctk.CTkLabel(card, text=tag_text,
                             font=ctk.CTkFont(size=10, weight="bold"),
                             text_color=tag_col).pack(anchor="w", padx=12, pady=(0,5))
                badge_row = ctk.CTkFrame(card, fg_color="transparent")
                badge_row.pack(anchor="w", padx=10, pady=(0,10))
                for i, ws in enumerate(r["workshops"]):
                    ctk.CTkLabel(badge_row, text=f"  {ws}  ",
                                 font=ctk.CTkFont(size=9),
                                 fg_color=PALETTE[i % len(PALETTE)],
                                 corner_radius=8, text_color="white").pack(side="left", padx=2)

    # ── 4. ALLOCATION ─────────────────────────────────────────────────────────

    def _build_alloc_panel(self):
        panel = ctk.CTkFrame(self.content, fg_color=BG)
        self._panels["Allocation"] = panel
        self._alloc_outer = panel

    def _refresh_alloc(self):
        for w in self._alloc_outer.winfo_children():
            w.destroy()
        self._header(self._alloc_outer, "Allocation Helper",
                     "Confirms multi-registered attendees into the least-subscribed workshop.")

        alloc = self.alloc_df
        if alloc.empty:
            ctk.CTkLabel(self._alloc_outer,
                         text="🎉  No attendee is registered in more than one workshop.",
                         font=ctk.CTkFont(size=14), text_color=GREEN).pack(pady=60)
            return

        # KPIs
        ws_counts = self.master_df.groupby("workshop")["email"].nunique()
        krow = ctk.CTkFrame(self._alloc_outer, fg_color="transparent")
        krow.pack(fill="x", padx=28, pady=(0,14))
        for title, val, col in [
            ("Attendees to Decide", str(len(alloc)),      AMBER),
            ("Most Subscribed",     ws_counts.idxmax(),   RED),
            ("Least Subscribed",    ws_counts.idxmin(),   GREEN),
        ]:
            c = ctk.CTkFrame(krow, fg_color=CARD, corner_radius=12)
            c.pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkLabel(c, text=val, font=ctk.CTkFont(size=15, weight="bold"),
                         text_color=col, wraplength=190).pack(pady=(12,2))
            ctk.CTkLabel(c, text=title, font=ctk.CTkFont(size=10),
                         text_color=SUBTEXT).pack(pady=(0,10))

        # Workshop filter
        bar = ctk.CTkFrame(self._alloc_outer, fg_color=PANEL, corner_radius=8)
        bar.pack(fill="x", padx=28, pady=(0,8))
        ctk.CTkLabel(bar, text="Filter — confirm in:", text_color=SUBTEXT,
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=12, pady=8)
        ws_opts = ["All"] + sorted(alloc["confirm_workshop"].unique())
        ws_var  = ctk.StringVar(value="All")

        # Table container
        table_frame = ctk.CTkFrame(self._alloc_outer, fg_color="transparent")
        table_frame.pack(fill="both", expand=True, padx=28, pady=(0,16))
        scroll_holder = [None]

        def refresh_table(*_):
            if scroll_holder[0]:
                scroll_holder[0].destroy()
            flt = alloc if ws_var.get() == "All" else \
                  alloc[alloc["confirm_workshop"] == ws_var.get()]
            sc = ctk.CTkScrollableFrame(table_frame, fg_color=PANEL, corner_radius=10)
            sc.pack(fill="both", expand=True)
            scroll_holder[0] = sc

            # Header
            hdr = ctk.CTkFrame(sc, fg_color="transparent")
            hdr.pack(fill="x", padx=4, pady=(4,0))
            for col_name, w in [("Email",220),("Name",160),("Year",90),
                                 ("# WS",50),("✅ Confirm in",180),("Also registered in",300)]:
                ctk.CTkLabel(hdr, text=col_name, width=w,
                             font=ctk.CTkFont(size=10, weight="bold"),
                             text_color=SUBTEXT, anchor="w").pack(side="left", padx=4)

            for _, r in flt.iterrows():
                rw = ctk.CTkFrame(sc, fg_color=CARD, corner_radius=8)
                rw.pack(fill="x", padx=4, pady=3)
                for val, w, highlight in [
                    (r["email"],            220, False),
                    (r["name"] or "—",      160, False),
                    (r["year"],              90, False),
                    (str(r["n_registered"]), 50, False),
                    (r["confirm_workshop"], 180, True),
                    (r["other_workshops"],  300, False),
                ]:
                    ctk.CTkLabel(rw, text=val, width=w,
                                 font=ctk.CTkFont(size=10),
                                 text_color=GREEN if highlight else TEXT,
                                 anchor="w").pack(side="left", padx=4, pady=6)

        ctk.CTkComboBox(bar, values=ws_opts, variable=ws_var, width=220, height=32,
                        fg_color=CARD, border_color=BLUE, button_color=BLUE,
                        command=refresh_table).pack(side="left", padx=4, pady=6)
        refresh_table()

    # ── 5. EXPORT ─────────────────────────────────────────────────────────────

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
            ctk.CTkLabel(card, text=label, font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=TEXT).pack(anchor="w", padx=16, pady=(12,2))
            ctk.CTkLabel(card, text=desc, font=ctk.CTkFont(size=10),
                         text_color=SUBTEXT).pack(anchor="w", padx=16, pady=(0,8))
            ctk.CTkButton(card, text="⬇️  Export", width=130, height=34,
                          fg_color=BLUE, hover_color=BLUE_LT,
                          command=fn).pack(anchor="w", padx=16, pady=(0,12))

        export_card("All Profiles",
                    "One row per unique attendee with all workshops listed.",
                    self._export_profiles)
        export_card("Allocation Suggestions",
                    "Multi-registered attendees with their suggested confirmation workshop.",
                    self._export_alloc)
        export_card("Master Registration Log",
                    "Every registration entry in long format (email × workshop).",
                    self._export_master)

        if not self.alloc_df.empty:
            ctk.CTkLabel(box, text="Per-Workshop Confirmation Lists",
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=TEXT).pack(anchor="w", pady=(14,2))
            ctk.CTkLabel(box, text="One CSV per workshop — emails confirmed for that slot.",
                         font=ctk.CTkFont(size=10), text_color=SUBTEXT).pack(anchor="w", pady=(0,6))
            btn_row = ctk.CTkFrame(box, fg_color="transparent")
            btn_row.pack(fill="x")
            for ws in sorted(self.alloc_df["confirm_workshop"].unique()):
                ctk.CTkButton(btn_row, text=f"⬇️  {ws}", height=34,
                              fg_color=PANEL, hover_color=CARD,
                              border_color=BLUE, border_width=1,
                              command=lambda w=ws: self._export_ws_confirm(w)
                              ).pack(side="left", padx=4, pady=4)

    def _save_csv(self, df, default_name):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV","*.csv")],
            initialfile=default_name,
        )
        if path:
            df.to_csv(path, index=False)
            messagebox.showinfo("Saved", f"Saved to:\n{path}")

    def _export_profiles(self):
        if self.profiles_df.empty:
            messagebox.showwarning("No data","Build data first."); return
        df = self.profiles_df.copy()
        df["workshops"] = df["workshops"].apply(lambda ws: " | ".join(ws))
        self._save_csv(df, "camasm_all_profiles.csv")

    def _export_alloc(self):
        if self.alloc_df.empty:
            messagebox.showinfo("Nothing","No multi-registered attendees."); return
        self._save_csv(self.alloc_df, "camasm_allocation_suggestions.csv")

    def _export_master(self):
        if self.master_df.empty:
            messagebox.showwarning("No data","Build data first."); return
        self._save_csv(self.master_df, "camasm_master_log.csv")

    def _export_ws_confirm(self, ws):
        sub = self.alloc_df[self.alloc_df["confirm_workshop"] == ws][
            ["email","name","year","other_workshops"]
        ]
        self._save_csv(sub, f"confirm_{ws.replace(' ','_')}.csv")


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
