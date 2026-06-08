"""
Nissan Variance Analysis Tool
"""

import os
import sys
import shutil
import threading
import time
import json
from datetime import datetime

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    print("ERROR: tkinter not found. Reinstall Python and tick tcl/tk and IDLE.")
    sys.exit(1)


# ─────────────────────────────────────────────
#  THEME / COLOUR PALETTE
# ─────────────────────────────────────────────
C = {
    "bg_dark":    "#0A0A0A",
    "bg_panel":   "#141414",
    "bg_card":    "#1C1C1C",
    "bg_input":   "#242424",
    "nissan_red": "#C3002F",
    "red_dark":   "#8B0020",
    "red_glow":   "#FF1744",
    "silver":     "#C0C0C0",
    "silver_dim": "#888888",
    "white":      "#F0F0F0",
    "text_dim":   "#666666",
    "border":     "#2A2A2A",
    "success":    "#00C853",
    "warning":    "#FFD600",
    "accent":     "#E0E0E0",
    "highlight":  "#2A1A1A",
}

SCENARIOS = [
    "FC 2+10", "FC 5+7", "FC 7+5", "FC 8+4",
    "FC 9+3", "FC 10+2", "FC 11+1", "Actuals",
    "Budget", "LY Actuals"
]

QUARTERS = ["Q1 (Apr–Jun)", "Q2 (Jul–Sep)", "Q3 (Oct–Dec)", "Q4 (Jan–Mar)"]

MONTHS_BY_Q = {
    "Q1 (Apr–Jun)": ["April", "May", "June"],
    "Q2 (Jul–Sep)": ["July", "August", "September"],
    "Q3 (Oct–Dec)": ["October", "November", "December"],
    "Q4 (Jan–Mar)": ["January", "February", "March"],
}

REQUIRED_FILES = [
    "Sales_Data.xlsx", "Volume_Plan.xlsx", "Price_Mix.xlsx",
    "Cost_Summary.xlsx", "Budget_Template.xlsx",
    "Actuals_Upload.xlsx", "Variance_Master.xlsx",
]

PAIRS_FILE = "scenario_pairs.json"   # persisted scenario pairs


# ═══════════════════════════════════════════════════════════════════════
#  NISSAN LOGO
# ═══════════════════════════════════════════════════════════════════════
def draw_nissan_logo(canvas, cx, cy, scale=1.0):
    rw, rh = int(110 * scale), int(34 * scale)
    bw, bh = int(10 * scale), int(54 * scale)
    canvas.create_oval(cx-rw, cy-rh, cx+rw, cy+rh,
                       outline=C["silver"], width=int(3*scale), fill="")
    canvas.create_oval(cx-rw+int(6*scale), cy-rh+int(5*scale),
                       cx+rw-int(6*scale), cy+rh-int(5*scale),
                       outline=C["silver_dim"], width=1, fill="")
    canvas.create_rectangle(cx-bw, cy-bh, cx+bw, cy+bh,
                            fill=C["nissan_red"], outline=C["red_dark"], width=2)
    canvas.create_text(cx, cy, text="NISSAN",
                       fill=C["white"], font=("Arial", int(9*scale), "bold"))


# ═══════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════
class NissanVarianceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Nissan  |  Variance Analysis Tool")
        self.geometry("1100x860")
        self.minsize(960, 760)
        self.configure(bg=C["bg_dark"])
        self.resizable(True, True)

        # State
        self.input_folders: list = []
        self.files_found: dict = {}
        self.all_loaded = False

        # Scenario selection (max 2)
        self.scenario_vars: dict = {}
        self._selected_count = 0

        # Stored pairs  {label: (sc1, sc2)}
        self.stored_pairs: dict = {}
        self._load_pairs()

        self._build_ui()
        self._apply_styles()
        self._refresh_pairs_list()

    # ──────────────────────────────────────────
    #  PERSIST PAIRS
    # ──────────────────────────────────────────
    def _load_pairs(self):
        if os.path.exists(PAIRS_FILE):
            try:
                with open(PAIRS_FILE) as f:
                    self.stored_pairs = json.load(f)
            except Exception:
                self.stored_pairs = {}

    def _save_pairs(self):
        try:
            with open(PAIRS_FILE, "w") as f:
                json.dump(self.stored_pairs, f, indent=2)
        except Exception:
            pass

    # ──────────────────────────────────────────
    #  STYLES
    # ──────────────────────────────────────────
    def _apply_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TFrame",       background=C["bg_dark"])
        style.configure("Card.TFrame",  background=C["bg_card"])
        style.configure("Panel.TFrame", background=C["bg_panel"])

        style.configure("TLabel",
                        background=C["bg_card"], foreground=C["white"],
                        font=("Helvetica Neue", 10))
        style.configure("Header.TLabel",
                        background=C["bg_dark"], foreground=C["white"],
                        font=("Helvetica Neue", 13, "bold"))
        style.configure("Sub.TLabel",
                        background=C["bg_card"], foreground=C["silver_dim"],
                        font=("Helvetica Neue", 9))
        style.configure("TCombobox",
                        fieldbackground=C["bg_input"], background=C["bg_input"],
                        foreground=C["white"], arrowcolor=C["nissan_red"],
                        bordercolor=C["border"])
        style.map("TCombobox",
                  fieldbackground=[("readonly", C["bg_input"])],
                  foreground=[("readonly", C["white"])])

        for name, bg, fg, hover in [
            ("Red.TButton",     C["nissan_red"], C["white"],   C["red_glow"]),
            ("Green.TButton",   C["success"],    C["bg_dark"], "#00E676"),
            ("Ghost.TButton",   C["bg_input"],   C["silver"],  C["bg_card"]),
            ("Warning.TButton", C["warning"],    C["bg_dark"], "#FFE033"),
        ]:
            style.configure(name, background=bg, foreground=fg,
                            font=("Helvetica Neue", 10, "bold"),
                            borderwidth=0, padding=(12, 8))
            style.map(name, background=[("active", hover), ("pressed", bg)])

        style.configure("TProgressbar",
                        troughcolor=C["bg_input"], background=C["nissan_red"],
                        borderwidth=0, thickness=4)

    # ──────────────────────────────────────────
    #  BUILD UI
    # ──────────────────────────────────────────
    def _build_ui(self):
        self._build_header()

        main = tk.Frame(self, bg=C["bg_dark"])
        main.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        left = tk.Frame(main, bg=C["bg_dark"])
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right = tk.Frame(main, bg=C["bg_dark"], width=300)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        self._build_input_section(left)
        self._build_period_section(left)
        self._build_scenario_section(left)   # ← updated
        self._build_output_section(left)
        self._build_file_panel(right)
        self._build_statusbar()

    # ── HEADER ──────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=C["bg_dark"], height=88)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C["nissan_red"], height=3).pack(fill="x", side="top")
        inner = tk.Frame(hdr, bg=C["bg_dark"])
        inner.pack(fill="both", expand=True, padx=24)

        logo_c = tk.Canvas(inner, width=250, height=70,
                           bg=C["bg_dark"], highlightthickness=0)
        logo_c.pack(side="left", pady=8)
        draw_nissan_logo(logo_c, 125, 38)

        tf = tk.Frame(inner, bg=C["bg_dark"])
        tf.pack(side="left", padx=28, pady=8)
        tk.Label(tf, text="VARIANCE ANALYSIS",
                 bg=C["bg_dark"], fg=C["white"],
                 font=("Georgia", 18, "bold")).pack(anchor="w")
        tk.Label(tf, text="Financial Planning & Analysis  •  Forecast Control",
                 bg=C["bg_dark"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 9)).pack(anchor="w")

        tk.Label(inner, text=datetime.now().strftime("%d %b %Y"),
                 bg=C["bg_dark"], fg=C["text_dim"],
                 font=("Courier", 9)).pack(side="right", pady=8)
        tk.Frame(hdr, bg=C["border"], height=1).pack(fill="x", side="bottom")

    # ── SECTION helper ──────────────────────────
    def _section(self, parent, title):
        outer = tk.Frame(parent, bg=C["bg_dark"])
        outer.pack(fill="x", pady=(0, 10))
        lr = tk.Frame(outer, bg=C["bg_dark"])
        lr.pack(fill="x", pady=(0, 6))
        tk.Frame(lr, bg=C["nissan_red"], width=3).pack(side="left", fill="y")
        tk.Label(lr, text=f"  {title.upper()}",
                 bg=C["bg_dark"], fg=C["silver"],
                 font=("Helvetica Neue", 9, "bold")).pack(side="left")
        card = tk.Frame(outer, bg=C["bg_card"],
                        highlightbackground=C["border"], highlightthickness=1)
        card.pack(fill="x")
        return card

    # ── INPUT FOLDERS ─────────────────────────
    def _build_input_section(self, parent):
        card = self._section(parent, "01  Input Folders")
        card.configure(padx=16, pady=14)

        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")
        self.folder_entry = tk.Entry(row, bg=C["bg_input"], fg=C["silver"],
                                     insertbackground=C["white"],
                                     relief="flat", font=("Courier", 9), bd=0)
        self.folder_entry.pack(side="left", fill="x", expand=True, ipady=7, ipadx=6)
        self.folder_entry.insert(0, "Paste folder path or browse…")
        self.folder_entry.bind("<FocusIn>", self._clear_placeholder)
        ttk.Button(row, text="Browse", style="Ghost.TButton",
                   command=self._browse_folder).pack(side="left", padx=(8, 0))
        ttk.Button(row, text="Add", style="Red.TButton",
                   command=self._add_folder).pack(side="left", padx=(6, 0))

        self.folder_listbox = tk.Listbox(card, bg=C["bg_input"], fg=C["silver"],
                                          selectbackground=C["nissan_red"],
                                          selectforeground=C["white"],
                                          relief="flat", bd=0,
                                          font=("Courier", 8), height=3,
                                          activestyle="none")
        self.folder_listbox.pack(fill="x", pady=(10, 0))

        br = tk.Frame(card, bg=C["bg_card"])
        br.pack(fill="x", pady=(8, 0))
        ttk.Button(br, text="✕  Remove Selected", style="Ghost.TButton",
                   command=self._remove_folder).pack(side="left")
        ttk.Button(br, text="⟳  Scan Files", style="Red.TButton",
                   command=self._scan_files).pack(side="right")

    # ── PERIOD ────────────────────────────────
    def _build_period_section(self, parent):
        card = self._section(parent, "02  Financial Period")
        card.configure(padx=16, pady=14)
        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")

        qf = tk.Frame(row, bg=C["bg_card"])
        qf.pack(side="left", fill="x", expand=True, padx=(0, 16))
        tk.Label(qf, text="Quarter", bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 8)).pack(anchor="w", pady=(0, 4))
        self.quarter_var = tk.StringVar()
        self.quarter_cb = ttk.Combobox(qf, textvariable=self.quarter_var,
                                        values=QUARTERS, state="readonly",
                                        font=("Helvetica Neue", 10))
        self.quarter_cb.pack(fill="x", ipady=4)
        self.quarter_cb.bind("<<ComboboxSelected>>", self._on_quarter_change)

        mf = tk.Frame(row, bg=C["bg_card"])
        mf.pack(side="left", fill="x", expand=True)
        tk.Label(mf, text="Month", bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 8)).pack(anchor="w", pady=(0, 4))
        self.month_var = tk.StringVar()
        self.month_cb = ttk.Combobox(mf, textvariable=self.month_var,
                                      values=[], state="readonly",
                                      font=("Helvetica Neue", 10))
        self.month_cb.pack(fill="x", ipady=4)

    # ══════════════════════════════════════════════════════════════
    #  SCENARIO SECTION  — max 2 selections + store pairs
    # ══════════════════════════════════════════════════════════════
    def _build_scenario_section(self, parent):
        card = self._section(parent, "03  Forecast Scenarios  (select 1 or 2  →  MTD + YTD tabs each)")
        card.configure(padx=16, pady=14)

        # ── top: checkboxes ──────────────────────
        top = tk.Frame(card, bg=C["bg_card"])
        top.pack(fill="x")

        left_col = tk.Frame(top, bg=C["bg_card"])
        left_col.pack(side="left", fill="both", expand=True)

        right_col = tk.Frame(top, bg=C["bg_card"], width=220)
        right_col.pack(side="right", fill="y", padx=(16, 0))
        right_col.pack_propagate(False)

        # counter badge
        self.sc_counter = tk.Label(left_col,
                                    text="0 selected  (1 or 2)",
                                    bg=C["bg_card"], fg=C["silver_dim"],
                                    font=("Helvetica Neue", 8, "italic"))
        self.sc_counter.pack(anchor="w", pady=(0, 8))

        grid = tk.Frame(left_col, bg=C["bg_card"])
        grid.pack(fill="x")

        cols = 2
        for i, sc in enumerate(SCENARIOS):
            var = tk.BooleanVar()
            self.scenario_vars[sc] = var
            cb = tk.Checkbutton(
                grid, text=sc, variable=var,
                bg=C["bg_card"], fg=C["white"],
                selectcolor=C["nissan_red"],
                activebackground=C["bg_card"],
                activeforeground=C["white"],
                font=("Helvetica Neue", 9),
                padx=6, pady=2,
                relief="flat", bd=0,
                command=lambda s=sc: self._on_scenario_toggle(s)
            )
            cb.grid(row=i // cols, column=i % cols, sticky="w", padx=4, pady=2)

        # clear btn
        btn_row = tk.Frame(left_col, bg=C["bg_card"])
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="Clear Selection", style="Ghost.TButton",
                   command=self._clear_scenarios).pack(side="left")

        # ── right col: store pair ──────────────────
        tk.Label(right_col, text="STORED PAIRS",
                 bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 8, "bold")).pack(anchor="w", pady=(0, 6))

        # pair name entry + store button
        pe_row = tk.Frame(right_col, bg=C["bg_card"])
        pe_row.pack(fill="x", pady=(0, 6))
        self.pair_name_entry = tk.Entry(pe_row, bg=C["bg_input"], fg=C["silver"],
                                         insertbackground=C["white"],
                                         relief="flat", font=("Courier", 8), bd=0)
        self.pair_name_entry.pack(side="left", fill="x", expand=True,
                                   ipady=5, ipadx=4)
        self.pair_name_entry.insert(0, "Pair label…")
        self.pair_name_entry.bind("<FocusIn>", self._clear_pair_placeholder)
        ttk.Button(pe_row, text="Store", style="Red.TButton",
                   command=self._store_pair).pack(side="left", padx=(6, 0))

        # pairs listbox
        self.pairs_listbox = tk.Listbox(
            right_col, bg=C["bg_input"], fg=C["silver"],
            selectbackground=C["nissan_red"], selectforeground=C["white"],
            relief="flat", bd=0, font=("Courier", 8), height=6,
            activestyle="none"
        )
        self.pairs_listbox.pack(fill="x", pady=(0, 6))
        self.pairs_listbox.bind("<<ListboxSelect>>", self._on_pair_select)

        pbr = tk.Frame(right_col, bg=C["bg_card"])
        pbr.pack(fill="x")
        ttk.Button(pbr, text="Load", style="Ghost.TButton",
                   command=self._load_pair).pack(side="left", padx=(0, 4))
        ttk.Button(pbr, text="Delete", style="Ghost.TButton",
                   command=self._delete_pair).pack(side="left")

        # active pair indicator
        self.active_pair_label = tk.Label(
            right_col, text="No pair active",
            bg=C["bg_card"], fg=C["text_dim"],
            font=("Helvetica Neue", 8, "italic"), wraplength=200, justify="left"
        )
        self.active_pair_label.pack(anchor="w", pady=(8, 0))

    # ── OUTPUT ────────────────────────────────
    def _build_output_section(self, parent):
        card = self._section(parent, "04  Output")
        card.configure(padx=16, pady=14)

        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")
        self.output_entry = tk.Entry(row, bg=C["bg_input"], fg=C["silver"],
                                      insertbackground=C["white"],
                                      relief="flat", font=("Courier", 9), bd=0)
        self.output_entry.pack(side="left", fill="x", expand=True,
                               ipady=7, ipadx=6)
        self.output_entry.insert(0, "Select output folder…")
        self.output_entry.bind("<FocusIn>", self._clear_output_placeholder)
        ttk.Button(row, text="Browse", style="Ghost.TButton",
                   command=self._browse_output).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(card, mode="determinate",
                                         style="TProgressbar", maximum=100)
        self.progress.pack(fill="x", pady=(12, 6))
        self.progress_label = tk.Label(card, text="",
                                        bg=C["bg_card"], fg=C["silver_dim"],
                                        font=("Helvetica Neue", 8))
        self.progress_label.pack(anchor="w")

        # two action buttons side-by-side
        btn_row = tk.Frame(card, bg=C["bg_card"])
        btn_row.pack(fill="x", pady=(14, 0))

        ttk.Button(btn_row, text="⚙   GENERATE OUTPUT FILES",
                   style="Red.TButton",
                   command=self._generate_files).pack(side="left", fill="x",
                                                       expand=True, ipady=4,
                                                       padx=(0, 8))

        ttk.Button(btn_row, text="▶   RUN ANALYSIS",
                   style="Green.TButton",
                   command=self._run_analysis).pack(side="left", fill="x",
                                                     expand=True, ipady=4)

    # ── FILE PANEL (right) ────────────────────
    def _build_file_panel(self, parent):
        tk.Frame(parent, bg=C["bg_dark"]).pack(fill="x")
        tk.Label(parent, text="FILE MANIFEST",
                 bg=C["bg_dark"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 8, "bold")).pack(anchor="w", padx=4, pady=(0, 6))
        panel = tk.Frame(parent, bg=C["bg_panel"],
                         highlightbackground=C["border"], highlightthickness=1)
        panel.pack(fill="both", expand=True)
        self.file_text = tk.Text(panel, bg=C["bg_panel"], fg=C["silver"],
                                  relief="flat", bd=0, font=("Courier", 8),
                                  wrap="none", state="disabled",
                                  insertbackground=C["white"])
        scroll = ttk.Scrollbar(panel, command=self.file_text.yview)
        self.file_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.file_text.pack(fill="both", expand=True, padx=8, pady=8)
        self.file_text.tag_configure("folder",  foreground=C["nissan_red"],
                                      font=("Courier", 8, "bold"))
        self.file_text.tag_configure("ok",      foreground=C["success"])
        self.file_text.tag_configure("missing", foreground=C["warning"])
        self.file_text.tag_configure("dim",     foreground=C["text_dim"])
        self.file_summary = tk.Label(parent, text="No folders loaded",
                                      bg=C["bg_dark"], fg=C["text_dim"],
                                      font=("Helvetica Neue", 8),
                                      wraplength=280, justify="left")
        self.file_summary.pack(anchor="w", padx=6, pady=(8, 0))

    # ── STATUS BAR ───────────────────────────
    def _build_statusbar(self):
        bar = tk.Frame(self, bg=C["bg_panel"], height=26)
        bar.pack(fill="x", side="bottom")
        tk.Frame(bar, bg=C["border"], height=1).pack(fill="x", side="top")
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(bar, textvariable=self.status_var,
                 bg=C["bg_panel"], fg=C["text_dim"],
                 font=("Helvetica Neue", 8)).pack(side="left", padx=12, pady=4)
        tk.Label(bar, text="v2.0  •  Nissan FP&A",
                 bg=C["bg_panel"], fg=C["text_dim"],
                 font=("Helvetica Neue", 8)).pack(side="right", padx=12, pady=4)

    # ══════════════════════════════════════════════════════════════
    #  SCENARIO LOGIC — max 2
    # ══════════════════════════════════════════════════════════════
    def _on_scenario_toggle(self, toggled_scenario):
        selected = [s for s, v in self.scenario_vars.items() if v.get()]
        count = len(selected)

        if count > 2:
            # Uncheck the one just toggled (enforce max 2)
            self.scenario_vars[toggled_scenario].set(False)
            self.status_var.set("Maximum 2 scenarios allowed. Deselect one first.")
            count = 2

        self._selected_count = count
        if count == 0:
            colour = C["silver_dim"]
            label  = "0 selected  (1 or 2)"
        elif count == 1:
            colour = C["warning"]
            label  = f"1 selected  →  2 tabs (MTD + YTD)"
        else:
            colour = C["success"]
            label  = f"2 selected  →  4 tabs (MTD + YTD each)"
        self.sc_counter.configure(text=label, fg=colour)

    def _clear_scenarios(self):
        for v in self.scenario_vars.values():
            v.set(False)
        self._selected_count = 0
        self.sc_counter.configure(text="0 selected  (1 or 2)", fg=C["silver_dim"])
        self.active_pair_label.configure(text="No pair active", fg=C["text_dim"])

    def _get_selected_scenarios(self):
        return [s for s, v in self.scenario_vars.items() if v.get()]

    # ══════════════════════════════════════════════════════════════
    #  PAIR STORE / LOAD / DELETE
    # ══════════════════════════════════════════════════════════════
    def _clear_pair_placeholder(self, _):
        if self.pair_name_entry.get() == "Pair label…":
            self.pair_name_entry.delete(0, "end")

    def _store_pair(self):
        selected = self._get_selected_scenarios()
        if len(selected) != 2:
            messagebox.showwarning("Select 2 Scenarios",
                                   "Please select exactly 2 scenarios before storing a pair.")
            return
        label = self.pair_name_entry.get().strip()
        if not label or label == "Pair label…":
            # Auto-generate label
            label = f"{selected[0]} vs {selected[1]}"

        self.stored_pairs[label] = selected
        self._save_pairs()
        self._refresh_pairs_list()
        self.pair_name_entry.delete(0, "end")
        self.pair_name_entry.insert(0, "Pair label…")
        self.active_pair_label.configure(
            text=f"Stored: {label}\n{selected[0]}  ↔  {selected[1]}",
            fg=C["success"]
        )
        self.status_var.set(f"Pair stored: {label}")

    def _refresh_pairs_list(self):
        self.pairs_listbox.delete(0, "end")
        for label in self.stored_pairs:
            sc1, sc2 = self.stored_pairs[label]
            self.pairs_listbox.insert("end", f"  {label}  ({sc1} / {sc2})")

    def _on_pair_select(self, _=None):
        pass  # selection tracked via listbox curselection

    def _load_pair(self):
        sel = self.pairs_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select a Pair", "Click a stored pair first, then Load.")
            return
        label = list(self.stored_pairs.keys())[sel[0]]
        sc1, sc2 = self.stored_pairs[label]
        # Clear all then set the two
        for v in self.scenario_vars.values():
            v.set(False)
        self.scenario_vars[sc1].set(True)
        self.scenario_vars[sc2].set(True)
        self._selected_count = 2
        self.sc_counter.configure(text="2 selected  →  4 tabs (MTD + YTD each)", fg=C["success"])
        self.active_pair_label.configure(
            text=f"Active: {label}\n{sc1}  ↔  {sc2}",
            fg=C["success"]
        )
        self.status_var.set(f"Loaded pair: {label}")

    def _delete_pair(self):
        sel = self.pairs_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select a Pair", "Click a stored pair first, then Delete.")
            return
        label = list(self.stored_pairs.keys())[sel[0]]
        if messagebox.askyesno("Delete Pair", f"Delete stored pair '{label}'?"):
            del self.stored_pairs[label]
            self._save_pairs()
            self._refresh_pairs_list()
            self.active_pair_label.configure(text="No pair active", fg=C["text_dim"])
            self.status_var.set(f"Deleted pair: {label}")

    # ══════════════════════════════════════════════════════════════
    #  RUN ANALYSIS  ← calls variance_engine.py
    # ══════════════════════════════════════════════════════════════
    def _run_analysis(self):
        # Validate inputs
        selected = self._get_selected_scenarios()
        if len(selected) == 0:
            messagebox.showwarning("No Scenario",
                                   "Please select at least 1 scenario to run the analysis.")
            return
        if not self.quarter_var.get():
            messagebox.showwarning("No Quarter", "Please select a Financial Quarter.")
            return
        if not self.month_var.get():
            messagebox.showwarning("No Month", "Please select a Month.")
            return
        if not self.all_loaded or not self.files_found:
            messagebox.showwarning("Files Not Loaded",
                                   "Please add folders and click 'Scan Files' first.")
            return
        out_dir = self.output_entry.get().strip()
        if not out_dir or out_dir == "Select output folder…":
            messagebox.showwarning("No Output Folder", "Please select an output folder.")
            return

        # Collect all scanned file paths
        all_files = []
        for folder, files in self.files_found.items():
            for f in files:
                all_files.append(os.path.join(folder, f))

        # Build arguments dict passed to the engine
        # 'scenarios' drives tab generation; legacy keys kept for compatibility
        run_args = {
            "scenarios":     selected,
            "scenario_1":    selected[0],
            "scenario_2":    selected[1] if len(selected) > 1 else "",
            "quarter":       self.quarter_var.get(),
            "month":         self.month_var.get(),
            "input_folders": self.input_folders,
            "input_files":   all_files,
            "output_folder": out_dir,
            "timestamp":     datetime.now().strftime("%Y%m%d_%H%M%S"),
        }

        sc_lines = "\n".join(f"  Scenario {i+1}  : {s}"
                             for i, s in enumerate(selected))
        confirm = (
            f"Run Variance Analysis with:\n\n"
            f"{sc_lines}\n"
            f"  Quarter    : {self.quarter_var.get()}\n"
            f"  Month      : {self.month_var.get()}\n"
            f"  Output     : {out_dir}\n\n"
            f"Each scenario will produce 2 tabs (MTD + YTD).\n\nProceed?"
        )
        if not messagebox.askyesno("Confirm Run", confirm):
            return

        self.status_var.set("Running analysis…")
        self.progress["value"] = 0
        self.progress_label.configure(text="Starting engine…")
        threading.Thread(target=self._run_thread, args=(run_args,), daemon=True).start()

    def _run_thread(self, args: dict):
        try:
            import variance_engine
            variance_engine.run_variance(args)
            self.after(0, self._on_run_complete, args["output_folder"], None)
        except ImportError:
            self.after(0, self._on_run_complete, args["output_folder"],
                       "variance_engine.py not found in the same folder.")
        except Exception as e:
            self.after(0, self._on_run_complete, args["output_folder"], str(e))

    def _on_run_complete(self, out_dir, error):
        self.progress["value"] = 100
        if error:
            self.progress_label.configure(text=f"Error: {error}")
            self.status_var.set("Run failed.")
            messagebox.showerror("Run Error", error)
        else:
            self.progress_label.configure(text="Analysis complete ✓")
            self.status_var.set(f"Analysis saved to: {out_dir}")
            messagebox.showinfo("Done", f"Analysis complete.\nOutput: {out_dir}")

    # ══════════════════════════════════════════════════════════════
    #  FOLDER / FILE LOGIC (unchanged from original)
    # ══════════════════════════════════════════════════════════════
    def _clear_placeholder(self, _):
        if self.folder_entry.get() == "Paste folder path or browse…":
            self.folder_entry.delete(0, "end")

    def _clear_output_placeholder(self, _):
        if self.output_entry.get() == "Select output folder…":
            self.output_entry.delete(0, "end")

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select Input Folder")
        if path:
            self.folder_entry.delete(0, "end")
            self.folder_entry.insert(0, path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, path)

    def _add_folder(self):
        path = self.folder_entry.get().strip()
        if not path or path == "Paste folder path or browse…":
            messagebox.showwarning("No Path", "Please enter or browse for a folder path.")
            return
        if not os.path.isdir(path):
            messagebox.showerror("Invalid Folder", f"Folder not found:\n{path}")
            return
        if path in self.input_folders:
            messagebox.showinfo("Duplicate", "This folder is already in the list.")
            return
        self.input_folders.append(path)
        self.folder_listbox.insert("end", f"  📁  {path}")
        self.status_var.set(f"Folder added: {os.path.basename(path)}")
        self.folder_entry.delete(0, "end")
        self.folder_entry.insert(0, "Paste folder path or browse…")

    def _remove_folder(self):
        sel = self.folder_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.folder_listbox.delete(idx)
        removed = self.input_folders.pop(idx)
        self.status_var.set(f"Removed: {os.path.basename(removed)}")

    def _on_quarter_change(self, _=None):
        q = self.quarter_var.get()
        months = MONTHS_BY_Q.get(q, [])
        self.month_cb.configure(values=months)
        self.month_var.set("")

    def _scan_files(self):
        if not self.input_folders:
            messagebox.showwarning("No Folders", "Please add at least one input folder first.")
            return
        self.all_loaded = False
        self.files_found.clear()
        self._update_file_panel()
        self.status_var.set("Scanning files…")
        threading.Thread(target=self._scan_thread, daemon=True).start()

    def _scan_thread(self):
        for folder in self.input_folders:
            found = []
            try:
                for f in os.listdir(folder):
                    if os.path.isfile(os.path.join(folder, f)):
                        found.append(f)
            except Exception:
                pass
            self.files_found[folder] = found
            time.sleep(0.15)
        self.after(0, self._on_scan_complete)

    def _on_scan_complete(self):
        self._update_file_panel()
        total = sum(len(v) for v in self.files_found.values())
        self.all_loaded = True
        self.status_var.set(
            f"Scan complete — {total} file(s) found across {len(self.input_folders)} folder(s)"
        )
        self._show_loaded_popup(total)

    def _update_file_panel(self):
        self.file_text.configure(state="normal")
        self.file_text.delete("1.0", "end")
        if not self.files_found:
            self.file_text.insert("end", "\n  No files scanned yet.\n", "dim")
            self.file_summary.configure(text="No folders loaded")
            self.file_text.configure(state="disabled")
            return
        total = 0
        for folder, files in self.files_found.items():
            name = os.path.basename(folder) or folder
            self.file_text.insert("end", f"\n📁  {name}\n", "folder")
            self.file_text.insert("end", f"   {folder}\n", "dim")
            if files:
                for f in sorted(files):
                    tag = "ok" if any(rf.lower() in f.lower() for rf in REQUIRED_FILES) else "dim"
                    mark = "✓" if tag == "ok" else "·"
                    self.file_text.insert("end", f"   {mark}  {f}\n", tag)
                    total += 1
            else:
                self.file_text.insert("end", "   (empty folder)\n", "missing")
            self.file_text.insert("end", "\n")
        self.file_summary.configure(
            text=f"{total} file(s) across {len(self.files_found)} folder(s)\n✓ = required  ·  · = other",
            fg=C["silver_dim"]
        )
        self.file_text.configure(state="disabled")

    def _show_loaded_popup(self, total):
        popup = tk.Toplevel(self)
        popup.title("")
        popup.configure(bg=C["bg_card"])
        popup.resizable(False, False)
        popup.geometry("380x200")
        popup.grab_set()
        popup.focus_set()
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 380) // 2
        y = self.winfo_y() + (self.winfo_height() - 200) // 2
        popup.geometry(f"+{x}+{y}")
        tk.Frame(popup, bg=C["nissan_red"], height=4).pack(fill="x")
        tk.Label(popup, text="✓", bg=C["bg_card"], fg=C["success"],
                 font=("Helvetica Neue", 36)).pack(pady=(18, 4))
        tk.Label(popup, text="All Files Loaded",
                 bg=C["bg_card"], fg=C["white"],
                 font=("Helvetica Neue", 14, "bold")).pack()
        tk.Label(popup, text=f"{total} file(s) scanned from {len(self.input_folders)} folder(s).",
                 bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 9)).pack(pady=(4, 16))
        ttk.Button(popup, text="Continue", style="Red.TButton",
                   command=popup.destroy).pack(pady=(0, 16))

    def _generate_files(self):
        if not self.all_loaded or not self.files_found:
            messagebox.showwarning("Files Not Loaded",
                                   "Please add folders and click 'Scan Files' first.")
            return
        if not self.quarter_var.get():
            messagebox.showwarning("No Quarter", "Please select a Financial Quarter.")
            return
        if not self.month_var.get():
            messagebox.showwarning("No Month", "Please select a Month.")
            return
        selected = self._get_selected_scenarios()
        if not selected:
            messagebox.showwarning("No Scenario", "Please select at least 1 forecast scenario.")
            return
        out_dir = self.output_entry.get().strip()
        if not out_dir or out_dir == "Select output folder…":
            messagebox.showwarning("No Output Folder", "Please select an output folder.")
            return
        if not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output folder:\n{e}")
                return
        confirm_msg = (
            f"Generate output files for:\n\n"
            f"  Quarter    : {self.quarter_var.get()}\n"
            f"  Month      : {self.month_var.get()}\n"
            f"  Scenarios  : {', '.join(selected)}\n\n"
            f"Files will be saved to:\n  {out_dir}\n\nProceed?"
        )
        if not messagebox.askyesno("Confirm Generation", confirm_msg):
            return
        threading.Thread(target=self._generate_thread,
                         args=(out_dir, selected), daemon=True).start()

    def _generate_thread(self, out_dir, scenarios):
        total_steps = len(scenarios) * max(
            sum(len(v) for v in self.files_found.values()), 1
        )
        step = 0
        quarter = self.quarter_var.get().split()[0]
        month = self.month_var.get()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        for scenario in scenarios:
            sc_label = scenario.replace(" ", "_").replace("+", "p")
            for folder, files in self.files_found.items():
                for fname in files:
                    src = os.path.join(folder, fname)
                    stem, ext = os.path.splitext(fname)
                    versioned = f"{stem}__{quarter}_{month}__{sc_label}__v{ts}{ext}"
                    dst = os.path.join(out_dir, versioned)
                    try:
                        shutil.copy2(src, dst)
                    except Exception:
                        pass
                    step += 1
                    pct = min(int(step / total_steps * 100), 99)
                    self.after(0, self._update_progress,
                               pct, f"Processing {scenario} — {fname}")
                    time.sleep(0.05)

        self.after(0, self._on_generation_complete, out_dir, ts)

    def _update_progress(self, pct, msg):
        self.progress["value"] = pct
        self.progress_label.configure(text=msg)
        self.status_var.set(msg)

    def _on_generation_complete(self, out_dir, ts):
        self.progress["value"] = 100
        self.progress_label.configure(text="Generation complete ✓")
        self.status_var.set(f"Output saved to: {out_dir}")
        popup = tk.Toplevel(self)
        popup.title("Generation Complete")
        popup.configure(bg=C["bg_card"])
        popup.resizable(False, False)
        popup.geometry("480x260")
        popup.grab_set()
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 480) // 2
        y = self.winfo_y() + (self.winfo_height() - 260) // 2
        popup.geometry(f"+{x}+{y}")
        tk.Frame(popup, bg=C["success"], height=4).pack(fill="x")
        tk.Label(popup, text="Files Generated Successfully",
                 bg=C["bg_card"], fg=C["white"],
                 font=("Helvetica Neue", 14, "bold")).pack(pady=(18, 4))
        tk.Label(popup,
                 text=f"Version stamp: {ts}\n\nOutput folder:\n{out_dir}",
                 bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 9), justify="center").pack(pady=(4, 16))
        br = tk.Frame(popup, bg=C["bg_card"])
        br.pack()
        ttk.Button(br, text="Open Folder", style="Ghost.TButton",
                   command=lambda: self._open_folder(out_dir)).pack(side="left", padx=6)
        ttk.Button(br, text="Close", style="Red.TButton",
                   command=popup.destroy).pack(side="left", padx=6)

    def _open_folder(self, path):
        import subprocess
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = NissanVarianceApp()
    app.mainloop()
