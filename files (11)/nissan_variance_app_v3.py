"""
Nissan Variance Analysis Tool  —  v3.1
Changes vs v3.0:
  • scenario_pairs.json deleted on startup (fresh session every run)
  • Max 2 scenarios selectable at once; clicking "Add Pair" appends to
    the session queue and resets the checkboxes
  • "Load" button removed — pairs accumulate via "Add Pair" only
  • Engine receives all session pairs; MTD+YTD tabs created for each
  • Master file is scanned at Generate time: any existing sheet whose
    name contains "mtd" or "ytd" (case-insensitive) is hidden in the
    output via zip/XML surgery — no link breakage
"""

import os
import sys
import threading
import time
import json
from datetime import datetime

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)

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

QUARTERS = ["Q1 (Apr-Jun)", "Q2 (Jul-Sep)", "Q3 (Oct-Dec)", "Q4 (Jan-Mar)"]

MONTHS_BY_Q = {
    "Q1 (Apr-Jun)": ["April", "May", "June"],
    "Q2 (Jul-Sep)": ["July", "August", "September"],
    "Q3 (Oct-Dec)": ["October", "November", "December"],
    "Q4 (Jan-Mar)": ["January", "February", "March"],
}

REQUIRED_FILES = [
    "Sales_Data.xlsx", "Volume_Plan.xlsx", "Price_Mix.xlsx",
    "Cost_Summary.xlsx", "Budget_Template.xlsx",
    "Actuals_Upload.xlsx", "Variance_Master.xlsx",
]

PAIRS_FILE = "scenario_pairs.json"


# ═══════════════════════════════════════════════════════════════════════
#  NISSAN LOGO
# ═══════════════════════════════════════════════════════════════════════
def draw_nissan_logo(canvas, cx, cy, scale=1.0):
    rw, rh = int(110 * scale), int(34 * scale)
    bw, bh = int(10 * scale), int(54 * scale)
    canvas.create_oval(cx - rw, cy - rh, cx + rw, cy + rh,
                       outline=C["silver"], width=int(3 * scale), fill="")
    canvas.create_oval(cx - rw + int(6 * scale), cy - rh + int(5 * scale),
                       cx + rw - int(6 * scale), cy + rh - int(5 * scale),
                       outline=C["silver_dim"], width=1, fill="")
    canvas.create_rectangle(cx - bw, cy - bh, cx + bw, cy + bh,
                            fill=C["nissan_red"], outline=C["red_dark"], width=2)
    canvas.create_text(cx, cy, text="NISSAN",
                       fill=C["white"], font=("Arial", int(9 * scale), "bold"))


# ═══════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════
class NissanVarianceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Nissan  |  Variance Analysis Tool")
        self.geometry("1160x900")
        self.minsize(980, 780)
        self.configure(bg=C["bg_dark"])
        self.resizable(True, True)

        # ── Delete stale pairs file at startup (fresh session) ──
        if os.path.exists(PAIRS_FILE):
            try:
                os.remove(PAIRS_FILE)
            except Exception:
                pass

        # State
        self.input_folders: list = []
        self.files_found:   dict = {}
        self.all_loaded            = False

        # Scenario checkboxes — exactly 2 allowed per "Add Pair" action
        self.scenario_vars: dict = {}

        # File paths
        self.master_file_var   = tk.StringVar(value="")
        self.template_file_var = tk.StringVar(value="")

        # Session pairs accumulated this run  {label: [sc1, sc2]}
        self.stored_pairs: dict = {}

        self._build_ui()
        self._apply_styles()
        self._refresh_pairs_list()

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
            ("Red.TButton",   C["nissan_red"], C["white"],   C["red_glow"]),
            ("Ghost.TButton", C["bg_input"],   C["silver"],  C["bg_card"]),
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
        main.pack(fill="both", expand=True, padx=20, pady=(0, 0))

        right = tk.Frame(main, bg=C["bg_dark"], width=300)
        right.pack(side="right", fill="y", pady=(0, 16))
        right.pack_propagate(False)

        left_outer = tk.Frame(main, bg=C["bg_dark"])
        left_outer.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self._build_action_bar(left_outer)

        self._scroll_canvas = tk.Canvas(left_outer, bg=C["bg_dark"],
                                        highlightthickness=0)
        vbar = ttk.Scrollbar(left_outer, orient="vertical",
                             command=self._scroll_canvas.yview)
        self._scroll_canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        self._scroll_canvas.pack(side="top", fill="both", expand=True)

        left = tk.Frame(self._scroll_canvas, bg=C["bg_dark"])
        self._canvas_window = self._scroll_canvas.create_window(
            (0, 0), window=left, anchor="nw")

        def _on_frame_configure(e):
            self._scroll_canvas.configure(
                scrollregion=self._scroll_canvas.bbox("all"))

        def _on_canvas_configure(e):
            self._scroll_canvas.itemconfig(
                self._canvas_window, width=e.width)

        left.bind("<Configure>", _on_frame_configure)
        self._scroll_canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(e):
            self._scroll_canvas.yview_scroll(
                int(-1 * (e.delta / 120)), "units")
        self._scroll_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self._build_input_section(left)
        self._build_period_section(left)
        self._build_scenario_section(left)
        self._build_master_file_section(left)
        self._build_template_file_section(left)
        self._build_output_folder_section(left)

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
        tk.Label(tf, text="Financial Planning & Analysis  -  Forecast Control",
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

    # ── 01  INPUT FOLDERS ───────────────────────
    def _build_input_section(self, parent):
        card = self._section(parent, "01  Input Folders")
        card.configure(padx=16, pady=14)

        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")
        self.folder_entry = tk.Entry(row, bg=C["bg_input"], fg=C["silver"],
                                     insertbackground=C["white"],
                                     relief="flat", font=("Courier", 9), bd=0)
        self.folder_entry.pack(side="left", fill="x", expand=True, ipady=7, ipadx=6)
        self.folder_entry.insert(0, "Paste folder path or browse...")
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
        ttk.Button(br, text="x  Remove Selected", style="Ghost.TButton",
                   command=self._remove_folder).pack(side="left")
        ttk.Button(br, text="Scan Files", style="Red.TButton",
                   command=self._scan_files).pack(side="right")

    # ── 02  FINANCIAL PERIOD ────────────────────
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

    # ── 03  SCENARIO SECTION ─────────────────────────────────────
    def _build_scenario_section(self, parent):
        card = self._section(
            parent,
            "03  Forecast Scenarios  (select exactly 2 → Add Pair → repeat)")
        card.configure(padx=16, pady=14)

        top = tk.Frame(card, bg=C["bg_card"])
        top.pack(fill="x")

        left_col = tk.Frame(top, bg=C["bg_card"])
        left_col.pack(side="left", fill="both", expand=True)

        right_col = tk.Frame(top, bg=C["bg_card"], width=230)
        right_col.pack(side="right", fill="y", padx=(16, 0))
        right_col.pack_propagate(False)

        # Status label — shows selection state
        self.sc_counter = tk.Label(
            left_col,
            text="Select exactly 2 scenarios, then click  Add Pair",
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

        btn_row = tk.Frame(left_col, bg=C["bg_card"])
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="Clear Selection", style="Ghost.TButton",
                   command=self._clear_scenarios).pack(side="left")
        ttk.Button(btn_row, text="Add Pair  →", style="Red.TButton",
                   command=self._add_pair).pack(side="left", padx=(10, 0))

        # Right col — session queue display
        tk.Label(right_col, text="SESSION QUEUE",
                 bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 8, "bold")).pack(anchor="w", pady=(0, 6))

        self.pairs_listbox = tk.Listbox(
            right_col, bg=C["bg_input"], fg=C["silver"],
            selectbackground=C["nissan_red"], selectforeground=C["white"],
            relief="flat", bd=0, font=("Courier", 8), height=8,
            activestyle="none")
        self.pairs_listbox.pack(fill="x", pady=(0, 6))

        pbr = tk.Frame(right_col, bg=C["bg_card"])
        pbr.pack(fill="x")
        ttk.Button(pbr, text="Remove Selected", style="Ghost.TButton",
                   command=self._delete_pair).pack(side="left")

        self.queue_summary = tk.Label(
            right_col, text="No pairs added yet",
            bg=C["bg_card"], fg=C["text_dim"],
            font=("Helvetica Neue", 8, "italic"), wraplength=210, justify="left")
        self.queue_summary.pack(anchor="w", pady=(8, 0))

    # ── 04  MASTER INPUT FILE ────────────────────────────────────
    def _build_master_file_section(self, parent):
        card = self._section(
            parent,
            "04  Master File  (base output — existing MTD/YTD tabs will be hidden)")
        card.configure(padx=16, pady=14)

        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")
        self.master_entry = tk.Entry(row, textvariable=self.master_file_var,
                                      bg=C["bg_input"], fg=C["silver"],
                                      insertbackground=C["white"],
                                      relief="flat", font=("Courier", 9), bd=0)
        self.master_entry.pack(side="left", fill="x", expand=True, ipady=7, ipadx=6)
        ttk.Button(row, text="Browse", style="Ghost.TButton",
                   command=self._browse_master_file).pack(side="left", padx=(8, 0))

        self.master_status = tk.Label(card, text="No master file selected",
                                       bg=C["bg_card"], fg=C["text_dim"],
                                       font=("Helvetica Neue", 8, "italic"))
        self.master_status.pack(anchor="w", pady=(6, 0))

    # ── 05  TEMPLATE FILE ────────────────────────────────────────
    def _build_template_file_section(self, parent):
        card = self._section(parent, "05  Template File  (contains MTD and YTD template sheets)")
        card.configure(padx=16, pady=14)

        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")
        self.template_entry = tk.Entry(row, textvariable=self.template_file_var,
                                        bg=C["bg_input"], fg=C["silver"],
                                        insertbackground=C["white"],
                                        relief="flat", font=("Courier", 9), bd=0)
        self.template_entry.pack(side="left", fill="x", expand=True, ipady=7, ipadx=6)
        ttk.Button(row, text="Browse", style="Ghost.TButton",
                   command=self._browse_template_file).pack(side="left", padx=(8, 0))

        self.template_status = tk.Label(card, text="No template file selected",
                                         bg=C["bg_card"], fg=C["text_dim"],
                                         font=("Helvetica Neue", 8, "italic"))
        self.template_status.pack(anchor="w", pady=(6, 0))

    # ── 06  OUTPUT FOLDER + PROGRESS ────────────────────────────
    def _build_output_folder_section(self, parent):
        card = self._section(parent, "06  Output Folder")
        card.configure(padx=16, pady=14)

        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")
        self.output_entry = tk.Entry(row, bg=C["bg_input"], fg=C["silver"],
                                      insertbackground=C["white"],
                                      relief="flat", font=("Courier", 9), bd=0)
        self.output_entry.pack(side="left", fill="x", expand=True,
                               ipady=7, ipadx=6)
        self.output_entry.insert(0, "Select output folder...")
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

    # ── ACTION BAR (pinned) ──────────────────────────────────────
    def _build_action_bar(self, parent):
        bar = tk.Frame(parent, bg=C["bg_dark"])
        bar.pack(side="bottom", fill="x", pady=(8, 16))
        tk.Frame(bar, bg=C["border"], height=1).pack(fill="x", pady=(0, 10))
        ttk.Button(bar, text="  GENERATE FILE",
                   style="Red.TButton",
                   command=self._generate_file).pack(fill="x", ipady=8)

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
        tk.Label(bar, text="v3.1  -  Nissan FP&A",
                 bg=C["bg_panel"], fg=C["text_dim"],
                 font=("Helvetica Neue", 8)).pack(side="right", padx=12, pady=4)

    # ══════════════════════════════════════════════════════════════
    #  SCENARIO LOGIC — max 2 at a time
    # ══════════════════════════════════════════════════════════════
    def _on_scenario_toggle(self, changed: str):
        selected = self._get_selected_scenarios()
        count = len(selected)

        # Enforce max-2 rule: uncheck the changed box if we'd exceed 2
        if count > 2:
            self.scenario_vars[changed].set(False)
            selected = self._get_selected_scenarios()
            count = 2

        if count == 0:
            self.sc_counter.configure(
                text="Select exactly 2 scenarios, then click  Add Pair",
                fg=C["silver_dim"])
        elif count == 1:
            self.sc_counter.configure(
                text=f"1 selected — choose one more to complete the pair",
                fg=C["warning"])
        else:  # count == 2
            self.sc_counter.configure(
                text=f"Pair ready:  {selected[0]}  vs  {selected[1]}  — click Add Pair",
                fg=C["success"])

    def _get_selected_scenarios(self) -> list:
        return [s for s in SCENARIOS if self.scenario_vars[s].get()]

    def _clear_scenarios(self):
        for v in self.scenario_vars.values():
            v.set(False)
        self.sc_counter.configure(
            text="Select exactly 2 scenarios, then click  Add Pair",
            fg=C["silver_dim"])

    # ── Add current 2-selection to session queue ─────────────────
    def _add_pair(self):
        selected = self._get_selected_scenarios()
        if len(selected) != 2:
            messagebox.showwarning(
                "Select Exactly 2",
                f"Please select exactly 2 scenarios before adding.\n"
                f"Currently selected: {len(selected)}")
            return

        sc1, sc2 = selected[0], selected[1]
        label = f"{sc1} vs {sc2}"

        if label in self.stored_pairs:
            messagebox.showinfo("Duplicate", f"Pair already in session:\n{label}")
            return

        self.stored_pairs[label] = [sc1, sc2]
        self._refresh_pairs_list()
        self._clear_scenarios()
        self.status_var.set(f"Added: {label}")

    def _delete_pair(self):
        sel = self.pairs_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select a Pair", "Click a pair in the queue first.")
            return
        label = list(self.stored_pairs.keys())[sel[0]]
        if messagebox.askyesno("Remove Pair", f"Remove  '{label}'  from session queue?"):
            del self.stored_pairs[label]
            self._refresh_pairs_list()
            self.status_var.set(f"Removed: {label}")

    def _refresh_pairs_list(self):
        self.pairs_listbox.delete(0, "end")
        for i, (label, scenarios) in enumerate(self.stored_pairs.items(), 1):
            self.pairs_listbox.insert("end", f"  {i}.  {label}")
        n = len(self.stored_pairs)
        if n == 0:
            self.queue_summary.configure(text="No pairs added yet", fg=C["text_dim"])
        else:
            tabs = n * 2
            self.queue_summary.configure(
                text=f"{n} pair(s) → {tabs} tabs (MTD+YTD each)",
                fg=C["success"])

    # ══════════════════════════════════════════════════════════════
    #  GENERATE FILE
    # ══════════════════════════════════════════════════════════════
    def _generate_file(self):
        if not self.quarter_var.get():
            messagebox.showwarning("No Quarter", "Please select a Financial Quarter.")
            return
        if not self.month_var.get():
            messagebox.showwarning("No Month", "Please select a Month.")
            return

        if not self.stored_pairs:
            messagebox.showwarning(
                "No Pairs",
                "No scenario pairs in the session queue.\n"
                "Select 2 scenarios and click  'Add Pair'  to build the queue.")
            return

        master = self.master_file_var.get().strip()
        if not master or not os.path.isfile(master):
            messagebox.showwarning("No Master File",
                                   "Please select a valid Master File in section 04.")
            return

        template = self.template_file_var.get().strip()
        if not template or not os.path.isfile(template):
            messagebox.showwarning("No Template File",
                                   "Please select a valid Template File in section 05.")
            return

        out_dir = self.output_entry.get().strip()
        if not out_dir or out_dir == "Select output folder...":
            messagebox.showwarning("No Output Folder", "Please select an output folder.")
            return

        pairs = [(v[0], v[1]) for v in self.stored_pairs.values()]
        all_files = [os.path.join(folder, f)
                     for folder, files in self.files_found.items()
                     for f in files]

        pair_lines = "\n".join(f"    {a} vs {b}" for a, b in pairs)
        confirm_msg = (
            f"Generate output file:\n\n"
            f"  Pairs:\n{pair_lines}\n\n"
            f"  Quarter    : {self.quarter_var.get()}\n"
            f"  Month      : {self.month_var.get()}\n"
            f"  Master     : {os.path.basename(master)}\n"
            f"  Template   : {os.path.basename(template)}\n"
            f"  Output     : {out_dir}\n\n"
            f"Existing MTD/YTD tabs in the master will be hidden.\n"
            f"Each pair produces MTD + YTD tabs.\n\nProceed?"
        )
        if not messagebox.askyesno("Confirm", confirm_msg):
            return

        os.makedirs(out_dir, exist_ok=True)

        run_args = {
            "scenario_pairs":  pairs,
            "quarter":         self.quarter_var.get(),
            "month":           self.month_var.get(),
            "input_folders":   list(self.input_folders),
            "input_files":     all_files,
            "master_file":     master,
            "template_file":   template,
            "output_folder":   out_dir,
            "timestamp":       datetime.now().strftime("%Y%m%d_%H%M%S"),
        }

        self.status_var.set("Generating output file…")
        self.progress["value"] = 0
        self.progress_label.configure(text="Starting…")
        threading.Thread(target=self._generate_thread,
                         args=(run_args,), daemon=True).start()

    def _generate_thread(self, args: dict):
        try:
            if _APP_DIR not in sys.path:
                sys.path.insert(0, _APP_DIR)

            engine_path = os.path.join(_APP_DIR, "variance_engine.py")
            if not os.path.isfile(engine_path):
                raise FileNotFoundError(
                    f"variance_engine.py not found in:\n{_APP_DIR}\n\n"
                    "Both files must be in the same folder.")

            import importlib.util
            spec   = importlib.util.spec_from_file_location(
                        "variance_engine", engine_path)
            engine = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(engine)

            self.after(0, self._update_progress, 30, "Copying master file…")
            out_path = engine.run_variance(args)
            self.after(0, self._update_progress, 95, "Saving workbook…")
            time.sleep(0.3)
            self.after(0, self._on_generate_complete,
                       args["output_folder"], out_path, None)

        except Exception as e:
            self.after(0, self._on_generate_complete,
                       args["output_folder"], None, str(e))

    def _update_progress(self, pct, msg):
        self.progress["value"] = pct
        self.progress_label.configure(text=msg)
        self.status_var.set(msg)

    def _on_generate_complete(self, out_dir, out_path, error):
        self.progress["value"] = 100 if not error else 0
        if error:
            self.progress_label.configure(text="Error")
            self.status_var.set("Generation failed.")
            messagebox.showerror("Error", error)
            return

        fname = os.path.basename(out_path) if out_path else ""
        self.progress_label.configure(text=f"Done  —  {fname}")
        self.status_var.set(f"Saved: {out_path}")

        popup = tk.Toplevel(self)
        popup.title("File Generated")
        popup.configure(bg=C["bg_card"])
        popup.resizable(False, False)
        popup.geometry("480x220")
        popup.grab_set()
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width()  - 480) // 2
        y = self.winfo_y() + (self.winfo_height() - 220) // 2
        popup.geometry(f"+{x}+{y}")

        tk.Frame(popup, bg=C["success"], height=4).pack(fill="x")
        tk.Label(popup, text="Output File Generated",
                 bg=C["bg_card"], fg=C["white"],
                 font=("Helvetica Neue", 14, "bold")).pack(pady=(18, 4))
        tk.Label(popup,
                 text=f"{fname}\n\n{out_dir}",
                 bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 9), justify="center").pack(pady=(4, 16))
        br = tk.Frame(popup, bg=C["bg_card"])
        br.pack()
        ttk.Button(br, text="Open Folder", style="Ghost.TButton",
                   command=lambda: self._open_folder(out_dir)).pack(side="left", padx=6)
        ttk.Button(br, text="Close", style="Red.TButton",
                   command=popup.destroy).pack(side="left", padx=6)

    # ══════════════════════════════════════════════════════════════
    #  FOLDER / FILE LOGIC
    # ══════════════════════════════════════════════════════════════
    def _clear_placeholder(self, _):
        if self.folder_entry.get() == "Paste folder path or browse...":
            self.folder_entry.delete(0, "end")

    def _clear_output_placeholder(self, _):
        if self.output_entry.get() == "Select output folder...":
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

    def _browse_master_file(self):
        path = filedialog.askopenfilename(
            title="Select Master File",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if path:
            self.master_file_var.set(path)
            self.master_status.configure(
                text=f"Selected: {os.path.basename(path)}", fg=C["success"])
            self.status_var.set(f"Master file set: {os.path.basename(path)}")

    def _browse_template_file(self):
        path = filedialog.askopenfilename(
            title="Select Template File",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if path:
            self.template_file_var.set(path)
            self.template_status.configure(
                text=f"Selected: {os.path.basename(path)}", fg=C["success"])
            self.status_var.set(f"Template file set: {os.path.basename(path)}")

    def _add_folder(self):
        path = self.folder_entry.get().strip()
        if not path or path == "Paste folder path or browse...":
            messagebox.showwarning("No Path", "Please enter or browse for a folder path.")
            return
        if not os.path.isdir(path):
            messagebox.showerror("Invalid Folder", f"Folder not found:\n{path}")
            return
        if path in self.input_folders:
            messagebox.showinfo("Duplicate", "This folder is already in the list.")
            return
        self.input_folders.append(path)
        self.folder_listbox.insert("end", f"  {path}")
        self.status_var.set(f"Folder added: {os.path.basename(path)}")
        self.folder_entry.delete(0, "end")
        self.folder_entry.insert(0, "Paste folder path or browse...")

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
            f"Scan complete — {total} file(s) across {len(self.input_folders)} folder(s)")
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
            self.file_text.insert("end", f"\n  {name}\n", "folder")
            self.file_text.insert("end", f"   {folder}\n", "dim")
            if files:
                for f in sorted(files):
                    tag  = "ok" if any(rf.lower() in f.lower()
                                       for rf in REQUIRED_FILES) else "dim"
                    mark = "v" if tag == "ok" else "-"
                    self.file_text.insert("end", f"   {mark}  {f}\n", tag)
                    total += 1
            else:
                self.file_text.insert("end", "   (empty folder)\n", "missing")
            self.file_text.insert("end", "\n")
        self.file_summary.configure(
            text=f"{total} file(s) across {len(self.files_found)} folder(s)\n"
                 "v = required  |  - = other",
            fg=C["silver_dim"])
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
        x = self.winfo_x() + (self.winfo_width()  - 380) // 2
        y = self.winfo_y() + (self.winfo_height() - 200) // 2
        popup.geometry(f"+{x}+{y}")
        tk.Frame(popup, bg=C["nissan_red"], height=4).pack(fill="x")
        tk.Label(popup, text="All Files Loaded",
                 bg=C["bg_card"], fg=C["white"],
                 font=("Helvetica Neue", 14, "bold")).pack(pady=(24, 4))
        tk.Label(popup,
                 text=f"{total} file(s) from {len(self.input_folders)} folder(s).",
                 bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 9)).pack(pady=(4, 16))
        ttk.Button(popup, text="Continue", style="Red.TButton",
                   command=popup.destroy).pack(pady=(0, 16))

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
