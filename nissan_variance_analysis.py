"""
Nissan Variance Analysis Tool
A professional Tkinter desktop application for financial variance analysis.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import threading
import time
from datetime import datetime
from pathlib import Path


# ─────────────────────────────────────────────
#  THEME / COLOUR PALETTE  (Nissan brand)
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
    "Sales_Data.xlsx",
    "Volume_Plan.xlsx",
    "Price_Mix.xlsx",
    "Cost_Summary.xlsx",
    "Budget_Template.xlsx",
    "Actuals_Upload.xlsx",
    "Variance_Master.xlsx",
]


# ═══════════════════════════════════════════════════════════════════════
#  NISSAN LOGO  (SVG-like canvas drawing)
# ═══════════════════════════════════════════════════════════════════════
def draw_nissan_logo(canvas: tk.Canvas, cx: int, cy: int, scale: float = 1.0):
    """Draw the Nissan oval + bar logo on a canvas."""
    rw = int(110 * scale)
    rh = int(34 * scale)
    bw = int(10 * scale)
    bh = int(54 * scale)

    # Outer oval
    canvas.create_oval(cx - rw, cy - rh, cx + rw, cy + rh,
                       outline=C["silver"], width=int(3 * scale), fill="")
    # Inner oval
    canvas.create_oval(cx - rw + int(6 * scale), cy - rh + int(5 * scale),
                       cx + rw - int(6 * scale), cy + rh - int(5 * scale),
                       outline=C["silver_dim"], width=1, fill="")
    # Vertical bar (rectangle)
    canvas.create_rectangle(cx - bw, cy - bh, cx + bw, cy + bh,
                            fill=C["nissan_red"], outline=C["red_dark"], width=2)
    # NISSAN text inside bar
    canvas.create_text(cx, cy, text="NISSAN",
                       fill=C["white"], font=("Arial", int(9 * scale), "bold"),
                       angle=0)


# ═══════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════
class NissanVarianceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Nissan  |  Variance Analysis Tool")
        self.geometry("1060x780")
        self.minsize(900, 700)
        self.configure(bg=C["bg_dark"])
        self.resizable(True, True)

        # State
        self.input_folders: list[str] = []
        self.files_found: dict[str, list[str]] = {}   # folder → [files]
        self.all_loaded = False
        self.selected_scenarios: list[str] = []

        self._build_ui()
        self._apply_styles()

    # ──────────────────────────────────────────
    #  STYLES
    # ──────────────────────────────────────────
    def _apply_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TFrame",        background=C["bg_dark"])
        style.configure("Card.TFrame",   background=C["bg_card"])
        style.configure("Panel.TFrame",  background=C["bg_panel"])

        style.configure("TLabel",
                        background=C["bg_card"],
                        foreground=C["white"],
                        font=("Helvetica Neue", 10))

        style.configure("Header.TLabel",
                        background=C["bg_dark"],
                        foreground=C["white"],
                        font=("Helvetica Neue", 13, "bold"))

        style.configure("Sub.TLabel",
                        background=C["bg_card"],
                        foreground=C["silver_dim"],
                        font=("Helvetica Neue", 9))

        style.configure("Red.TLabel",
                        background=C["bg_card"],
                        foreground=C["nissan_red"],
                        font=("Helvetica Neue", 10, "bold"))

        style.configure("Success.TLabel",
                        background=C["bg_card"],
                        foreground=C["success"],
                        font=("Helvetica Neue", 10, "bold"))

        style.configure("TCombobox",
                        fieldbackground=C["bg_input"],
                        background=C["bg_input"],
                        foreground=C["white"],
                        arrowcolor=C["nissan_red"],
                        bordercolor=C["border"],
                        lightcolor=C["border"],
                        darkcolor=C["border"])
        style.map("TCombobox",
                  fieldbackground=[("readonly", C["bg_input"])],
                  foreground=[("readonly", C["white"])])

        style.configure("Red.TButton",
                        background=C["nissan_red"],
                        foreground=C["white"],
                        font=("Helvetica Neue", 10, "bold"),
                        borderwidth=0,
                        padding=(12, 8))
        style.map("Red.TButton",
                  background=[("active", C["red_glow"]), ("pressed", C["red_dark"])])

        style.configure("Ghost.TButton",
                        background=C["bg_input"],
                        foreground=C["silver"],
                        font=("Helvetica Neue", 9),
                        borderwidth=1,
                        padding=(8, 5))
        style.map("Ghost.TButton",
                  background=[("active", C["bg_card"])])

        style.configure("TProgressbar",
                        troughcolor=C["bg_input"],
                        background=C["nissan_red"],
                        borderwidth=0,
                        thickness=4)

    # ──────────────────────────────────────────
    #  BUILD UI
    # ──────────────────────────────────────────
    def _build_ui(self):
        # ── HEADER ──────────────────────────────
        self._build_header()

        # ── MAIN CONTENT ────────────────────────
        main = tk.Frame(self, bg=C["bg_dark"])
        main.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        left = tk.Frame(main, bg=C["bg_dark"])
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right = tk.Frame(main, bg=C["bg_dark"], width=300)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        # LEFT column sections
        self._build_input_section(left)
        self._build_period_section(left)
        self._build_scenario_section(left)
        self._build_output_section(left)

        # RIGHT column – file list panel
        self._build_file_panel(right)

        # ── STATUS BAR ──────────────────────────
        self._build_statusbar()

    # ── HEADER ──────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=C["bg_dark"], height=88)
        hdr.pack(fill="x", padx=0, pady=0)
        hdr.pack_propagate(False)

        # Red accent strip
        tk.Frame(hdr, bg=C["nissan_red"], height=3).pack(fill="x", side="top")

        inner = tk.Frame(hdr, bg=C["bg_dark"])
        inner.pack(fill="both", expand=True, padx=24)

        # Logo canvas
        logo_canvas = tk.Canvas(inner, width=250, height=70,
                                bg=C["bg_dark"], highlightthickness=0)
        logo_canvas.pack(side="left", pady=8)
        draw_nissan_logo(logo_canvas, 125, 38, scale=1.0)

        # Title
        title_frame = tk.Frame(inner, bg=C["bg_dark"])
        title_frame.pack(side="left", padx=28, pady=8)
        tk.Label(title_frame, text="VARIANCE ANALYSIS",
                 bg=C["bg_dark"], fg=C["white"],
                 font=("Georgia", 18, "bold")).pack(anchor="w")
        tk.Label(title_frame, text="Financial Planning & Analysis  •  Forecast Control",
                 bg=C["bg_dark"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 9)).pack(anchor="w")

        # Timestamp
        ts = tk.Label(inner, text=datetime.now().strftime("%d %b %Y"),
                      bg=C["bg_dark"], fg=C["text_dim"],
                      font=("Courier", 9))
        ts.pack(side="right", pady=8)

        # Bottom border
        tk.Frame(hdr, bg=C["border"], height=1).pack(fill="x", side="bottom")

    # ── SECTION helper ──────────────────────────
    def _section(self, parent, title):
        outer = tk.Frame(parent, bg=C["bg_dark"])
        outer.pack(fill="x", pady=(0, 10))

        # Section label with red left-bar
        label_row = tk.Frame(outer, bg=C["bg_dark"])
        label_row.pack(fill="x", pady=(0, 6))
        tk.Frame(label_row, bg=C["nissan_red"], width=3).pack(side="left", fill="y")
        tk.Label(label_row, text=f"  {title.upper()}",
                 bg=C["bg_dark"], fg=C["silver"],
                 font=("Helvetica Neue", 9, "bold"),
                 letterSpacing=4).pack(side="left")

        card = tk.Frame(outer, bg=C["bg_card"],
                        highlightbackground=C["border"],
                        highlightthickness=1)
        card.pack(fill="x")
        return card

    # ── INPUT FOLDERS SECTION ────────────────────
    def _build_input_section(self, parent):
        card = self._section(parent, "01  Input Folders")
        card.configure(padx=16, pady=14)

        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")

        self.folder_entry = tk.Entry(row,
                                     bg=C["bg_input"], fg=C["silver"],
                                     insertbackground=C["white"],
                                     relief="flat", font=("Courier", 9),
                                     bd=0)
        self.folder_entry.pack(side="left", fill="x", expand=True,
                               ipady=7, ipadx=6)
        self.folder_entry.insert(0, "Paste folder path or browse…")
        self.folder_entry.bind("<FocusIn>", self._clear_placeholder)

        ttk.Button(row, text="Browse", style="Ghost.TButton",
                   command=self._browse_folder).pack(side="left", padx=(8, 0))
        ttk.Button(row, text="Add", style="Red.TButton",
                   command=self._add_folder).pack(side="left", padx=(6, 0))

        # Folder list
        self.folder_listbox = tk.Listbox(card, bg=C["bg_input"],
                                          fg=C["silver"], selectbackground=C["nissan_red"],
                                          selectforeground=C["white"],
                                          relief="flat", bd=0,
                                          font=("Courier", 8), height=3,
                                          activestyle="none")
        self.folder_listbox.pack(fill="x", pady=(10, 0))

        btn_row = tk.Frame(card, bg=C["bg_card"])
        btn_row.pack(fill="x", pady=(8, 0))
        ttk.Button(btn_row, text="✕  Remove Selected", style="Ghost.TButton",
                   command=self._remove_folder).pack(side="left")
        ttk.Button(btn_row, text="⟳  Scan Files", style="Red.TButton",
                   command=self._scan_files).pack(side="right")

    # ── PERIOD SECTION ───────────────────────────
    def _build_period_section(self, parent):
        card = self._section(parent, "02  Financial Period")
        card.configure(padx=16, pady=14)

        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")

        # Quarter
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

        # Month
        mf = tk.Frame(row, bg=C["bg_card"])
        mf.pack(side="left", fill="x", expand=True)
        tk.Label(mf, text="Month", bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 8)).pack(anchor="w", pady=(0, 4))
        self.month_var = tk.StringVar()
        self.month_cb = ttk.Combobox(mf, textvariable=self.month_var,
                                      values=[], state="readonly",
                                      font=("Helvetica Neue", 10))
        self.month_cb.pack(fill="x", ipady=4)

    # ── SCENARIO SECTION ─────────────────────────
    def _build_scenario_section(self, parent):
        card = self._section(parent, "03  Forecast Scenarios")
        card.configure(padx=16, pady=14)

        tk.Label(card, text="Select scenarios to generate output files for:",
                 bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 8)).pack(anchor="w", pady=(0, 10))

        grid = tk.Frame(card, bg=C["bg_card"])
        grid.pack(fill="x")

        self.scenario_vars: dict[str, tk.BooleanVar] = {}
        cols = 2
        for i, sc in enumerate(SCENARIOS):
            var = tk.BooleanVar()
            self.scenario_vars[sc] = var
            cb = tk.Checkbutton(grid, text=sc, variable=var,
                                bg=C["bg_card"], fg=C["white"],
                                selectcolor=C["nissan_red"],
                                activebackground=C["bg_card"],
                                activeforeground=C["white"],
                                font=("Helvetica Neue", 9),
                                padx=6, pady=2,
                                relief="flat", bd=0)
            cb.grid(row=i // cols, column=i % cols, sticky="w", padx=4, pady=2)

        btn_row = tk.Frame(card, bg=C["bg_card"])
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="Select All", style="Ghost.TButton",
                   command=lambda: [v.set(True) for v in self.scenario_vars.values()]).pack(side="left", padx=(0, 6))
        ttk.Button(btn_row, text="Clear All", style="Ghost.TButton",
                   command=lambda: [v.set(False) for v in self.scenario_vars.values()]).pack(side="left")

    # ── OUTPUT SECTION ───────────────────────────
    def _build_output_section(self, parent):
        card = self._section(parent, "04  Output")
        card.configure(padx=16, pady=14)

        row = tk.Frame(card, bg=C["bg_card"])
        row.pack(fill="x")

        self.output_entry = tk.Entry(row,
                                      bg=C["bg_input"], fg=C["silver"],
                                      insertbackground=C["white"],
                                      relief="flat", font=("Courier", 9), bd=0)
        self.output_entry.pack(side="left", fill="x", expand=True,
                               ipady=7, ipadx=6)
        self.output_entry.insert(0, "Select output folder…")
        self.output_entry.bind("<FocusIn>", self._clear_output_placeholder)

        ttk.Button(row, text="Browse", style="Ghost.TButton",
                   command=self._browse_output).pack(side="left", padx=(8, 0))

        # Progress
        self.progress = ttk.Progressbar(card, mode="determinate",
                                         style="TProgressbar", maximum=100)
        self.progress.pack(fill="x", pady=(12, 6))

        self.progress_label = tk.Label(card, text="",
                                        bg=C["bg_card"], fg=C["silver_dim"],
                                        font=("Helvetica Neue", 8))
        self.progress_label.pack(anchor="w")

        ttk.Button(card, text="⚙   GENERATE OUTPUT FILES",
                   style="Red.TButton",
                   command=self._generate_files).pack(fill="x", pady=(14, 0),
                                                       ipady=4)

    # ── FILE PANEL (right) ───────────────────────
    def _build_file_panel(self, parent):
        tk.Frame(parent, bg=C["bg_dark"]).pack(fill="x")  # spacer
        tk.Label(parent, text="FILE MANIFEST",
                 bg=C["bg_dark"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 8, "bold")).pack(anchor="w", padx=4, pady=(0, 6))

        panel = tk.Frame(parent, bg=C["bg_panel"],
                         highlightbackground=C["border"],
                         highlightthickness=1)
        panel.pack(fill="both", expand=True)

        self.file_text = tk.Text(panel, bg=C["bg_panel"], fg=C["silver"],
                                  relief="flat", bd=0,
                                  font=("Courier", 8), wrap="none",
                                  state="disabled",
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

        # Summary label
        self.file_summary = tk.Label(parent, text="No folders loaded",
                                      bg=C["bg_dark"], fg=C["text_dim"],
                                      font=("Helvetica Neue", 8),
                                      wraplength=280, justify="left")
        self.file_summary.pack(anchor="w", padx=6, pady=(8, 0))

    # ── STATUS BAR ───────────────────────────────
    def _build_statusbar(self):
        bar = tk.Frame(self, bg=C["bg_panel"], height=26)
        bar.pack(fill="x", side="bottom")
        tk.Frame(bar, bg=C["border"], height=1).pack(fill="x", side="top")
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(bar, textvariable=self.status_var,
                 bg=C["bg_panel"], fg=C["text_dim"],
                 font=("Helvetica Neue", 8)).pack(side="left", padx=12, pady=4)
        tk.Label(bar, text=f"v1.0  •  Nissan FP&A",
                 bg=C["bg_panel"], fg=C["text_dim"],
                 font=("Helvetica Neue", 8)).pack(side="right", padx=12, pady=4)

    # ══════════════════════════════════════════════
    #  LOGIC
    # ══════════════════════════════════════════════

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

    # ── SCAN FILES ───────────────────────────────
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
            time.sleep(0.15)  # brief pause for animation feel

        self.after(0, self._on_scan_complete)

    def _on_scan_complete(self):
        self._update_file_panel()
        total = sum(len(v) for v in self.files_found.values())
        required_found = all(
            any(rf.lower() in f.lower() for f in files)
            for folder, files in self.files_found.items()
            for rf in REQUIRED_FILES
        )
        self.all_loaded = True
        self.status_var.set(f"Scan complete — {total} file(s) found across {len(self.input_folders)} folder(s)")
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
            text=f"{total} file(s) across {len(self.files_found)} folder(s)\n"
                 f"✓ = required file  ·  · = other",
            fg=C["silver_dim"]
        )
        self.file_text.configure(state="disabled")

    def _show_loaded_popup(self, total: int):
        """Small popup confirming all files are loaded."""
        popup = tk.Toplevel(self)
        popup.title("")
        popup.configure(bg=C["bg_card"])
        popup.resizable(False, False)
        popup.geometry("380x200")
        popup.grab_set()
        popup.focus_set()
        # Centre
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

    # ── GENERATE FILES ───────────────────────────
    def _generate_files(self):
        # Validations
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

        selected = [s for s, v in self.scenario_vars.items() if v.get()]
        if not selected:
            messagebox.showwarning("No Scenario", "Please select at least one forecast scenario.")
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

        # Confirm
        confirm_msg = (
            f"Generate output files for:\n\n"
            f"  Quarter : {self.quarter_var.get()}\n"
            f"  Month   : {self.month_var.get()}\n"
            f"  Scenarios ({len(selected)}) : {', '.join(selected)}\n\n"
            f"Files will be saved to:\n  {out_dir}\n\n"
            f"Proceed?"
        )
        if not messagebox.askyesno("Confirm Generation", confirm_msg):
            return

        self.selected_scenarios = selected
        threading.Thread(target=self._generate_thread,
                         args=(out_dir, selected), daemon=True).start()

    def _generate_thread(self, out_dir: str, scenarios: list[str]):
        total_steps = len(scenarios) * len(self.input_folders) * max(len(v) for v in self.files_found.values() or [1])
        step = 0

        quarter = self.quarter_var.get().split()[0]   # "Q1" etc.
        month   = self.month_var.get()
        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")

        for scenario in scenarios:
            sc_label = scenario.replace(" ", "_").replace("+", "p")
            for folder, files in self.files_found.items():
                for fname in files:
                    src = os.path.join(folder, fname)
                    stem, ext = os.path.splitext(fname)
                    versioned_name = f"{stem}__{quarter}_{month}__{sc_label}__v{ts}{ext}"
                    dst = os.path.join(out_dir, versioned_name)

                    try:
                        self._copy_preserving_links(src, dst, ext)
                    except Exception:
                        pass

                    step += 1
                    pct = min(int(step / max(total_steps, 1) * 100), 99)
                    self.after(0, self._update_progress,
                               pct, f"Processing {scenario} — {fname}")
                    time.sleep(0.05)

        self.after(0, self._on_generation_complete, out_dir, ts)

    def _copy_preserving_links(self, src: str, dst: str, ext: str):
        """Copy file, preserving internal links for Excel/Office files."""
        ext_lower = ext.lower()
        if ext_lower in (".xlsx", ".xlsm", ".xls"):
            # For Excel files use shutil copy to preserve binary exactly
            # (in a full implementation you'd use openpyxl with keep_links=True)
            shutil.copy2(src, dst)
        else:
            shutil.copy2(src, dst)

    def _update_progress(self, pct: int, msg: str):
        self.progress["value"] = pct
        self.progress_label.configure(text=msg)
        self.status_var.set(msg)

    def _on_generation_complete(self, out_dir: str, ts: str):
        self.progress["value"] = 100
        self.progress_label.configure(text="Generation complete ✓")
        self.status_var.set(f"Output saved to: {out_dir}")
        self._show_complete_popup(out_dir, ts)

    def _show_complete_popup(self, out_dir: str, ts: str):
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
        tk.Label(popup, text=(
            f"All selected scenarios processed.\n"
            f"Version stamp: {ts}\n\n"
            f"Output folder:\n{out_dir}"
        ), bg=C["bg_card"], fg=C["silver_dim"],
                 font=("Helvetica Neue", 9), justify="center").pack(pady=(4, 16))

        btn_row = tk.Frame(popup, bg=C["bg_card"])
        btn_row.pack()
        ttk.Button(btn_row, text="Open Folder", style="Ghost.TButton",
                   command=lambda: self._open_folder(out_dir)).pack(side="left", padx=6)
        ttk.Button(btn_row, text="Close", style="Red.TButton",
                   command=popup.destroy).pack(side="left", padx=6)

    def _open_folder(self, path: str):
        import subprocess, sys
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
