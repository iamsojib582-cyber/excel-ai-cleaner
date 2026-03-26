# =========================
# FILE: ui.py
# FULL PRODUCT-LEVEL UI (DETAILED)
# =========================

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

from data_loader import load_file
from cleaner import Cleaner
from ai_engine import get_engine
from utils import THEME
BG_DARK = THEME["bg"]
BG_PANEL = THEME["panel"]
GOLD = THEME["gold"]
GOLD_BRIGHT = THEME["gold_bright"]
TEXT_PRIMARY = THEME["text"]
TEXT_DIM = THEME["dim"]
BORDER = THEME["border"]


class ExcelAICleanerApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel AI Cleaner — Professional Edition")
        self.root.geometry("1400x820")
        self.root.configure(bg=BG_DARK)

        self.df = pd.DataFrame()
        self.original_df = pd.DataFrame()

        self._build_ui()

    # =========================================================
    # MAIN UI STRUCTURE
    # =========================================================
    def _build_ui(self):

        # ───────── TOP BAR ─────────
        top_bar = tk.Frame(self.root, bg=BG_PANEL, height=60)
        top_bar.pack(fill="x")

        tk.Label(
            top_bar,
            text="Excel AI Cleaner",
            bg=BG_PANEL,
            fg=GOLD_BRIGHT,
            font=("Segoe UI", 18, "bold")
        ).pack(side="left", padx=20)

        self.status_label = tk.Label(
            top_bar,
            text="Ready",
            bg=BG_PANEL,
            fg=TEXT_DIM,
            font=("Segoe UI", 10)
        )
        self.status_label.pack(side="right", padx=20)

        # ───────── TOOLBAR ─────────
        toolbar = tk.Frame(self.root, bg=BG_DARK)
        toolbar.pack(fill="x", pady=5)

        self._btn(toolbar, "📂 Open File", self.open_file)
        self._btn(toolbar, "🤖 AI Scan", self.ai_scan)
        self._btn(toolbar, "✨ Auto Clean", self.auto_clean)
        self._btn(toolbar, "↩ Reset", self.reset_data)
        self._btn(toolbar, "💾 Export CSV", self.export_csv)

        # ───────── NOTEBOOK ─────────
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.preview_tab = tk.Frame(self.notebook, bg=BG_DARK)
        self.ai_tab = tk.Frame(self.notebook, bg=BG_DARK)
        self.clean_tab = tk.Frame(self.notebook, bg=BG_DARK)
        self.chart_tab = tk.Frame(self.notebook, bg=BG_DARK)

        self.notebook.add(self.preview_tab, text="📊 Data Preview")
        self.notebook.add(self.ai_tab, text="🤖 AI Analysis")
        self.notebook.add(self.clean_tab, text="🧹 Cleaning")
        self.notebook.add(self.chart_tab, text="📈 Visualization")

        self._build_preview_tab()
        self._build_ai_tab()
        self._build_clean_tab()
        self._build_chart_tab()

    # =========================================================
    # BUTTON STYLE
    # =========================================================
    def _btn(self, parent, text, command):
        tk.Button(
            parent,
            text=text,
            command=command,
            bg=BG_PANEL,
            fg=GOLD,
            activebackground=GOLD,
            activeforeground=BG_DARK,
            font=("Segoe UI", 10, "bold"),
            bd=0,
            padx=15,
            pady=8
        ).pack(side="left", padx=5)

    # =========================================================
    # PREVIEW TAB
    # =========================================================
    def _build_preview_tab(self):

        top = tk.Frame(self.preview_tab, bg=BG_DARK)
        top.pack(fill="x", padx=10, pady=5)

        self.summary_label = tk.Label(
            top,
            text="Rows: 0 | Columns: 0 | Missing: 0 | Duplicates: 0",
            bg=BG_DARK,
            fg=TEXT_PRIMARY,
            font=("Segoe UI", 10, "bold")
        )
        self.summary_label.pack(anchor="w")

        frame = tk.Frame(self.preview_tab, bg=BG_DARK)
        frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(frame, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        scroll_y.pack(side="right", fill="y")

        self.tree.configure(yscrollcommand=scroll_y.set)

    # =========================================================
    # AI TAB
    # =========================================================
    def _build_ai_tab(self):

        self.ai_text = tk.Text(
            self.ai_tab,
            bg="#0d1117",
            fg="#c9d1d9",
            font=("Consolas", 10),
            insertbackground="white"
        )
        self.ai_text.pack(fill="both", expand=True, padx=10, pady=10)

    # =========================================================
    # CLEAN TAB
    # =========================================================
    def _build_clean_tab(self):

        frame = tk.Frame(self.clean_tab, bg=BG_DARK)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.clean_log = tk.Text(
            frame,
            bg=BG_DARK,
            fg=TEXT_PRIMARY,
            font=("Consolas", 10)
        )
        self.clean_log.pack(fill="both", expand=True)

    # =========================================================
    # CHART TAB
    # =========================================================
    def _build_chart_tab(self):

        frame = tk.Frame(self.chart_tab, bg=BG_DARK)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(
            frame,
            text="(Chart feature placeholder)",
            bg=BG_DARK,
            fg=TEXT_DIM
        ).pack()

    # =========================================================
    # CORE FUNCTIONS
    # =========================================================
    def open_file(self):
        path = filedialog.askopenfilename()
        if not path:
            return

        try:
            self.df = load_file(path)
            self.original_df = self.df.copy()
            self.update_preview()
            self.status_label.config(text="File Loaded")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def update_preview(self):
        if self.df.empty:
            return

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(self.df.columns)

        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)

        for _, row in self.df.head(200).iterrows():
            self.tree.insert("", "end", values=list(row))

        missing = int(self.df.isna().sum().sum())
        duplicates = int(self.df.duplicated().sum())

        self.summary_label.config(
            text=f"Rows: {len(self.df)} | Columns: {len(self.df.columns)} | Missing: {missing} | Duplicates: {duplicates}"
        )

    def ai_scan(self):
        if self.df.empty:
            return

        self.ai_text.delete("1.0", "end")
        issues = run_ai_scan(self.df)

        for i in issues:
            self.ai_text.insert("end", i + "\n")

        self.status_label.config(text="AI Scan Complete")


        self.update_preview()

        for l in log:
            self.clean_log.insert("end", l + "\n")

        self.status_label.config(text="Data Cleaned")

    def reset_data(self):
        if self.original_df.empty:
            return

        self.df = self.original_df.copy()
        self.update_preview()
        self.clean_log.delete("1.0", "end")
        self.ai_text.delete("1.0", "end")

        self.status_label.config(text="Reset Done")

    def export_csv(self):
        if self.df.empty:
            return

        path = filedialog.asksaveasfilename(defaultextension=".csv")
        if not path:
            return

        self.df.to_csv(path, index=False)
        messagebox.showinfo("Saved", "File exported successfully")