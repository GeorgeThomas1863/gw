"""
ui/app.py — GrayWolfe main application window.
"""
from __future__ import annotations

import queue
import sqlite3
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import config
from core.import_data import detect_column_types, parse_import_input, run_default_import, run_unrelated_import
from core.search import parse_raw_input, run_search
from data.database import get_current_user
from data.sync import pull_from_master
from utils.errors import GWError
from utils.logger import get_logger

logger = get_logger(__name__)

_DELIMITER_OPTIONS = {
    "Auto": None,
    "!! (double bang)": "!!",
    "Tab": "\t",
    "Newline": "\n",
    "Comma": ",",
}

_SEARCH_MODE_OPTIONS = ("GW + S", "GW Only", "S Only")
_IMPORT_MODE_OPTIONS = ("Default Import", "Unrelated Import")
_TYPE_OPTIONS = ["Auto-Detect"] + list(config.SELECTOR_TYPES)


class GrayWolfeApp(tk.Tk):
    """Root application window with Search and Add Data tabs."""

    def __init__(self, conn: sqlite3.Connection) -> None:
        super().__init__()
        self.conn = conn
        self.username = get_current_user()

        self.title(f"{config.APP_NAME} v{config.APP_VERSION}")
        self.minsize(640, 480)
        self.resizable(True, True)

        self._result_queue: queue.Queue = queue.Queue()

        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)
        self._build_search_tab()
        self._build_add_tab()
        self._build_status_bar()

    # ---- Search Tab ----

    def _build_search_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Search")
        frame.columnconfigure(1, weight=1)

        # Input
        ttk.Label(frame, text="Search Input:").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 2))
        self._search_text = tk.Text(frame, height=8, wrap="none")
        self._search_text.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 6))
        frame.rowconfigure(1, weight=1)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self._search_text.yview)
        vsb.grid(row=1, column=2, sticky="ns", pady=(0, 6))
        self._search_text["yscrollcommand"] = vsb.set

        # Options row
        opts = ttk.Frame(frame)
        opts.grid(row=2, column=0, columnspan=2, sticky="ew")

        ttk.Label(opts, text="Delimiter:").pack(side="left")
        self._search_delim_var = tk.StringVar(value="Auto")
        ttk.Combobox(opts, textvariable=self._search_delim_var,
                     values=list(_DELIMITER_OPTIONS), state="readonly",
                     width=16).pack(side="left", padx=(4, 12))

        ttk.Label(opts, text="Search:").pack(side="left")
        self._search_mode_var = tk.StringVar(value="GW + S")
        search_mode_cb = ttk.Combobox(opts, textvariable=self._search_mode_var,
                                       values=list(_SEARCH_MODE_OPTIONS),
                                       state="readonly", width=12)
        search_mode_cb.pack(side="left", padx=(4, 12))
        search_mode_cb.bind("<<ComboboxSelected>>", self._on_search_mode_change)

        # Token row
        token_row = ttk.Frame(frame)
        token_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        self._token_label = ttk.Label(token_row, text="S Token:")
        self._token_label.pack(side="left")
        self._token_var = tk.StringVar()
        self._token_entry = ttk.Entry(token_row, textvariable=self._token_var,
                                       show="*", width=50)
        self._token_entry.pack(side="left", padx=(4, 8), fill="x", expand=True)

        # Search button
        btn_row = ttk.Frame(frame)
        btn_row.grid(row=4, column=0, columnspan=2, sticky="e", pady=(8, 0))
        self._btn_search = ttk.Button(btn_row, text="Search",
                                      command=self._do_search, width=12)
        self._btn_search.pack(side="right")

        self._on_search_mode_change()

    # ---- Add Data Tab ----

    def _build_add_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(frame, text="Add Data")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Import Input:").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 2))
        self._add_text = tk.Text(frame, height=8, wrap="none")
        self._add_text.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 6))
        frame.rowconfigure(1, weight=1)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self._add_text.yview)
        vsb.grid(row=1, column=2, sticky="ns", pady=(0, 6))
        self._add_text["yscrollcommand"] = vsb.set

        # Options row
        opts = ttk.Frame(frame)
        opts.grid(row=2, column=0, columnspan=2, sticky="ew")

        ttk.Label(opts, text="Delimiter:").pack(side="left")
        self._add_delim_var = tk.StringVar(value="Auto")
        ttk.Combobox(opts, textvariable=self._add_delim_var,
                     values=list(_DELIMITER_OPTIONS), state="readonly",
                     width=16).pack(side="left", padx=(4, 12))

        ttk.Label(opts, text="Mode:").pack(side="left")
        self._import_mode_var = tk.StringVar(value="Default Import")
        mode_cb = ttk.Combobox(opts, textvariable=self._import_mode_var,
                               values=list(_IMPORT_MODE_OPTIONS),
                               state="readonly", width=18)
        mode_cb.pack(side="left", padx=(4, 12))
        mode_cb.bind("<<ComboboxSelected>>", self._on_import_mode_change)

        # Type override (Unrelated only)
        type_row = ttk.Frame(frame)
        type_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self._type_override_label = ttk.Label(type_row, text="Type Override:")
        self._type_override_label.pack(side="left")
        self._type_override_var = tk.StringVar(value="Auto-Detect")
        self._type_override_cb = ttk.Combobox(
            type_row, textvariable=self._type_override_var,
            values=_TYPE_OPTIONS, state="readonly", width=16)
        self._type_override_cb.pack(side="left", padx=(4, 0))

        # Add Data button
        btn_row = ttk.Frame(frame)
        btn_row.grid(row=4, column=0, columnspan=2, sticky="e", pady=(8, 0))
        self._btn_add = ttk.Button(btn_row, text="Add Data",
                                    command=self._do_add, width=12)
        self._btn_add.pack(side="right")

        self._on_import_mode_change()

    # ---- Status Bar ----

    def _build_status_bar(self) -> None:
        bar = ttk.Frame(self, relief="sunken", borderwidth=1)
        bar.pack(fill="x", side="bottom", padx=0, pady=0)

        ttk.Label(bar, text=f"User: {self.username}").pack(side="left", padx=8, pady=3)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", pady=2)

        self._btn_pull = ttk.Button(bar, text="Pull from Master",
                                    command=self._do_pull_master)
        self._btn_pull.pack(side="left", padx=8, pady=2)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", pady=2)

        self._status_var = tk.StringVar(value="Ready")
        ttk.Label(bar, textvariable=self._status_var).pack(side="left", padx=8, pady=3)

    # ------------------------------------------------------------------
    # Mode change handlers
    # ------------------------------------------------------------------

    def _on_search_mode_change(self, event=None) -> None:
        mode = self._search_mode_var.get()
        needs_token = mode in ("GW + S", "S Only")
        state = "normal" if needs_token else "disabled"
        self._token_entry.configure(state=state)
        self._token_label.configure(foreground="" if needs_token else "#aaa")

    def _on_import_mode_change(self, event=None) -> None:
        mode = self._import_mode_var.get()
        is_unrelated = mode == "Unrelated Import"
        self._type_override_cb.configure(state="readonly" if is_unrelated else "disabled")
        self._type_override_label.configure(
            foreground="" if is_unrelated else "#aaa"
        )

    # ------------------------------------------------------------------
    # Search action
    # ------------------------------------------------------------------

    def _do_search(self) -> None:
        raw = self._search_text.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("Empty Input", "Please enter search terms.", parent=self)
            return

        mode = self._search_mode_var.get()
        search_gw = mode in ("GW + S", "GW Only")
        search_s = mode in ("GW + S", "S Only")

        s_client = None
        if search_s:
            token = self._token_var.get().strip()
            if not token:
                messagebox.showwarning("S Token Required",
                                       "Paste your S API token to search S.", parent=self)
                return
            try:
                from data.s_api import SApiClient
                s_client = SApiClient(token)
            except GWError as exc:
                messagebox.showerror(f"Error [GW{exc.code}]", exc.message, parent=self)
                return

        delim = _DELIMITER_OPTIONS[self._search_delim_var.get()]
        query_terms = parse_raw_input(raw, delim)

        self._set_status("Searching…")
        self._run_in_thread(
            self._search_worker,
            raw, delim, search_gw, search_s, s_client, query_terms,
            on_complete=self._on_search_complete,
            on_error=self._on_worker_error,
        )

    def _search_worker(self, raw, delim, search_gw, search_s, s_client, query_terms):
        gw_results, s_results = run_search(
            raw, delim, self.conn,
            s_client=s_client,
            search_gw=search_gw,
            search_s_flag=search_s,
        )
        return gw_results, s_results, query_terms, s_client

    def _on_search_complete(self, result) -> None:
        gw_results, s_results, query_terms, s_client = result
        self._set_status("Ready")
        from ui.results_window import ResultsWindow
        ResultsWindow(self, gw_results, s_results, query_terms,
                      self.conn, s_client=s_client)

    # ------------------------------------------------------------------
    # Add Data action
    # ------------------------------------------------------------------

    def _do_add(self) -> None:
        raw = self._add_text.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("Empty Input", "Please enter data to import.", parent=self)
            return

        mode = self._import_mode_var.get()
        delim = _DELIMITER_OPTIONS[self._add_delim_var.get()]

        if mode == "Unrelated Import":
            sel_type_display = self._type_override_var.get()
            sel_type = "auto" if sel_type_display == "Auto-Detect" else sel_type_display
            self._set_status("Importing…")
            self._run_in_thread(
                run_unrelated_import, raw, sel_type, self.conn, self.username, delim,
                on_complete=lambda n: self._on_import_complete(n),
                on_error=self._on_worker_error,
            )
        else:
            # Default Import — open schema detection dialog first
            rows = parse_import_input(raw, delim)
            if not rows:
                messagebox.showwarning("Empty Input", "No data rows found.", parent=self)
                return
            detected = detect_column_types(rows)
            from ui.schema_detection import SchemaDetectionDialog
            SchemaDetectionDialog(
                self, rows, detected,
                on_confirm=lambda types: self._run_default_import(rows, types),
            )

    def _run_default_import(self, rows: list, confirmed_types: list) -> None:
        self._set_status("Importing…")
        self._run_in_thread(
            run_default_import, rows, confirmed_types, self.conn, self.username,
            on_complete=lambda n: self._on_import_complete(n),
            on_error=self._on_worker_error,
        )

    def _on_import_complete(self, count: int) -> None:
        self._set_status("Ready")
        messagebox.showinfo("Import Complete",
                            f"{count} selector(s) imported successfully.", parent=self)

    # ------------------------------------------------------------------
    # Pull from Master
    # ------------------------------------------------------------------

    def _do_pull_master(self) -> None:
        self._set_status("Syncing from master…")
        self._run_in_thread(
            pull_from_master, self.conn, config.MASTER_DB_PATH,
            on_complete=self._on_pull_complete,
            on_error=self._on_worker_error,
        )

    def _on_pull_complete(self, stats: dict) -> None:
        self._set_status("Ready")
        msg = (
            f"Sync complete.\n"
            f"  Selectors added: {stats.get('selectors_added', 0)}\n"
            f"  Targets added:   {stats.get('targets_added', 0)}\n"
            f"  Norks added:     {stats.get('norks_added', 0)}"
        )
        messagebox.showinfo("Pull from Master", msg, parent=self)

    # ------------------------------------------------------------------
    # Threading
    # ------------------------------------------------------------------

    def _set_busy(self, busy: bool) -> None:
        """Disable/enable all action buttons while a background thread runs."""
        state = "disabled" if busy else "normal"
        self._btn_search.configure(state=state)
        self._btn_add.configure(state=state)
        self._btn_pull.configure(state=state)

    def _run_in_thread(self, func, *args, on_complete=None, on_error=None) -> None:
        self._set_busy(True)

        def worker():
            try:
                result = func(*args)
                self._result_queue.put(("ok", result, on_complete))
            except Exception as exc:
                self._result_queue.put(("err", exc, on_error))

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _poll_queue(self) -> None:
        """Drain the result queue on the main thread. Called repeatedly via after()."""
        try:
            while True:
                status, payload, callback = self._result_queue.get_nowait()
                self._set_busy(False)  # always re-enable before invoking callback
                if status == "ok" and callback:
                    callback(payload)
                elif status == "err" and callback:
                    callback(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _on_worker_error(self, exc: Exception) -> None:
        self._set_status("Ready")
        if isinstance(exc, GWError):
            messagebox.showerror(f"Error [GW{exc.code}]", exc.message, parent=self)
        else:
            logger.exception("Unexpected error in worker thread")
            messagebox.showerror("Unexpected Error", str(exc), parent=self)

    # ------------------------------------------------------------------
    # Status bar
    # ------------------------------------------------------------------

    def _set_status(self, text: str) -> None:
        self._status_var.set(text)
        self.update_idletasks()
