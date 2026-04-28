"""
ui/app.py — GrayWolfe main application window.
"""
from __future__ import annotations

import queue
import sqlite3
import threading
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from tkinter import messagebox, ttk

import config
from data.s_api import SApiClient
from data.sync import pull_from_master
from data.database import get_current_user
from display import strings
from display.results_window import ResultsWindow
from display.schema_detection import SchemaDetectionDialog
from src.import_data import detect_column_types, parse_import_input, run_default_import, run_unrelated_import
from src.search import parse_raw_input, run_search
from util.errors import GWError
from util.logger import get_logger

logger = get_logger(__name__)

_DELIMITER_OPTIONS = {
    "Auto": None,
    "!! (double bang)": "!!",
    "Tab": "\t",
    "Newline": "\n",
    "Comma": ",",
    "Semicolon": ";",
    "Pipe": "|",
}

_SEARCH_MODE_OPTIONS = ("GW + S", "GW Only", "S Only")
_IMPORT_MODE_OPTIONS = ("Default Import", "Unrelated Import")
_TYPE_OPTIONS = ["Auto-Detect"] + list(config.SELECTOR_TYPES)
_FONT_SIZE_OPTIONS = ("8", "9", "10", "11", "12", "14", "16", "18", "20")

_SEARCH_BORDER_COLOR = "#4472c4"
_ADD_BORDER_COLOR    = "#70ad47"
_BORDER_WIDTH        = 4
_BORDER_RADIUS       = 12


def _rounded_rect_points(x1: float, y1: float, x2: float, y2: float, r: float) -> list[float]:
    """Return polygon points for a smooth rounded rectangle (for tk.Canvas smooth polygon)."""
    return [
        x1+r, y1,   x2-r, y1,
        x2,   y1,   x2,   y1+r,
        x2,   y2-r, x2,   y2,
        x2,   y2,   x2-r, y2,
        x1+r, y2,   x1,   y2,
        x1,   y2,   x1,   y2-r,
        x1,   y1+r, x1,   y1,
        x1,   y1,   x1+r, y1,
    ]


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
        self._ask_queue: queue.Queue = queue.Queue()
        self._answer_queue: queue.Queue = queue.Queue()

        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _get_bg(self) -> str:
        """Return the actual hex background color of the root window."""
        raw = self.cget("bg")
        try:
            r, g, b = self.winfo_rgb(raw)
            return f"#{r >> 8:02x}{g >> 8:02x}{b >> 8:02x}"
        except Exception:
            return "#f0f0f0"

    def _make_rounded_border(
        self, parent: tk.Widget, color: str,
        radius: int = _BORDER_RADIUS, border_width: int = _BORDER_WIDTH, padding: int = 10,
    ) -> tuple[tk.Canvas, tk.Frame]:
        bg = self._get_bg()
        canvas = tk.Canvas(parent, bg=bg, highlightthickness=0)
        inner = tk.Frame(canvas, bg=bg)
        win_id = canvas.create_window(0, 0, anchor="nw", window=inner)

        def _redraw(event: tk.Event | None = None) -> None:
            w = canvas.winfo_width()
            h = canvas.winfo_height()
            if w < 4 or h < 4:
                return
            canvas.delete("rr")
            half = border_width / 2
            pts = _rounded_rect_points(half, half, w - half, h - half, radius)
            canvas.create_polygon(
                pts, smooth=True,
                fill=bg, outline=color, width=border_width,
                tags="rr",
            )
            inner_offset = border_width + padding
            canvas.coords(win_id, inner_offset, inner_offset)
            canvas.itemconfig(
                win_id,
                width=max(1, w - 2 * inner_offset),
                height=max(1, h - 2 * inner_offset),
            )

        canvas.bind("<Configure>", _redraw)
        return canvas, inner

    def _build_ui(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)
        self._build_search_tab()
        self._build_add_tab()
        self._build_status_bar()

    # ---- Search Tab ----

    def _build_search_tab(self) -> None:
        tab_outer = ttk.Frame(self.notebook)
        self.notebook.add(tab_outer, text="Search")

        # Disclaimer at top
        ttk.Label(
            tab_outer,
            text=strings.DISCLAIMER_SEARCH,
            foreground="#555",
            font=("TkDefaultFont", 9, "italic"),
            wraplength=720,
            justify="left",
        ).pack(side="top", anchor="w", padx=8, pady=(6, 2))

        # Rounded border canvas
        canvas, inner = self._make_rounded_border(tab_outer, _SEARCH_BORDER_COLOR)
        canvas.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        inner.columnconfigure(0, weight=1)
        inner.rowconfigure(3, weight=1)

        # Row 0 — Search Location label
        ttk.Label(inner, text="Search Location", font=("TkDefaultFont", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )

        # Row 1 — loc_row: mode combobox | token entry (expand) | eye | help
        loc_row = tk.Frame(inner, bg=self._get_bg())
        loc_row.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        loc_row.columnconfigure(1, weight=1)

        self._search_mode_var = tk.StringVar(value="GW + S")
        search_mode_cb = ttk.Combobox(
            loc_row, textvariable=self._search_mode_var,
            values=list(_SEARCH_MODE_OPTIONS),
            state="readonly", width=16,
        )
        search_mode_cb.grid(row=0, column=0, sticky="w")
        search_mode_cb.bind("<<ComboboxSelected>>", self._on_search_mode_change)

        self._token_var = tk.StringVar()
        self._token_entry = ttk.Entry(loc_row, textvariable=self._token_var, width=40)
        self._token_entry.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self._setup_entry_placeholder(self._token_entry, strings.TOKEN_PLACEHOLDER)
        self._token_visible = False

        self._token_toggle_btn = ttk.Button(
            loc_row, text="\U0001F441", width=3,
            command=self._toggle_token_visibility,
        )
        self._token_toggle_btn.grid(row=0, column=2, padx=(4, 0))

        self._token_help_btn = ttk.Button(
            loc_row, text="?", width=2,
            command=self._open_token_help,
        )
        self._token_help_btn.grid(row=0, column=3, padx=(2, 0))

        self._token_entry._on_placeholder_restore = lambda: (  # type: ignore[attr-defined]
            setattr(self, "_token_visible", False)
            or self._token_toggle_btn.configure(text="\U0001F441")
        )

        # Row 2 — "ITW Selectors to Search" label (left) + Delimiter + Font (right)
        lbl_row = tk.Frame(inner, bg=self._get_bg())
        lbl_row.grid(row=2, column=0, sticky="ew", pady=(0, 2))
        lbl_row.columnconfigure(0, weight=1)

        ttk.Label(lbl_row, text="ITW Selectors to Search", font=("TkDefaultFont", 10, "bold")).grid(
            row=0, column=0, sticky="w"
        )

        controls_frame = tk.Frame(lbl_row, bg=self._get_bg())
        controls_frame.grid(row=0, column=1, sticky="e")

        ttk.Label(controls_frame, text="Delimiter:").pack(side="left")
        self._search_delim_var = tk.StringVar(value="Auto")
        ttk.Combobox(
            controls_frame, textvariable=self._search_delim_var,
            values=list(_DELIMITER_OPTIONS), state="readonly", width=14,
        ).pack(side="left", padx=(4, 12))

        ttk.Label(controls_frame, text="Font:").pack(side="left")
        self._search_font_size_var = tk.StringVar(value="11")
        search_font_cb = ttk.Combobox(
            controls_frame, textvariable=self._search_font_size_var,
            values=_FONT_SIZE_OPTIONS, state="readonly", width=4,
        )
        search_font_cb.pack(side="left", padx=(2, 0))
        search_font_cb.bind("<<ComboboxSelected>>", self._on_search_font_size_change)

        self._search_font_revealed = True

        # Row 3 — text area + scrollbar
        self._search_font = tkfont.nametofont("TkDefaultFont").copy()
        self._search_font.configure(size=11)
        self._search_text = tk.Text(inner, height=8, wrap="none", font=self._search_font)
        self._search_text.grid(row=3, column=0, sticky="nsew", pady=(0, 6))

        vsb = ttk.Scrollbar(inner, orient="vertical", command=self._search_text.yview)
        vsb.grid(row=3, column=1, sticky="ns", pady=(0, 6))
        self._search_text["yscrollcommand"] = vsb.set
        self._search_text.bind(
            "<KeyRelease>",
            lambda _e: self._schedule_delim_detect(
                self._search_text, self._search_delim_var, "_search_delim_after_id"
            ),
            add=True,
        )
        self._search_text.bind(
            "<<Paste>>",
            lambda _e: self._schedule_delim_detect(
                self._search_text, self._search_delim_var, "_search_delim_after_id"
            ),
            add=True,
        )
        self._setup_placeholder(self._search_text, lambda: strings.SEARCH_PLACEHOLDER)

        # Row 4 — buttons
        btn_row = tk.Frame(inner, bg=self._get_bg())
        btn_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))

        self._btn_search = tk.Button(
            btn_row, text="Submit",
            bg=_SEARCH_BORDER_COLOR, fg="white",
            activebackground="#2f5496",
            relief="flat", padx=12, pady=6,
            command=self._do_search,
        )
        self._btn_search.pack(side="left", padx=(0, 6))

        self._btn_clear_search = ttk.Button(
            btn_row, text="Clear", command=self._clear_search, width=12,
        )
        self._btn_clear_search.pack(side="left")

        self._on_search_mode_change()

    # ---- Add Data Tab ----

    def _build_add_tab(self) -> None:
        tab_outer = ttk.Frame(self.notebook)
        self.notebook.add(tab_outer, text="Add Selectors")

        # Disclaimer at top
        ttk.Label(
            tab_outer,
            text=strings.DISCLAIMER_ADD,
            foreground="#555",
            font=("TkDefaultFont", 9, "italic"),
            wraplength=720,
            justify="left",
        ).pack(side="top", anchor="w", padx=8, pady=(6, 2))

        # Rounded border canvas
        canvas, inner = self._make_rounded_border(tab_outer, _ADD_BORDER_COLOR)
        canvas.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        inner.columnconfigure(0, weight=1)
        inner.rowconfigure(4, weight=1)

        # Row 0 — two-col header: "Import Mode" | "Selector Type"
        ttk.Label(inner, text="Import Mode", font=("TkDefaultFont", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 2)
        )
        ttk.Label(inner, text="Selector Type", font=("TkDefaultFont", 10, "bold")).grid(
            row=0, column=1, sticky="w", pady=(0, 2), padx=(8, 0)
        )

        # Row 1 — import_mode_cb | type_override_cb
        self._import_mode_var = tk.StringVar(value="Default Import")
        mode_cb = ttk.Combobox(
            inner, textvariable=self._import_mode_var,
            values=list(_IMPORT_MODE_OPTIONS),
            state="readonly", width=18,
        )
        mode_cb.grid(row=1, column=0, sticky="w", pady=(0, 6))
        mode_cb.bind("<<ComboboxSelected>>", self._on_import_mode_change)

        self._type_override_var = tk.StringVar(value="Auto-Detect")
        self._type_override_cb = ttk.Combobox(
            inner, textvariable=self._type_override_var,
            values=_TYPE_OPTIONS, state="readonly", width=16,
        )
        self._type_override_cb.grid(row=1, column=1, sticky="w", pady=(0, 6), padx=(8, 0))
        # Keep label reference for _on_import_mode_change foreground logic
        self._type_override_label = ttk.Label(inner, text="")  # hidden placeholder for compat

        # Row 2 — also_s_row: checkbox | token label | token entry | eye | ? | Delimiter | Font
        also_row = tk.Frame(inner, bg=self._get_bg())
        also_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        also_row.columnconfigure(2, weight=1)

        self._s_add_var = tk.BooleanVar(value=False)
        self._s_add_cb = ttk.Checkbutton(
            also_row, text="Also Search S?",
            variable=self._s_add_var, command=self._on_s_add_toggle,
        )
        self._s_add_cb.grid(row=0, column=0, sticky="w")

        self._s_add_token_label = ttk.Label(also_row, text="S Token:")
        self._s_add_token_label.grid(row=0, column=1, padx=(10, 0), sticky="w")

        self._s_add_token_var = tk.StringVar()
        self._s_add_token_entry = ttk.Entry(also_row, textvariable=self._s_add_token_var, width=30)
        self._s_add_token_entry.grid(row=0, column=2, sticky="ew", padx=(4, 0))
        self._setup_entry_placeholder(self._s_add_token_entry, strings.TOKEN_PLACEHOLDER)
        self._s_add_token_visible = False

        self._s_add_token_toggle_btn = ttk.Button(
            also_row, text="\U0001F441", width=3,
            command=self._toggle_s_add_token_visibility,
        )
        self._s_add_token_toggle_btn.grid(row=0, column=3, padx=(4, 0))

        s_add_help_btn = ttk.Button(also_row, text="?", width=2, command=lambda: None)
        s_add_help_btn.grid(row=0, column=4, padx=(2, 0))

        self._s_add_token_entry._on_placeholder_restore = lambda: (  # type: ignore[attr-defined]
            setattr(self, "_s_add_token_visible", False)
            or self._s_add_token_toggle_btn.configure(text="\U0001F441")
        )

        self._add_font_revealed = True

        # Row 3 — title (left) + Delimiter + Font (right)
        lbl_row = tk.Frame(inner, bg=self._get_bg())
        lbl_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 2))
        lbl_row.columnconfigure(0, weight=1)

        ttk.Label(lbl_row, text="ITW Selectors to Add to GrayWolfe", font=("TkDefaultFont", 10, "bold")).grid(
            row=0, column=0, sticky="w"
        )

        controls_frame = tk.Frame(lbl_row, bg=self._get_bg())
        controls_frame.grid(row=0, column=1, sticky="e")

        ttk.Label(controls_frame, text="Delimiter:").pack(side="left")
        self._add_delim_var = tk.StringVar(value="Auto")
        ttk.Combobox(
            controls_frame, textvariable=self._add_delim_var,
            values=list(_DELIMITER_OPTIONS), state="readonly", width=14,
        ).pack(side="left", padx=(4, 12))

        ttk.Label(controls_frame, text="Font:").pack(side="left")
        self._add_font_size_var = tk.StringVar(value="11")
        add_font_cb = ttk.Combobox(
            controls_frame, textvariable=self._add_font_size_var,
            values=_FONT_SIZE_OPTIONS, state="readonly", width=4,
        )
        add_font_cb.pack(side="left", padx=(2, 0))
        add_font_cb.bind("<<ComboboxSelected>>", self._on_add_font_size_change)

        # Row 4 — text area + scrollbar
        self._add_current_placeholder = strings.ADD_DEFAULT_PLACEHOLDER
        self._add_font = tkfont.nametofont("TkDefaultFont").copy()
        self._add_font.configure(size=11)
        self._add_text = tk.Text(inner, height=8, wrap="none", font=self._add_font)
        self._add_text.grid(row=4, column=0, sticky="nsew", pady=(0, 6))

        vsb = ttk.Scrollbar(inner, orient="vertical", command=self._add_text.yview)
        vsb.grid(row=4, column=1, sticky="ns", pady=(0, 6))
        self._add_text["yscrollcommand"] = vsb.set
        self._add_text.bind(
            "<KeyRelease>",
            lambda _e: self._schedule_delim_detect(
                self._add_text, self._add_delim_var, "_add_delim_after_id"
            ),
            add=True,
        )
        self._add_text.bind(
            "<<Paste>>",
            lambda _e: self._schedule_delim_detect(
                self._add_text, self._add_delim_var, "_add_delim_after_id"
            ),
            add=True,
        )
        self._setup_placeholder(self._add_text, lambda: self._add_current_placeholder)

        # Row 5 — buttons
        btn_row = tk.Frame(inner, bg=self._get_bg())
        btn_row.grid(row=5, column=0, columnspan=2, sticky="w", pady=(4, 0))

        self._btn_add = tk.Button(
            btn_row, text="Submit",
            bg=_ADD_BORDER_COLOR, fg="white",
            activebackground="#507e32",
            relief="flat", padx=12, pady=6,
            command=self._do_add,
        )
        self._btn_add.pack(side="left", padx=(0, 6))

        self._btn_clear_add = ttk.Button(
            btn_row, text="Clear", command=self._clear_add, width=12,
        )
        self._btn_clear_add.pack(side="left")

        # Footer label
        ttk.Label(
            tab_outer,
            text="Tool will auto search new inputs and make connections (inshaAllah)",
            foreground="#777",
            font=("TkDefaultFont", 9, "italic"),
        ).pack(side="bottom", pady=(4, 6))

        self._on_import_mode_change()
        self._on_s_add_toggle()

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
        self._token_toggle_btn.configure(state=state)
        self._token_help_btn.configure(state=state)

    def _on_import_mode_change(self, event=None) -> None:
        mode = self._import_mode_var.get()
        is_unrelated = mode == "Unrelated Import"
        self._type_override_cb.configure(state="readonly" if is_unrelated else "disabled")
        new_ph = strings.ADD_UNRELATED_PLACEHOLDER if is_unrelated else strings.ADD_DEFAULT_PLACEHOLDER
        self._add_current_placeholder = new_ph
        if self._add_text.tag_ranges("placeholder"):
            self._add_text.delete("1.0", "end")
            self._add_text.insert("1.0", new_ph)
            self._add_text.tag_add("placeholder", "1.0", "end-1c")

    def _on_s_add_toggle(self) -> None:
        """Enable/disable S token row based on checkbox state."""
        enabled = self._s_add_var.get()
        state = "normal" if enabled else "disabled"
        self._s_add_token_entry.configure(state=state)
        self._s_add_token_toggle_btn.configure(state=state)
        self._s_add_token_label.configure(foreground="" if enabled else "#aaa")

    def _toggle_token_visibility(self) -> None:
        self._token_visible = not self._token_visible
        placeholder_active = getattr(self._token_entry, "_placeholder_active", False)
        self._token_entry.configure(show="" if (self._token_visible or placeholder_active) else "*")
        self._token_toggle_btn.configure(
            text="\U0001F512" if self._token_visible else "\U0001F441"
        )

    def _toggle_s_add_token_visibility(self) -> None:
        self._s_add_token_visible = not self._s_add_token_visible
        placeholder_active = getattr(self._s_add_token_entry, "_placeholder_active", False)
        self._s_add_token_entry.configure(show="" if (self._s_add_token_visible or placeholder_active) else "*")
        self._s_add_token_toggle_btn.configure(
            text="\U0001F512" if self._s_add_token_visible else "\U0001F441"
        )

    def _on_search_font_size_change(self, event=None) -> None:
        self._search_font.configure(size=int(self._search_font_size_var.get()))

    def _on_add_font_size_change(self, event=None) -> None:
        self._add_font.configure(size=int(self._add_font_size_var.get()))

    def _detect_input_delimiter(self, text: str) -> str | None:
        """Detect the most likely column delimiter in *text*.

        Candidates checked in priority order: semicolon, pipe, tab, comma.
        Algorithm mirrors VBA DetectInputDelimiter:
          - Split into up to 20 non-empty rows.
          - Fewer than 4 rows: first candidate found in any row wins.
          - 4+ rows: candidate must appear in >74% of rows AND average
            hits-per-row > 1.
          - Returns the winning delimiter character, or None if undecided.
        """
        candidates = [";", "|", "\t", ","]
        rows = [r for r in text.split("\n") if r.strip()][:20]
        if not rows:
            return None
        if len(rows) < 4:
            for cand in candidates:
                if any(cand in row for row in rows):
                    return cand
            return None
        for cand in candidates:
            hit_rows = [row for row in rows if cand in row]
            if not hit_rows:
                continue
            hit_ratio = len(hit_rows) / len(rows)
            avg_hits = sum(row.count(cand) for row in hit_rows) / len(hit_rows)
            if hit_ratio > 0.74 and avg_hits > 1:
                return cand
        return None

    def _schedule_delim_detect(
        self,
        widget: tk.Text,
        delim_var: tk.StringVar,
        after_id_attr: str,
    ) -> None:
        """Cancel any pending detection for this widget and schedule a fresh one in 500 ms.

        *after_id_attr* is the name of an instance attribute that stores the
        pending `after()` handle. Lazily initialised via getattr — no __init__
        assignment is needed.
        """
        pending = getattr(self, after_id_attr, None)
        if pending is not None:
            self.after_cancel(pending)
        handle = self.after(
            500,
            lambda: self._run_delim_detect(widget, delim_var, after_id_attr),
        )
        setattr(self, after_id_attr, handle)

    def _run_delim_detect(
        self,
        widget: tk.Text,
        delim_var: tk.StringVar,
        after_id_attr: str,
    ) -> None:
        """Detect delimiter in *widget* content and update *delim_var*.

        Empty/placeholder content resets the combobox to "Auto".
        Detected value is looked up in _DELIMITER_OPTIONS (None → "Auto").
        """
        setattr(self, after_id_attr, None)
        text = self._get_text_widget_value(widget)
        if not text:
            delim_var.set("Auto")
            return
        detected = self._detect_input_delimiter(text)
        # Build reverse map excluding the None/"Auto" entry so None-detected
        # falls through to the default "Auto".
        reverse = {v: k for k, v in _DELIMITER_OPTIONS.items() if v is not None}
        delim_var.set(reverse.get(detected, "Auto"))

    def _setup_placeholder(self, widget: tk.Text, get_text) -> None:
        """Insert placeholder text (gray) that clears on focus and restores when empty."""
        widget.tag_configure("placeholder", foreground="#aaa")
        widget.insert("1.0", get_text())
        widget.tag_add("placeholder", "1.0", "end-1c")

        def on_focus_in(_event):
            if widget.tag_ranges("placeholder"):
                widget.delete("1.0", "end")

        def on_focus_out(_event):
            if not widget.get("1.0", "end").strip():
                ph = get_text()
                widget.insert("1.0", ph)
                widget.tag_add("placeholder", "1.0", "end-1c")

        def on_key(_event):
            if widget.tag_ranges("placeholder"):
                widget.delete("1.0", "end")

        widget.bind("<FocusIn>", on_focus_in, add=True)
        widget.bind("<FocusOut>", on_focus_out, add=True)
        widget.bind("<Key>", on_key, add=True)

    def _setup_entry_placeholder(self, entry: ttk.Entry, placeholder: str) -> None:
        """Insert placeholder text (gray) in an Entry that clears on focus and restores when empty.

        For password-style entries: shows placeholder unmasked (show=""), switches to
        show="*" on first interaction, restores placeholder on empty focus-out.
        """
        entry._placeholder_active = True  # type: ignore[attr-defined]
        entry.insert(0, placeholder)
        entry.configure(foreground="#aaa", show="")

        def on_focus_in(_event):
            if getattr(entry, "_placeholder_active", False):
                entry.delete(0, "end")
                entry.configure(foreground="", show="*")
                entry._placeholder_active = False  # type: ignore[attr-defined]

        def on_focus_out(_event):
            if not entry.get().strip():
                entry.delete(0, "end")
                entry.insert(0, placeholder)
                entry.configure(foreground="#aaa", show="")
                entry._placeholder_active = True  # type: ignore[attr-defined]
                restore = getattr(entry, "_on_placeholder_restore", None)
                if callable(restore):
                    restore()

        entry.bind("<FocusIn>", on_focus_in, add=True)
        entry.bind("<FocusOut>", on_focus_out, add=True)

    def _get_text_widget_value(self, widget: tk.Text) -> str:
        """Return text content; returns empty string if placeholder is currently active."""
        if widget.tag_ranges("placeholder"):
            return ""
        return widget.get("1.0", "end").strip()

    def _clear_search(self) -> None:
        self._search_text.delete("1.0", "end")
        self._search_text.insert("1.0", strings.SEARCH_PLACEHOLDER)
        self._search_text.tag_add("placeholder", "1.0", "end-1c")

    def _clear_add(self) -> None:
        self._add_text.delete("1.0", "end")
        self._add_text.insert("1.0", self._add_current_placeholder)
        self._add_text.tag_add("placeholder", "1.0", "end-1c")

    def _open_token_help(self) -> None:
        if config.S_TOKEN_HELP_URL:
            webbrowser.open(config.S_TOKEN_HELP_URL)

    # ------------------------------------------------------------------
    # Search action
    # ------------------------------------------------------------------

    def _do_search(self) -> None:
        raw = self._get_text_widget_value(self._search_text)
        if not raw:
            messagebox.showwarning("Empty Input", "Please enter search terms.", parent=self)
            return

        mode = self._search_mode_var.get()
        search_gw = mode in ("GW + S", "GW Only")
        search_s = mode in ("GW + S", "S Only")

        s_client = None
        if search_s:
            token = self._token_var.get().strip()
            if not token or getattr(self._token_entry, "_placeholder_active", False):
                messagebox.showwarning("S Token Required",
                                       "Paste your S API token to search S.", parent=self)
                return
            try:
                s_client = SApiClient(token)
            except GWError as exc:
                messagebox.showerror(f"Error [GW{exc.code}]", exc.message, parent=self)
                return

        delim = _DELIMITER_OPTIONS[self._search_delim_var.get()]
        query_terms = parse_raw_input(raw, delim)

        # Drain any stale answer left over from a previous timed-out ask_cb dialog
        while True:
            try:
                self._answer_queue.get_nowait()
            except queue.Empty:
                break

        def progress_cb(selector: str, idx: int, total: int) -> None:
            self._result_queue.put(("progress", strings.rate_limit_text(selector, idx, total), None))

        def ask_cb(selector: str, num_found: int) -> bool:
            self._ask_queue.put((selector, num_found))
            try:
                return self._answer_queue.get(timeout=30)
            except queue.Empty:
                return False  # treat timeout as "skip"

        self._set_status("Searching…")
        self._run_in_thread(
            self._search_worker,
            raw, delim, search_gw, search_s, s_client, query_terms, progress_cb, ask_cb,
            on_complete=self._on_search_complete,
            on_error=self._on_worker_error,
        )

    def _search_worker(self, raw, delim, search_gw, search_s, s_client, query_terms, progress_cb, ask_cb):
        gw_results, s_results = run_search(
            raw, delim, self.conn,
            s_client=s_client,
            search_gw=search_gw,
            search_s_flag=search_s,
            progress_cb=progress_cb,
            ask_cb=ask_cb,
        )
        return gw_results, s_results, query_terms, s_client

    def _on_search_complete(self, result) -> None:
        gw_results, s_results, query_terms, s_client = result
        self._set_status("Ready")
        ResultsWindow(self, gw_results, s_results, query_terms,
                      self.conn, s_client=s_client)

    # ------------------------------------------------------------------
    # Add Data action
    # ------------------------------------------------------------------

    def _do_add(self) -> None:
        raw = self._get_text_widget_value(self._add_text)
        if not raw:
            messagebox.showwarning("Empty Input", "Please enter data to import.", parent=self)
            return

        mode = self._import_mode_var.get()
        delim = _DELIMITER_OPTIONS[self._add_delim_var.get()]

        # Resolve S client if "Also Search S?" is checked
        s_client = None
        if self._s_add_var.get():
            token = self._s_add_token_var.get().strip()
            if not token or getattr(self._s_add_token_entry, "_placeholder_active", False):
                messagebox.showwarning(
                    "S Token Required",
                    "Paste your S API token to search S after import.",
                    parent=self,
                )
                return
            try:
                s_client = SApiClient(token)
            except GWError as exc:
                messagebox.showerror(f"Error [GW{exc.code}]", exc.message, parent=self)
                return

        if mode == "Unrelated Import":
            sel_type_display = self._type_override_var.get()
            sel_type = "auto" if sel_type_display == "Auto-Detect" else sel_type_display
            self._set_status("Importing…")
            self._run_in_thread(
                run_unrelated_import, raw, sel_type, self.conn, self.username, delim,
                on_complete=lambda result: self._on_import_complete(result, raw=raw, delim=delim, s_client=s_client),
                on_error=self._on_worker_error,
            )
        else:
            # Default Import — open schema detection dialog first
            rows = parse_import_input(raw, delim)
            if not rows:
                messagebox.showwarning("Empty Input", "No data rows found.", parent=self)
                return
            detected = detect_column_types(rows)
            SchemaDetectionDialog(
                self, rows, detected,
                on_confirm=lambda types: self._run_default_import(rows, types, raw=raw, delim=delim, s_client=s_client),
            )

    def _run_default_import(
        self,
        rows: list,
        confirmed_types: list,
        raw: str = "",
        delim: str | None = None,
        s_client=None,
    ) -> None:
        self._set_status("Importing…")
        self._run_in_thread(
            run_default_import, rows, confirmed_types, self.conn, self.username,
            on_complete=lambda result: self._on_import_complete(result, raw=raw, delim=delim, s_client=s_client),
            on_error=self._on_worker_error,
        )

    def _on_import_complete(
        self,
        result: tuple[int, int],
        raw: str = "",
        delim: str | None = None,
        s_client=None,
    ) -> None:
        self._set_status("Ready")
        inserted, skipped = result
        submitted = inserted + skipped
        messagebox.showinfo(
            "Import Complete",
            strings.import_result_text(submitted, inserted, skipped),
            parent=self,
        )
        if s_client is not None and raw:
            self._do_s_search_after_import(raw, delim, s_client)

    def _do_s_search_after_import(
        self,
        raw: str,
        delim: str | None,
        s_client,
    ) -> None:
        """Run an S-only search on the just-imported data and open a ResultsWindow.

        Called synchronously from _on_import_complete (main thread). The
        messagebox in _on_import_complete blocks until the user clicks OK, so
        by the time this runs the import dialog is already dismissed.
        """
        query_terms = parse_raw_input(raw, delim)

        # Drain any stale answer left over from a previous timed-out ask_cb dialog
        while True:
            try:
                self._answer_queue.get_nowait()
            except queue.Empty:
                break

        def progress_cb(selector: str, idx: int, total: int) -> None:
            self._result_queue.put(("progress", strings.rate_limit_text(selector, idx, total), None))

        def ask_cb(selector: str, num_found: int) -> bool:
            self._ask_queue.put((selector, num_found))
            try:
                return self._answer_queue.get(timeout=30)
            except queue.Empty:
                return False  # treat timeout as "skip"

        def _s_search_worker():
            return run_search(
                raw, delim, self.conn, s_client, False, True,
                progress_cb=progress_cb,
                ask_cb=ask_cb,
            )

        self._set_status("Searching S…")
        self._run_in_thread(
            _s_search_worker,
            on_complete=lambda result: self._on_s_search_after_import_complete(
                result, query_terms, s_client
            ),
            on_error=self._on_worker_error,
        )

    def _on_s_search_after_import_complete(
        self,
        result: tuple,
        query_terms: list[str],
        s_client,
    ) -> None:
        self._set_status("Ready")
        gw_results, s_results = result
        ResultsWindow(self, gw_results, s_results, query_terms, self.conn, s_client=s_client)

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
        self._btn_clear_search.configure(state=state)
        self._btn_clear_add.configure(state=state)
        self._s_add_cb.configure(state=state)
        if busy:
            self._s_add_token_entry.configure(state="disabled")
            self._s_add_token_toggle_btn.configure(state="disabled")
        else:
            # Restore token entry state based on current checkbox value
            self._on_s_add_toggle()

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
                if status == "progress":
                    if payload:
                        self._set_status(payload)
                else:
                    self._set_busy(False)  # always re-enable before invoking callback
                    if status == "ok" and callback:
                        callback(payload)
                    elif status == "err" and callback:
                        callback(payload)
        except queue.Empty:
            pass
        try:
            selector, num_found = self._ask_queue.get_nowait()
            answer = messagebox.askyesno(
                "Large Search",
                strings.search_check_text(selector, num_found),
                parent=self,
            )
            self._answer_queue.put(answer)
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
