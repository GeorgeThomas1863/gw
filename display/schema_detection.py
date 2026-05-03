"""
ui/schema_detection.py — Modal dialog for confirming column types during default import.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable

from config import SELECTOR_TYPES
from display import strings
from models import ColumnTypeInfo, SelectorType
from util.logger import get_logger

logger = get_logger(__name__)

# Column type choices shown in each combobox.
_TYPE_CHOICES = list(SELECTOR_TYPES) + ["null"]


class SchemaDetectionDialog(tk.Toplevel):
    """Modal dialog that shows auto-detected column types and lets the user override them."""

    def __init__(
        self,
        parent: tk.Widget,
        rows: list[list[str]],
        detected_types: list[ColumnTypeInfo],
        on_confirm: Callable[[list[str]], None],
    ) -> None:
        super().__init__(parent)
        self.rows = rows
        self.detected_types = detected_types
        self.on_confirm = on_confirm

        self.title("Confirm Column Types")
        self.resizable(True, False)
        self.grab_set()  # modal

        self._combos: list[ttk.Combobox] = []
        self._build_ui()

        self.transient(parent)
        self.wait_visibility()
        self.focus_set()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 4}
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        # --- Preview table ---
        preview_label = ttk.Label(main, text="Data Preview (first 5 rows):",
                                  font=("", 9, "bold"))
        preview_label.pack(anchor="w", pady=(0, 4))

        num_cols = max((len(r) for r in self.rows), default=1)
        col_ids = [str(i) for i in range(num_cols)]

        preview_frame = ttk.Frame(main, relief="sunken", borderwidth=1)
        preview_frame.pack(fill="x", pady=(0, 8))

        tree = ttk.Treeview(preview_frame, columns=col_ids, show="headings", height=5)
        for i, cid in enumerate(col_ids):
            tree.heading(cid, text=f"Col {i + 1}")
            tree.column(cid, width=120, anchor="w", stretch=True)

        for row in self.rows[:5]:
            padded = list(row) + [""] * (num_cols - len(row))
            tree.insert("", "end", values=padded)

        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=hsb.set)
        tree.pack(fill="x")
        hsb.pack(fill="x")

        # --- Column type selectors ---
        type_frame = ttk.LabelFrame(main, text="Column Types", padding=8)
        type_frame.pack(fill="x", pady=(0, 8))

        for col_idx, info in enumerate(self.detected_types):
            if col_idx >= 6:
                break  # cap at 6 columns for UI clarity

            col_type = info.selector_type.value if info.selector_type else "other"
            confidence = info.confidence
            pct = f"{confidence * 100:.0f}%"

            row_frame = ttk.Frame(type_frame)
            row_frame.pack(fill="x", pady=2)

            ttk.Label(row_frame, text=f"Column {col_idx + 1}:",
                      width=12, anchor="w").pack(side="left")

            cb = ttk.Combobox(row_frame, values=_TYPE_CHOICES, state="readonly", width=14)
            cb.set(col_type if col_type in _TYPE_CHOICES else "other")
            cb.pack(side="left", padx=(4, 8))
            self._combos.append(cb)

            ttk.Label(row_frame, text=f"({pct} confident)",
                      foreground="#666").pack(side="left")

        # --- Buttons ---
        btn_frame = ttk.Frame(main)
        btn_frame.pack(pady=(4, 0))
        ttk.Button(btn_frame, text="Submit", command=self._submit,
                   width=10).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy,
                   width=10).pack(side="left", padx=8)

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def _submit(self) -> None:
        confirmed_types = [cb.get() for cb in self._combos]
        logger.debug("Schema confirmed: %s", confirmed_types)

        # typeNull check: columns where GW couldn't detect a type (confidence == 0.0)
        for col_idx, info in enumerate(self.detected_types):
            if col_idx >= len(self._combos):
                break
            if info.confidence == 0.0 and confirmed_types[col_idx] != "null":
                sample = next(
                    (r[col_idx] for r in self.rows if col_idx < len(r) and r[col_idx].strip()),
                    f"Column {col_idx + 1}"
                )
                yes = messagebox.askyesno(
                    "Undetected Type", strings.type_null_text(sample), parent=self
                )
                if not yes:
                    return  # user wants to fix it; dialog stays open
                confirmed_types[col_idx] = "null"  # user chose to skip this column

        # typeWrong check: columns where the user overrode a high-confidence detection (>= 50%)
        for col_idx, (cb, info) in enumerate(zip(self._combos, self.detected_types)):
            if cb.get() != (info.selector_type.value if info.selector_type else "other") and info.confidence >= 0.5:
                sample = next(
                    (r[col_idx] for r in self.rows if col_idx < len(r) and r[col_idx].strip()),
                    f"Column {col_idx + 1}"
                )
                yes = messagebox.askyesno(
                    "Type Override", strings.type_wrong_text(sample, cb.get()), parent=self
                )
                if not yes:
                    cb.set(info.selector_type.value if info.selector_type else "other")  # revert the combo to detected type
                    return               # abort submit; dialog stays open for user to fix

        try:
            self.on_confirm(confirmed_types)
        except Exception as exc:
            logger.exception("on_confirm callback raised")
            messagebox.showerror("Import Error", str(exc), parent=self)
            return

        self.destroy()
