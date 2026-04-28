"""
ui/target_details.py — Window for viewing and editing a single target.
"""
from __future__ import annotations

import sqlite3
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from typing import Callable, Optional

from data.database import get_current_user, get_target, update_target, update_target_id
from display import strings
from util.errors import GWError
from util.logger import get_logger

logger = get_logger(__name__)


class TargetDetailsWindow(tk.Toplevel):
    """Non-modal window for viewing/editing a target and its selectors."""

    def __init__(
        self,
        parent: tk.Widget,
        conn: sqlite3.Connection,
        target_id: str,
        on_save_complete: Optional[Callable[[str], None]] = None,
    ) -> None:
        super().__init__(parent)
        self.conn = conn
        self.target_id = target_id
        self.on_save_complete = on_save_complete

        self.title(f"Target Details — {target_id}")
        self.minsize(520, 440)
        self.resizable(True, True)

        self._build_ui()
        self._load_target()

        self.transient(parent)
        self.focus_set()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)

        # --- Metadata fields ---
        row = 0
        ttk.Label(main, text="Target ID:").grid(row=row, column=0, sticky="w", pady=3)
        self._id_var = tk.StringVar()
        ttk.Entry(main, textvariable=self._id_var, state="readonly",
                  width=22).grid(row=row, column=1, sticky="w", pady=3)
        ttk.Button(main, text="Change ID", command=self._change_target_id, width=9).grid(
            row=0, column=2, sticky="w", padx=(4, 0), pady=3
        )

        row += 1
        ttk.Label(main, text="Target Name:").grid(row=row, column=0, sticky="w", pady=3)
        self._name_var = tk.StringVar()
        ttk.Entry(main, textvariable=self._name_var, width=40).grid(
            row=row, column=1, sticky="ew", pady=3)

        row += 1
        ttk.Label(main, text="Case Number:").grid(row=row, column=0, sticky="w", pady=3)
        self._case_var = tk.StringVar()
        ttk.Entry(main, textvariable=self._case_var, width=30).grid(
            row=row, column=1, sticky="ew", pady=3)

        row += 1
        ttk.Label(main, text="Laptop Count:").grid(row=row, column=0, sticky="w", pady=3)
        self._laptop_var = tk.IntVar(value=0)
        ttk.Spinbox(main, textvariable=self._laptop_var, from_=0, to=999,
                    width=8).grid(row=row, column=1, sticky="w", pady=3)

        # --- Selectors treeview ---
        row += 1
        ttk.Separator(main, orient="horizontal").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(8, 4))

        row += 1
        ttk.Label(main, text="Selectors", font=("", 10, "bold")).grid(
            row=row, column=0, columnspan=2, sticky="w")

        row += 1
        tree_frame = ttk.Frame(main)
        tree_frame.grid(row=row, column=0, columnspan=2, sticky="nsew", pady=(4, 0))
        main.rowconfigure(row, weight=1)

        cols = ("type", "selector", "count")
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=10)
        self._tree.heading("type", text="Type")
        self._tree.heading("selector", text="Selector")
        self._tree.heading("count", text="Count")
        self._tree.column("type", width=90, anchor="w")
        self._tree.column("selector", width=300, anchor="w")
        self._tree.column("count", width=60, anchor="center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # --- Buttons ---
        row += 1
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=(10, 0), sticky="e")

        ttk.Button(btn_frame, text="Merge Target",
                   command=self._open_merge).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Reset",
                   command=lambda: self._load_target(confirm=True)).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Save",
                   command=self._save).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Close",
                   command=self.destroy).pack(side="left", padx=4)

    # ------------------------------------------------------------------
    # Data load / save
    # ------------------------------------------------------------------

    def _load_target(self, confirm: bool = False) -> None:
        if confirm:
            proceed = messagebox.askyesno(
                "Reset Target", strings.CONFIRM_RESET_TARGETS, parent=self
            )
            if not proceed:
                return

        target = get_target(self.conn, self.target_id)
        if target is None:
            messagebox.showerror("Not Found",
                                 f"Target {self.target_id!r} not found.",
                                 parent=self)
            self.destroy()
            return

        self._id_var.set(target.get("target_id", ""))
        self._name_var.set(target.get("target_name") or "")
        self._case_var.set(target.get("case_number") or "")
        self._laptop_var.set(int(target.get("laptop_count") or 0))

        self._load_selectors()

    def _load_selectors(self) -> None:
        self._tree.delete(*self._tree.get_children())
        rows = self.conn.execute(
            """
            SELECT selector_type, selector, COUNT(*) AS cnt
            FROM selectors
            WHERE target_id = ?
            GROUP BY selector_clean
            ORDER BY cnt DESC, selector_type, selector
            """,
            (self.target_id,),
        ).fetchall()
        for r in rows:
            self._tree.insert("", "end", values=(r[0], r[1], r[2]))

    def _save(self) -> None:
        try:
            update_target(
                self.conn,
                self.target_id,
                {
                    "target_name": self._name_var.get().strip() or None,
                    "case_number": self._case_var.get().strip() or None,
                    "laptop_count": self._laptop_var.get(),
                },
                get_current_user(),
            )
        except GWError as exc:
            messagebox.showerror(f"Error [GW{exc.code}]", exc.message, parent=self)
            return
        except Exception as exc:
            logger.exception("Unexpected error saving target")
            messagebox.showerror("Unexpected Error", str(exc), parent=self)
            return

        messagebox.showinfo("Saved", "Target updated successfully.", parent=self)

        if self.on_save_complete:
            try:
                self.on_save_complete(self.target_id)
            except Exception:
                logger.exception("on_save_complete callback raised")

    def _change_target_id(self) -> None:
        if not messagebox.askyesno("Change Target ID", strings.WARN_CHANGE_TARGET_ID, parent=self):
            return

        new_id = simpledialog.askstring(
            "New Target ID",
            "Enter new target ID:",
            initialvalue=self.target_id,
            parent=self,
        )
        if not new_id or not new_id.strip() or new_id.strip() == self.target_id:
            return

        new_id = new_id.strip()
        try:
            update_target_id(self.conn, self.target_id, new_id, get_current_user())
        except GWError as exc:
            messagebox.showerror(f"Error [GW{exc.code}]", exc.message, parent=self)
            return
        except Exception as exc:
            logger.exception("Unexpected error changing target ID")
            messagebox.showerror("Unexpected Error", str(exc), parent=self)
            return

        self.target_id = new_id
        self.title(f"Target Details — {new_id}")
        self._load_target()  # reload with new ID (no confirm)

    # ------------------------------------------------------------------
    # Merge
    # ------------------------------------------------------------------

    def _open_merge(self) -> None:
        from display.merge_modal import MergeTargetsDialog
        MergeTargetsDialog(
            self,
            self.conn,
            keep_id=self.target_id,
            on_merge_complete=self._on_merged,
        )

    def _on_merged(self, surviving_id: str) -> None:
        if surviving_id != self.target_id:
            self.target_id = surviving_id
            self.title(f"Target Details — {surviving_id}")
        self._load_target()
