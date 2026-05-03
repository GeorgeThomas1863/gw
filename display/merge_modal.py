"""
ui/merge_targets.py — Modal dialog for merging two targets.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Optional
import sqlite3

from data.database import get_target
from src.targets import merge_targets
from util.errors import GWError
from util.logger import get_logger

logger = get_logger("gw.ui.merge_targets")


class MergeTargetsDialog(tk.Toplevel):
    """Modal dialog for merging two targets.

    The 'keep' target absorbs the 'absorb' target — all selectors from
    absorb move to keep, then absorb is deleted.
    """

    def __init__(
        self,
        parent: tk.Widget,
        conn: sqlite3.Connection,
        keep_id: str = "",
        on_merge_complete: Optional[Callable[[str], None]] = None,
    ) -> None:
        super().__init__(parent)
        self.conn = conn
        self.on_merge_complete = on_merge_complete

        self.title("Merge Targets")
        self.resizable(False, False)
        self.grab_set()  # modal

        self._build_ui()
        self._center_on_parent(parent)

        # Pre-fill keep_id if provided
        if keep_id:
            self._keep_id_var.set(keep_id)
            self._lookup_target("keep")

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        # ---- Keep Target section ----
        keep_frame = ttk.LabelFrame(self, text="Keep Target", padding=8)
        keep_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 4))
        keep_frame.columnconfigure(1, weight=1)

        ttk.Label(keep_frame, text="ID:").grid(row=0, column=0, sticky="w", **pad)
        self._keep_id_var = tk.StringVar()
        self._keep_id_entry = ttk.Entry(keep_frame, textvariable=self._keep_id_var, width=28)
        self._keep_id_entry.grid(row=0, column=1, sticky="ew", **pad)
        self._keep_id_entry.bind("<FocusOut>", lambda _e: self._lookup_target("keep"))

        ttk.Button(
            keep_frame, text="Lookup", command=lambda: self._lookup_target("keep")
        ).grid(row=0, column=2, padx=(0, 4), pady=6)

        ttk.Label(keep_frame, text="Name:").grid(row=1, column=0, sticky="w", **pad)
        self._keep_name_var = tk.StringVar()
        ttk.Entry(
            keep_frame, textvariable=self._keep_name_var, state="readonly", width=36
        ).grid(row=1, column=1, columnspan=2, sticky="ew", **pad)

        ttk.Label(keep_frame, text="Selectors:").grid(row=2, column=0, sticky="w", **pad)
        self._keep_count_var = tk.StringVar()
        ttk.Entry(
            keep_frame, textvariable=self._keep_count_var, state="readonly", width=10
        ).grid(row=2, column=1, sticky="w", **pad)

        # ---- Swap button ----
        ttk.Button(self, text="⇅ Swap", command=self._swap).grid(
            row=1, column=0, pady=4
        )

        # ---- Absorb Target section ----
        absorb_frame = ttk.LabelFrame(self, text="Absorb Target", padding=8)
        absorb_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=(4, 4))
        absorb_frame.columnconfigure(1, weight=1)

        ttk.Label(absorb_frame, text="ID:").grid(row=0, column=0, sticky="w", **pad)
        self._absorb_id_var = tk.StringVar()
        self._absorb_id_entry = ttk.Entry(
            absorb_frame, textvariable=self._absorb_id_var, width=28
        )
        self._absorb_id_entry.grid(row=0, column=1, sticky="ew", **pad)
        self._absorb_id_entry.bind("<FocusOut>", lambda _e: self._lookup_target("absorb"))

        ttk.Button(
            absorb_frame, text="Lookup", command=lambda: self._lookup_target("absorb")
        ).grid(row=0, column=2, padx=(0, 4), pady=6)

        ttk.Label(absorb_frame, text="Name:").grid(row=1, column=0, sticky="w", **pad)
        self._absorb_name_var = tk.StringVar()
        ttk.Entry(
            absorb_frame, textvariable=self._absorb_name_var, state="readonly", width=36
        ).grid(row=1, column=1, columnspan=2, sticky="ew", **pad)

        ttk.Label(absorb_frame, text="Selectors:").grid(row=2, column=0, sticky="w", **pad)
        self._absorb_count_var = tk.StringVar()
        ttk.Entry(
            absorb_frame, textvariable=self._absorb_count_var, state="readonly", width=10
        ).grid(row=2, column=1, sticky="w", **pad)

        # ---- Buttons ----
        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=3, column=0, pady=(8, 12))

        ttk.Button(btn_frame, text="Merge", command=self._do_merge).grid(
            row=0, column=0, padx=16
        )
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).grid(
            row=0, column=1, padx=16
        )

        self.columnconfigure(0, weight=1)

    # ------------------------------------------------------------------
    # Target lookup
    # ------------------------------------------------------------------

    def _lookup_target(self, side: str) -> None:
        """Look up target by ID and populate name + selector count fields."""
        if side == "keep":
            id_var = self._keep_id_var
            name_var = self._keep_name_var
            count_var = self._keep_count_var
        else:
            id_var = self._absorb_id_var
            name_var = self._absorb_name_var
            count_var = self._absorb_count_var

        target_id = id_var.get().strip()
        if not target_id:
            name_var.set("")
            count_var.set("")
            return

        try:
            target = get_target(self.conn, target_id)
            if target is None:
                name_var.set("[Not found]")
                count_var.set("—")
                return

            name_var.set(target.target_name or "")
            row = self.conn.execute(
                "SELECT COUNT(*) FROM selectors WHERE target_id = ?", (target_id,)
            ).fetchone()
            count_var.set(str(row[0]) if row else "0")
        except Exception as exc:
            logger.warning("_lookup_target(%s): %s", target_id, exc)
            name_var.set("[Error]")
            count_var.set("—")

    # ------------------------------------------------------------------
    # Swap
    # ------------------------------------------------------------------

    def _swap(self) -> None:
        keep_id = self._keep_id_var.get()
        keep_name = self._keep_name_var.get()
        keep_count = self._keep_count_var.get()

        self._keep_id_var.set(self._absorb_id_var.get())
        self._keep_name_var.set(self._absorb_name_var.get())
        self._keep_count_var.set(self._absorb_count_var.get())

        self._absorb_id_var.set(keep_id)
        self._absorb_name_var.set(keep_name)
        self._absorb_count_var.set(keep_count)

    # ------------------------------------------------------------------
    # Merge action
    # ------------------------------------------------------------------

    def _do_merge(self) -> None:
        keep_id = self._keep_id_var.get().strip()
        absorb_id = self._absorb_id_var.get().strip()

        if not keep_id:
            messagebox.showwarning("Missing ID", "Please enter a Keep Target ID.", parent=self)
            return
        if not absorb_id:
            messagebox.showwarning(
                "Missing ID", "Please enter an Absorb Target ID.", parent=self
            )
            return

        try:
            merge_targets(keep_id, absorb_id, self.conn)
        except GWError as e:
            messagebox.showerror(
                f"Error [GW{e.code}]", e.message, parent=self
            )
            return
        except Exception as exc:
            logger.exception("Unexpected error during merge")
            messagebox.showerror("Unexpected Error", str(exc), parent=self)
            return

        messagebox.showinfo(
            "Merge Complete",
            f"Target {absorb_id!r} has been merged into {keep_id!r}.",
            parent=self,
        )

        if self.on_merge_complete is not None:
            try:
                self.on_merge_complete(keep_id)
            except Exception:
                logger.exception("on_merge_complete callback raised")

        self.destroy()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _center_on_parent(self, parent: tk.Widget) -> None:
        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width() // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        self.geometry(f"+{pw - w // 2}+{ph - h // 2}")
