"""
ui/results_window.py — Results display window with GW and S tabs.
"""
from __future__ import annotations

import datetime
import sqlite3
import tkinter as tk
import webbrowser
from tkinter import ttk
from typing import Optional

from utils.logger import get_logger

logger = get_logger(__name__)


class ResultsWindow(tk.Toplevel):
    """Displays GrayWolfe and S search results in separate tabs with filters."""

    def __init__(
        self,
        parent: tk.Widget,
        gw_results: list[dict],
        s_results: list[dict],
        query_terms: list[str],
        conn: sqlite3.Connection,
        s_client=None,
    ) -> None:
        super().__init__(parent)
        self.gw_results = gw_results
        self.s_results = s_results
        self.query_terms = query_terms
        self.conn = conn
        self.s_client = s_client

        self.title("GrayWolfe — Search Results")
        self.minsize(800, 520)
        self.resizable(True, True)

        self._build_ui()
        self._populate_gw()
        self._populate_s()

        self.transient(parent)
        self.focus_set()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=4, pady=4)

        self._build_gw_tab()
        self._build_s_tab()

    # ---- GW Tab ----

    def _build_gw_tab(self) -> None:
        gw_frame = ttk.Frame(self.notebook)
        self.notebook.add(gw_frame, text="GrayWolfe")

        # Stats bar
        stats_frame = ttk.Frame(gw_frame)
        stats_frame.pack(fill="x", padx=6, pady=(6, 2))
        self._gw_searched_lbl = ttk.Label(stats_frame, text="Searched: 0")
        self._gw_searched_lbl.pack(side="left", padx=6)
        self._gw_hits_lbl = ttk.Label(stats_frame, text="GW Hits: 0")
        self._gw_hits_lbl.pack(side="left", padx=6)
        self._gw_targets_lbl = ttk.Label(stats_frame, text="Targets: 0")
        self._gw_targets_lbl.pack(side="left", padx=6)
        self._gw_in_s_lbl = ttk.Label(stats_frame, text="In S: 0")
        self._gw_in_s_lbl.pack(side="left", padx=6)
        self._gw_not_in_s_lbl = ttk.Label(stats_frame, text="Not In S: 0")
        self._gw_not_in_s_lbl.pack(side="left", padx=6)

        # Filter bar
        filter_frame = ttk.Frame(gw_frame)
        filter_frame.pack(fill="x", padx=6, pady=2)

        ttk.Label(filter_frame, text="Filter by Target:").pack(side="left")
        self._gw_target_filter_var = tk.StringVar(value="All Targets")
        self._gw_target_combo = ttk.Combobox(
            filter_frame, textvariable=self._gw_target_filter_var,
            state="readonly", width=24
        )
        self._gw_target_combo.pack(side="left", padx=4)
        self._gw_target_combo.bind("<<ComboboxSelected>>", lambda _: self._apply_gw_filter())

        self._gw_in_s_btn = ttk.Button(filter_frame, text="In S",
                                        command=lambda: self._toggle_gw_filter("in_s"))
        self._gw_in_s_btn.pack(side="left", padx=2)
        self._gw_not_in_s_btn = ttk.Button(filter_frame, text="Not In S",
                                             command=lambda: self._toggle_gw_filter("not_in_s"))
        self._gw_not_in_s_btn.pack(side="left", padx=2)
        self._gw_in_gw_btn = ttk.Button(filter_frame, text="In GW",
                                         command=lambda: self._toggle_gw_filter("in_gw"))
        self._gw_in_gw_btn.pack(side="left", padx=2)
        self._gw_not_in_gw_btn = ttk.Button(filter_frame, text="Not In GW",
                                              command=lambda: self._toggle_gw_filter("not_in_gw"))
        self._gw_not_in_gw_btn.pack(side="left", padx=2)
        ttk.Button(filter_frame, text="Clear Filters",
                   command=self._clear_gw_filters).pack(side="left", padx=4)

        self._gw_active_filter: Optional[str] = None

        # Treeview
        tree_frame = ttk.Frame(gw_frame)
        tree_frame.pack(fill="both", expand=True, padx=6, pady=(2, 6))

        cols = ("selector", "type", "target", "in_gw", "s_hits")
        self._gw_tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        self._gw_tree.heading("selector", text="Selector")
        self._gw_tree.heading("type", text="Type")
        self._gw_tree.heading("target", text="Target")
        self._gw_tree.heading("in_gw", text="In GW")
        self._gw_tree.heading("s_hits", text="S Hits")
        self._gw_tree.column("selector", width=240, anchor="w")
        self._gw_tree.column("type", width=90, anchor="w")
        self._gw_tree.column("target", width=180, anchor="w")
        self._gw_tree.column("in_gw", width=60, anchor="center")
        self._gw_tree.column("s_hits", width=60, anchor="center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._gw_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._gw_tree.xview)
        self._gw_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._gw_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self._gw_tree.bind("<Double-1>", self._gw_double_click)

    # ---- S Tab ----

    def _build_s_tab(self) -> None:
        s_frame = ttk.Frame(self.notebook)
        self.notebook.add(s_frame, text="S")

        # Stats bar
        stats_frame = ttk.Frame(s_frame)
        stats_frame.pack(fill="x", padx=6, pady=(6, 2))
        self._s_hits_lbl = ttk.Label(stats_frame, text="S Hits: 0")
        self._s_hits_lbl.pack(side="left", padx=6)
        self._s_searched_lbl = ttk.Label(stats_frame, text="Searched: 0")
        self._s_searched_lbl.pack(side="left", padx=6)

        # Filter bar
        filter_frame = ttk.Frame(s_frame)
        filter_frame.pack(fill="x", padx=6, pady=2)

        self._s_fd302_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(filter_frame, text="FD-302 Only",
                        variable=self._s_fd302_var,
                        command=self._apply_s_filter).pack(side="left", padx=4)

        self._s_last_year_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(filter_frame, text="Last Year",
                        variable=self._s_last_year_var,
                        command=self._apply_s_filter).pack(side="left", padx=4)

        self._s_no_attach_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(filter_frame, text="No Attachments",
                        variable=self._s_no_attach_var,
                        command=self._apply_s_filter).pack(side="left", padx=4)

        ttk.Label(filter_frame, text="By Selector:").pack(side="left", padx=(8, 2))
        self._s_selector_filter_var = tk.StringVar(value="All")
        self._s_selector_combo = ttk.Combobox(
            filter_frame, textvariable=self._s_selector_filter_var,
            state="readonly", width=22
        )
        self._s_selector_combo.pack(side="left", padx=2)
        self._s_selector_combo.bind("<<ComboboxSelected>>", lambda _: self._apply_s_filter())

        # Treeview
        tree_frame = ttk.Frame(s_frame)
        tree_frame.pack(fill="both", expand=True, padx=6, pady=(2, 6))

        cols = ("selector", "doc_type", "case", "title", "author", "date")
        self._s_tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        self._s_tree.heading("selector", text="Selector")
        self._s_tree.heading("doc_type", text="Doc Type")
        self._s_tree.heading("case", text="Case")
        self._s_tree.heading("title", text="Title")
        self._s_tree.heading("author", text="Author")
        self._s_tree.heading("date", text="Date")
        self._s_tree.column("selector", width=160, anchor="w")
        self._s_tree.column("doc_type", width=80, anchor="w")
        self._s_tree.column("case", width=130, anchor="w")
        self._s_tree.column("title", width=240, anchor="w")
        self._s_tree.column("author", width=90, anchor="w")
        self._s_tree.column("date", width=90, anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._s_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._s_tree.xview)
        self._s_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._s_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self._s_tree.bind("<Double-1>", self._s_double_click)

        # Store full row data keyed by tree item id
        self._s_item_data: dict[str, dict] = {}

    # ------------------------------------------------------------------
    # Populate
    # ------------------------------------------------------------------

    def _populate_gw(self) -> None:
        # Build s_hit_count lookup from s_results
        s_hits: dict[str, int] = {}
        for r in self.s_results:
            sel = r.get("selector", "")
            s_hits[sel] = s_hits.get(sel, 0) + 1

        self._gw_all_rows: list[dict] = []
        targets_seen: set[str] = set()
        target_names: list[str] = ["All Targets"]

        for r in self.gw_results:
            selector = r.get("selector", r.get("query_value", ""))
            hits = s_hits.get(selector, 0)
            in_s = hits > 0
            row = {
                "selector": selector,
                "type": r.get("selector_type") or "",
                "target": r.get("target_name") or "",
                "target_id": r.get("target_id"),
                "in_gw": "Yes" if r.get("in_gray_wolfe") else "No",
                "in_s": in_s,
                "s_hits": str(hits) if hits else "",
            }
            self._gw_all_rows.append(row)

            tid = r.get("target_id")
            tname = r.get("target_name")
            if tid and tid not in targets_seen:
                targets_seen.add(tid)
                if tname:
                    target_names.append(tname)

        self._gw_target_combo["values"] = target_names
        self._gw_target_filter_var.set("All Targets")

        self._apply_gw_filter()

        # Stats
        total = len(self.query_terms)
        in_gw = sum(1 for r in self._gw_all_rows if r["in_gw"] == "Yes")
        in_s_count = sum(1 for r in self._gw_all_rows if r["in_s"])

        self._gw_searched_lbl.config(text=f"Searched: {total}")
        self._gw_hits_lbl.config(text=f"GW Hits: {in_gw}")
        self._gw_targets_lbl.config(text=f"Targets: {len(targets_seen)}")
        self._gw_in_s_lbl.config(text=f"In S: {in_s_count}")
        self._gw_not_in_s_lbl.config(text=f"Not In S: {total - in_s_count}")

    def _populate_s(self) -> None:
        self._s_all_rows = self.s_results[:]

        selectors = sorted({r.get("selector", "") for r in self._s_all_rows} - {""})
        self._s_selector_combo["values"] = ["All"] + selectors
        self._s_selector_filter_var.set("All")

        self._apply_s_filter()
        self._s_hits_lbl.config(text=f"S Hits: {len(self._s_all_rows)}")
        self._s_searched_lbl.config(text=f"Searched: {len(self.query_terms)}")

    # ------------------------------------------------------------------
    # GW filtering
    # ------------------------------------------------------------------

    def _toggle_gw_filter(self, name: str) -> None:
        self._gw_active_filter = None if self._gw_active_filter == name else name
        self._apply_gw_filter()

    def _clear_gw_filters(self) -> None:
        self._gw_active_filter = None
        self._gw_target_filter_var.set("All Targets")
        self._apply_gw_filter()

    def _apply_gw_filter(self) -> None:
        target_filter = self._gw_target_filter_var.get()

        self._gw_tree.delete(*self._gw_tree.get_children())
        for row in self._gw_all_rows:
            if target_filter != "All Targets" and row["target"] != target_filter:
                continue
            if self._gw_active_filter == "in_s" and not row["in_s"]:
                continue
            if self._gw_active_filter == "not_in_s" and row["in_s"]:
                continue
            if self._gw_active_filter == "in_gw" and row["in_gw"] != "Yes":
                continue
            if self._gw_active_filter == "not_in_gw" and row["in_gw"] == "Yes":
                continue

            iid = self._gw_tree.insert(
                "", "end",
                values=(row["selector"], row["type"], row["target"],
                        row["in_gw"], row["s_hits"]),
                tags=(row.get("target_id") or "",),
            )

    # ------------------------------------------------------------------
    # S filtering
    # ------------------------------------------------------------------

    def _apply_s_filter(self) -> None:
        fd302_only = self._s_fd302_var.get()
        last_year = self._s_last_year_var.get()
        no_attach = self._s_no_attach_var.get()
        sel_filter = self._s_selector_filter_var.get()

        cutoff = (datetime.datetime.now() - datetime.timedelta(days=365)).isoformat()

        self._s_tree.delete(*self._s_tree.get_children())
        self._s_item_data.clear()

        for row in self._s_all_rows:
            if fd302_only and row.get("doc_type", "").upper() != "FD302":
                continue
            if last_year:
                date_str = row.get("created_date", "") or ""
                if date_str < cutoff:
                    continue
            if no_attach and "attachment" in (row.get("doc_sub_type") or "").lower():
                continue
            if sel_filter != "All" and row.get("selector", "") != sel_filter:
                continue

            iid = self._s_tree.insert(
                "", "end",
                values=(
                    row.get("selector", ""),
                    row.get("doc_type", ""),
                    row.get("case", ""),
                    row.get("doc_title", ""),
                    row.get("author", ""),
                    (row.get("created_date") or "")[:10],
                ),
            )
            self._s_item_data[iid] = row

    # ------------------------------------------------------------------
    # Double-click handlers
    # ------------------------------------------------------------------

    def _gw_double_click(self, event: tk.Event) -> None:
        item = self._gw_tree.identify_row(event.y)
        if not item:
            return
        tags = self._gw_tree.item(item, "tags")
        target_id = tags[0] if tags else None
        if not target_id:
            return
        from ui.target_details import TargetDetailsWindow
        TargetDetailsWindow(self, self.conn, target_id)

    def _s_double_click(self, event: tk.Event) -> None:
        item = self._s_tree.identify_row(event.y)
        if not item:
            return
        row = self._s_item_data.get(item)
        if not row:
            return
        link = row.get("link", "")
        if link:
            webbrowser.open(link)
