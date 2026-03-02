"""
widgets.py
Reusable UI helpers: ToolTip widget and ShadowStore data structure.

FIX: Removed the duplicate _show_tooltip / _hide_tooltip methods that existed
     alongside the ToolTip class in the original monolith. One tooltip system only.
"""

import tkinter as tk
from tkinter import ttk


# ---------------------------------------------------------------------------
# ToolTip
# ---------------------------------------------------------------------------

class ToolTip:
    """
    Lightweight hover tooltip. Bind to any widget:
        ToolTip(my_button, "Click to save")
    """

    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self._tip_window = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, event=None):
        if self._tip_window or not self.text:
            return
        try:
            x, y, _, _ = self.widget.bbox("insert")
        except Exception:
            x, y = 0, 0
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + 20

        self._tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.attributes("-topmost", True)
        tw.wm_geometry(f"+{x}+{y}")

        label = ttk.Label(
            tw,
            text=self.text,
            background="#ffffe0",
            padding=(6, 3),
            relief="solid",
        )
        label.pack()

    def _hide(self, event=None):
        if self._tip_window:
            self._tip_window.destroy()
            self._tip_window = None

    def update_text(self, new_text: str):
        self.text = new_text


# ---------------------------------------------------------------------------
# ShadowStore
# FIX: Re-indexing after row deletion is now built into the store so callers
#      don't have to remember to do it manually (which was the original bug).
# ---------------------------------------------------------------------------

class ShadowStore:
    """
    Parallel data store that tracks the raw / parsed / rounded state of every
    cell, keyed by (sheet_name, display_row_index, col_index).

    Treeview can only hold strings; this store holds the real data.
    """

    def __init__(self):
        self._data: dict = {}

    # ------------------------------------------------------------------
    # Core get/set
    # ------------------------------------------------------------------

    def set(self, sheet: str, row: int, col: int,
            raw_text, parsed_value, rounded_flag: bool = False):
        key = (sheet, int(row), int(col))
        self._data[key] = {
            "raw": None if raw_text is None else str(raw_text),
            "value": parsed_value,
            "rounded_flag": bool(rounded_flag),
        }

    def get(self, sheet: str, row: int, col: int) -> dict | None:
        return self._data.get((sheet, int(row), int(col)))

    def clear_sheet(self, sheet: str):
        self._data = {k: v for k, v in self._data.items() if k[0] != sheet}

    # ------------------------------------------------------------------
    # FIX: Row deletion re-indexing
    # When a row is deleted, all rows below it must have their index decremented
    # so the shadow keys stay in sync with the Treeview display indices.
    # Previously this was never done, causing hover/edit to read stale data.
    # ------------------------------------------------------------------

    def delete_row(self, sheet: str, deleted_display_row: int):
        """
        Remove the entry for the deleted row and shift all rows below it up by 1.
        Call this immediately after removing the row from the Treeview.
        """
        new_data = {}
        for (sname, ridx, cidx), val in self._data.items():
            if sname != sheet:
                new_data[(sname, ridx, cidx)] = val
            elif ridx == deleted_display_row:
                pass  # drop the deleted row's entries
            elif ridx > deleted_display_row:
                new_data[(sname, ridx - 1, cidx)] = val  # shift up
            else:
                new_data[(sname, ridx, cidx)] = val
        self._data = new_data

    # ------------------------------------------------------------------
    # Convenience
    # ------------------------------------------------------------------

    def is_rounded(self, sheet: str, row: int, col: int) -> bool:
        entry = self.get(sheet, row, col)
        return bool(entry and entry.get("rounded_flag"))

    def raw_value(self, sheet: str, row: int, col: int) -> str | None:
        entry = self.get(sheet, row, col)
        return entry.get("raw") if entry else None
