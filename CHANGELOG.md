# CHANGELOG

All changes from `aaa_gemini.py` → the refactored 4-file package.

---

## [refactor] Split monolith into 4-file package

`aaa_gemini.py` (1,349 lines, one file) →

```
data_entry_app/
  main.py        ~15 lines   — entry point only
  app.py         ~560 lines  — UI + event logic
  validation.py  ~145 lines  — pure validation helpers (no Tkinter)
  excel_io.py    ~100 lines  — file I/O + prefs (no Tkinter)
  widgets.py     ~110 lines  — ToolTip + ShadowStore classes
```

**Why:** `validation.py` and `excel_io.py` are now fully testable with
`pytest` without launching the GUI. `app.py` is ~60% shorter and reads
like a UI file, not a grab-bag.

---

## [fix] Cross-platform help file opening

**File:** `app.py` → `open_file_cross_platform()`
**Was:** `os.startfile(help_path) if os.name == 'nt' else None`
— silently did nothing on macOS and Linux.
**Now:** Routes to `os.startfile` (Windows), `open` (macOS),
or `xdg-open` (Linux) via `subprocess.Popen`.

---

## [fix] Shadow index desync after row deletion

**File:** `widgets.py` → `ShadowStore.delete_row()`
**Was:** When a row was deleted, the shadow dict keys were never
re-indexed. Any hover or edit on rows *below* the deleted one would
read stale/mismatched raw values.
**Now:** `ShadowStore.delete_row(sheet, row)` atomically removes the
deleted row's entries and decrements all row indices above it in one
pass. Called automatically from `_delete_row_after_flash()`.

---

## [fix] No-op duplicate try/except in add_row_from_inputs

**File:** `app.py` → `add_row_from_inputs()`
**Was:**
```python
try:
    sheet.cell(...).value = parsed
except Exception:
    sheet.cell(...).value = parsed   # identical — error was silently eaten
```
**Now:** Single write; the except block logs the error and shows a
status-bar warning so failures are visible.

---

## [fix] Ctrl+N shortcut mismatch (docs vs code)

**File:** `app.py` → `_bind_shortcuts()`
**Was:** README documented `Ctrl+N` = "Clear input fields" but the
code bound it to `new_file()`.
**Now:**
- `Ctrl+N` → `reset_to_add_mode()` (clear fields, matches README)
- `Ctrl+Shift+N` → `new_file()` (new workbook, aligns with most apps)
Menu labels updated to match.

---

## [fix] Broad "id" substring false positives in column detection

**File:** `validation.py` → `ID_PATTERN`, `infer_validation_rules()`
**Was:** `"id" in h_lower` matched "valid", "liquid", "modified",
"candidate", "invalid", etc., wrongly applying Strict duplicate policy.
**Now:** `ID_PATTERN = re.compile(r"(^|[_\s])id($|[_\s])|^id$|_id$")`
Only matches actual ID columns: `id`, `order_id`, `ref_id`, `ID`, etc.

---

## [fix] Prefs file no longer pollutes the Excel directory

**File:** `excel_io.py` → `get_prefs_path()`
**Was:** `filepath + ".prefs.json"` — created a config file inside
whatever folder held the spreadsheet (shared drives, Dropbox, etc.).
**Now:** Stored in the OS user-config directory:
- Windows: `%APPDATA%\tEppy_DataEntry\`
- macOS: `~/Library/Application Support/tEppy_DataEntry/`
- Linux: `~/.config/tEppy_DataEntry/`
Falls back to `~/.tEppy_DataEntry/` if `platformdirs` is not installed.

---

## [fix] Duplicate tooltip system removed

**File:** `widgets.py` (ToolTip class kept); `app.py` (`_show_tooltip` /
`_hide_tooltip` methods removed)
**Was:** Two parallel, overlapping tooltip implementations existed side
by side — the standalone `ToolTip` class and `_show_tooltip/_hide_tooltip`
instance methods on the app class.
**Now:** Only `ToolTip` (class-based, bind on construction). All usages
updated to use it.

---

## [fix] Removed dead stub _flush_shadow_to_workbook

**File:** `app.py`
**Was:** Method existed with a `pass` body and a comment "safety flush
if needed". Never called, never implemented.
**Now:** Removed entirely. The shadow store and workbook are kept in sync
during Add/Edit operations as before.

---

## [feat] Ctrl+D — Duplicate Row (was documented, never implemented)

**File:** `app.py` → `duplicate_selected_row()`
Copies the selected row's raw values (from ShadowStore, so pre-rounding
originals are preserved) into the input fields in add-mode. The user
can tweak values before pressing Add/Enter to save the new row.
Also added a ⧉ Duplicate toolbar button.

---

## [feat] Ctrl+Shift+I — Insert Blank Row (was documented, never implemented)

**File:** `app.py` → `insert_blank_row()`
Inserts an empty row in the workbook at the currently selected Treeview
position using `sheet.insert_rows()`. Shadow store keys below the
insertion point are shifted down by 1 to stay in sync.
If nothing is selected, clears the input fields for a manual append.

---

## [improve] Bare except clauses replaced with DEBUG logging

**File:** `app.py` → `_log()` helper
**Was:** `except: pass` or `except Exception: pass` in 20+ locations,
silently swallowing every error.
**Now:** All suppressed exceptions are passed to `_log(label, exc)`.
Set the environment variable `APP_DEBUG=1` before launching to print
them to the console:
```bash
APP_DEBUG=1 python main.py
```
Production behaviour (no console spam) is unchanged by default.

---

## [improve] Status bar shows row count and filter context

**File:** `app.py` → `_on_row_select()`, `_apply_filters()`,
`_load_active_sheet()`
- Load: "Loaded 'Sheet1' — 142 rows, 6 columns."
- Select: "Row 12 of 142"
- Filter active: "Filtered: 8 of 142 rows shown"

---

## [improve] all_rows cache kept in sync with deletions

**File:** `app.py` → `_delete_row_after_flash()`
**Was:** `all_rows` (the filter cache) was never updated on deletion,
so re-applying a filter after a delete could show ghost rows.
**Now:** The corresponding entry is popped from `all_rows` on deletion.
