# 🧮 Dynamic Excel Data Entry App (Tkinter + ttkbootstrap)

A modern **Excel-style data entry companion** built with **Tkinter** and **ttkbootstrap**, designed for users who prefer the familiarity of spreadsheets — with the reliability, structure, and validation of a professional tool.

> ✨ Ideal for office clerks, researchers, and small data teams handling recurring Excel templates.

---

## 🚀 Key Features

### 🗂 Excel Integration
- Opens and edits `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, and `.ods` files  
- Automatically converts non-`.xlsx` formats for safe saving  
- Preserves sheet headers and data structure intelligently  
- Supports multiple sheets with a quick selector  

### 💡 Adaptive Validation
- Auto-detects column types via headers (`Date`, `Email`, `Amount`, etc.)
- Built-in numeric formatting:
  - `amount`, `price`, `total`, `rate`, `cost`, `balance` → **2 decimal rounding**
  - `qty`, `age`, `count`, `number` → **integer rounding**
- Inline validation feedback (color-coded + per-field message)
- Header-based **duplicate control**:
  - Choose between **None / Warn / Strict**
- “Required” toggle for each column — simple and intuitive
- Smart duplicate detection excludes original value during edits

### ⚙️ Excel-Like Workflow
- Fast keyboard shortcuts:
  | Action | Shortcut |
  |--------|-----------|
  | Open file | Ctrl+O |
  | Save / Save As | Ctrl+S / Ctrl+Shift+S |
  | Clear input fields | Ctrl+N |
  | Add row | Enter (on last field) |
  | Edit selected row | F2 |
  | Duplicate row | Ctrl+D |
  | Delete row | Ctrl+Shift+D |
  | Insert blank row | Ctrl+Shift+I |
  | Cancel edit | Esc |
  | Quit | Ctrl+Q |
- Auto-increment IDs (`ID001 → ID002`, `Ref10 → Ref11`)
- Visual feedback (green flash = added, red = deleted)
- Optional **Auto-Save** for Add/Edit/Delete/Insert/Duplicate actions

### 🎨 Polished Interface
- Clean, responsive UI powered by **ttkbootstrap**
- Dynamic theme selector (Cosmo, Darkly, Flatly, etc.)
- Color-coded status bar with fade-out transitions
- Built-in tooltip and hover effects
- Scrollable input area for wide spreadsheets
- Inline help system via `help.txt` (❓ Help button in toolbar)

---

## 🧠 How to Use

1. **Prepare your Excel template**
   - First row should contain headers like `Name`, `Email`, `Amount`, `Date`.
   - Headers determine field validation automatically.
2. **Launch the app**
   ```bash
   python test.py
   ```
3. **Open your Excel file**
   - Answer **Yes** on startup prompt or use **Ctrl+O**.
4. **Enter or edit data**
   - Use **Tab** / **Enter** to move between fields.
   - **Enter** on last field → adds the row.
   - Double-click a row to edit it, press **Enter** to save.
5. **Save**
   - Use **Ctrl+S** or enable **Auto-Save** in the toolbar.
6. **Need help?**
   - Click the ❓ **Help** button or check `help.txt`.

---

## ✅ Best Practices

| Task | Recommendation |
|------|----------------|
| **Headers** | Use descriptive names like “Amount”, “Email”, “Date of Birth”. |
| **Numeric fields** | “Amount” and “Price” → decimal; “Qty” and “Age” → integer. |
| **Duplicate control** | Set **Strict** for IDs; **Warn** for fields like name or email. |
| **Required fields** | Mark essential data columns as Required. |
| **Auto-Save** | Enable it for safer workflows. |
| **Backup** | Keep a copy of your Excel before batch edits. |
| **Themes** | Experiment with ttkbootstrap themes for better visibility. |

---

## 🧩 Requirements

```bash
pip install ttkbootstrap openpyxl pandas pyxlsb odfpy xlrd
```

Requires **Python 3.10+**  
On Linux, ensure Tkinter is installed:
```bash
sudo apt-get install python3.11-tk
```

---

## 🧾 Credits

- **UI/Theme** → ttkbootstrap by [israel-dryer](https://github.com/israel-dryer)
- **Excel Engine** → openpyxl + pandas
- **Design & Code** → Built with ❤️ by [tEppy](https://github.com/tepong32) (2025)

---

## 📜 License

Personal and educational use allowed.  
Credit the author for commercial use or redistribution.

---

## 🧭 Version History

See [`CHANGELOG.md`](./changelog.md) for full release notes.
