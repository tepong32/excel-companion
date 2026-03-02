"""
validation.py
Stateless validation, parsing, and formatting helpers.
No Tkinter imports — fully testable without a GUI.
"""

import re
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP

# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# Matches "id", "ref_id", "_id", "order_id" — but NOT "valid", "liquid", etc.
# FIX: was `"id" in h_lower` which falsely matched unrelated words.
ID_PATTERN = re.compile(r"(^|[_\s])id($|[_\s])|^id$|_id$", re.IGNORECASE)

# Keywords that drive auto-detected column types
DECIMAL_KEYWORDS = ("amount", "price", "rate", "total", "cost", "balance", "value")
INTEGER_KEYWORDS = ("qty", "quantity", "number", "count", "age")

# ---------------------------------------------------------------------------
# Date parsing
# ---------------------------------------------------------------------------
_DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d"]

def try_parse_date(value: str):
    """Return a date object or None. Tries multiple common formats."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(s).date()
    except ValueError:
        return None

# ---------------------------------------------------------------------------
# Numeric helpers
# ---------------------------------------------------------------------------

def is_numeric(value: str) -> bool:
    if value is None:
        return False
    s = str(value).strip()
    if not s:
        return False
    try:
        float(s.replace(",", ""))
        return True
    except ValueError:
        return False


def normalize_numeric(value: str, fmt: str = "integer"):
    """
    Parse and normalise numeric input.
    Returns int for integer columns, float for decimal columns.
    Raises ValueError on bad input.
    """
    if value is None or str(value).strip() == "":
        return None
    s = str(value).replace(",", "").strip()
    try:
        num = float(s)
    except ValueError:
        raise ValueError(f"Invalid numeric input: {value!r}")

    if fmt == "decimal":
        return num
    # Integer format: cast only if lossless
    if num.is_integer():
        return int(num)
    return num


def round_half_up(value, decimal_places: int = 2):
    """
    Excel-style rounding: 0-4 rounds down, 5-9 rounds up.
    Uses Decimal to avoid float precision issues.
    """
    try:
        q = Decimal("1." + "0" * decimal_places)
        return float(Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP))
    except Exception:
        return value


def has_excess_precision(raw_text: str, decimal_limit: int = 2) -> bool:
    """True if the value has more decimal digits than decimal_limit."""
    if not raw_text:
        return False
    s = str(raw_text).replace(",", "").strip()
    if "." in s:
        return len(s.split(".")[1]) > decimal_limit
    return False


def detect_precision_mismatch(raw_text: str, parsed_value, decimal_places: int = 2) -> bool:
    """
    True if raw_text differs from its rounded form, indicating the stored
    value will be silently truncated.
    """
    if not raw_text:
        return False
    s = str(raw_text).strip()
    if not s:
        return False
    try:
        raw_num = float(s.replace(",", ""))
    except ValueError:
        return False
    rounded = round_half_up(raw_num, decimal_places)
    return abs(raw_num - rounded) > 1e-9

# ---------------------------------------------------------------------------
# Display formatting
# ---------------------------------------------------------------------------

def format_value_for_display(value, rule=None, decimal_places: int = 2) -> str:
    """Convert a stored value to a human-readable string for the Treeview."""
    if value is None:
        return ""
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        if rule and rule.get("format") == "decimal":
            return f"{value:.{decimal_places}f}"
        if value.is_integer():
            return str(int(value))
        return str(value)
    if isinstance(value, int):
        return str(value)
    return str(value)

# ---------------------------------------------------------------------------
# Rule inference
# ---------------------------------------------------------------------------

def infer_validation_rules(headers: list, tk_module) -> list:
    """
    Build a list of validation rule dicts from column headers.
    tk_module is passed in so this file stays GUI-framework-agnostic
    (pass `tkinter` from the caller).

    FIX: ID detection now uses ID_PATTERN regex instead of naive substring match
         to avoid false positives like "valid", "liquid", "modified".
    """
    rules = []
    for h in headers:
        h_lower = h.lower() if h else ""
        val_type = "text"
        num_format = None
        is_required_default = True
        duplicate_policy_default = "none"

        if any(k in h_lower for k in DECIMAL_KEYWORDS):
            val_type = "numeric"
            num_format = "decimal"
        elif any(k in h_lower for k in INTEGER_KEYWORDS):
            val_type = "numeric"
            num_format = "integer"
        elif "date" in h_lower:
            val_type = "date"
        elif "email" in h_lower:
            val_type = "email"

        if "optional" in h_lower:
            is_required_default = False

        # FIX: use regex instead of plain `"id" in h_lower`
        if ID_PATTERN.search(h_lower) or "code" in h_lower:
            duplicate_policy_default = "strict"
            is_required_default = False

        rule = {
            "name": h,
            "type": val_type,
            "format": num_format,
            "required_var": tk_module.BooleanVar(value=is_required_default),
            "duplicate_var": tk_module.StringVar(value=duplicate_policy_default),
            "required": is_required_default,
            "duplicate_policy": duplicate_policy_default,
        }
        rules.append(rule)
    return rules
