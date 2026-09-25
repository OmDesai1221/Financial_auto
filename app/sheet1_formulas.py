"""Lets the transformed workbooks keep the arithmetic a Sheet1 cell was typed
as (e.g. Total Balance "=600000/12*5"), so the person can still see how a
number was worked out in Excel's formula bar.

Sheet1 is normally read with cached values only. This reads the same sheet
WITHOUT caching to find Total Balance cells that hold a plain-arithmetic
formula. Formulas that point at other cells (subtotals like =SUM(L8:L25))
are ignored — those would mean something different outside Sheet1."""
import re

CELL_REF_RE = re.compile(r"\$?[A-Z]{1,3}\$?\d{1,4}")
GROSS_PREFIX = "gross sale"


def _pure_arithmetic(v):
    return (isinstance(v, str) and v.startswith("=") and len(v) > 1
            and not CELL_REF_RE.search(v) and "!" not in v)


def read_total_formulas(ws_formulas):
    """Returns (code -> formula text, gross_sales_formula_or_None)."""
    by_code, gross = {}, None
    for row in ws_formulas.iter_rows(min_row=7, max_row=ws_formulas.max_row):
        code, desc, total = row[1].value, row[2].value, row[11].value
        if not _pure_arithmetic(total):
            continue
        if code:
            by_code[code] = total
        elif isinstance(desc, str) and desc.strip().lower().startswith(GROSS_PREFIX):
            gross = total
    return by_code, gross


def compose(found, formulas, negate=False):
    """found: [(code, numeric_value)] for the codes that exist in Sheet1.
    Returns a formula string if any of them was typed as a formula in Sheet1
    (other codes are folded in as plain numbers), else None."""
    if not formulas or not any(c in formulas for c, _ in found):
        return None
    parts = []
    for code, val in found:
        f = formulas.get(code)
        parts.append(f[1:] if f else repr(val if val is not None else 0))
    body = parts[0] if len(parts) == 1 else "+".join(f"({p})" for p in parts)
    if negate:
        body = f"-({body})"
    return "=" + body
