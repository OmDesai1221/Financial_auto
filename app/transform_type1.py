"""
Type 1 transformation (PNL): merge a new month's GL trial-balance sheet
("Sheet1") into the running P&L workbook.

Rules:
- GL Code is the permanent match key between Sheet1 and P&L. A P&L row's
  GL Code cell may list several codes separated by commas
  (e.g. "X4002007001,X4002007002,X4002007007") — that row's value is then
  the sum of every listed code's Sheet1 balance, and none of those codes
  are ever inserted again as separate rows.
- The new month becomes the first month column; every existing month
  column shifts one column to the right. The first month column and the
  per-unit columns are detected from the workbook itself, so the client
  can add/remove unit columns without breaking anything.
- A Sheet1 code that isn't in P&L yet is first matched by DESCRIPTION
  against P&L rows that have no GL code (manual rows the client already
  keeps, e.g. "AUDIT FEES", "O/S - RAW MATERIAL", "DEPRECIATION"). If one
  matches, the code is attached to that existing row (its history stays
  intact) instead of inserting a duplicate row.
- Only a code with no such row is inserted as a new row: placed next to
  its neighbours in Sheet1 (staying inside its own section, and inside
  that section's SUM range), with 0 for every earlier month.
- Rows whose codes are all absent from Sheet1, and rows with no code at
  all (subtotals + manual accountant entries), are never zeroed or
  guessed: they are only auto-filled when a safe repeating pattern exists
  (subtotal formula, fiscal-YTD accrual, flat value); everything else is
  left blank with the previous value offered as a suggestion.
"""
import re
import copy
import datetime
from collections import namedtuple
from openpyxl.utils import get_column_letter, column_index_from_string

MONTHS = {m.lower(): i for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July",
     "August", "September", "October", "November", "December"], start=1)}

FISCAL_YEAR_START_MONTH = 4  # April

CELL_REF_RE = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d{1,4})')
RANGE_RE = re.compile(r'(\$?[A-Z]{1,3}\$?)(\d{1,4}):(\$?[A-Z]{1,3}\$?)(\d{1,4})')
TRAILING_MULT_RE = re.compile(r'([*])(\d{1,2})(\s*)$')

GL_COL = 2          # column B
DESC_COL = 3        # column C
FIRST_UNIT_COL = 4  # column D
HEADER_ROW = 6
DATE_ROW = 7
FIRST_DATA_ROW = 8

# Reconfigured per workbook by detect_layout(): the first month column and
# the unit columns before it.
TOTAL_COL = 12
UNIT_COLS = list(range(FIRST_UNIT_COL, 12))

GROSS_SALES_LABEL_PREFIX = "gross sale"  # "GROSS SALE" (P&L) / "Gross sales" (Sheet1)

SUBTOTAL_DESCRIPTIONS = {
    "NET SALES", "OTHER INCOME", "CLOSING STOCK", "TOTAL INCOME",
    "TOTAL EXPENSES", "PROFIT & LOSS BEFORE DEPRECIATION", "PROFIT & LOSS",
}


class NotANewMonthError(Exception):
    """Uploaded Sheet1's month is not strictly after the newest month."""


class NotTheLatestMonthError(Exception):
    """A 'replace latest month' upload doesn't match the newest month."""


def _norm(v):
    return " ".join(str(v or "").split()).casefold()


def parse_codes(cell_value):
    if not cell_value:
        return []
    return [c.strip() for c in str(cell_value).split(",") if c.strip()]


def parse_month_from_title(title):
    """'Profit & Loss Account as on July-2026' / 'as on July 2026' -> date(2026,7,1)"""
    m = re.search(r'as on\s+([A-Za-z]+)[-\s]+(\d{4})', title or "")
    if not m:
        raise ValueError(f"Could not parse month/year from title: {title!r}")
    month_name, year = m.group(1), int(m.group(2))
    month_num = MONTHS.get(month_name.lower())
    if not month_num:
        raise ValueError(f"Unrecognized month name: {month_name!r}")
    return datetime.date(year, month_num, 1)


def fiscal_month_number(d):
    """1 for April, 2 for May, ... 12 for March."""
    return (d.month - FISCAL_YEAR_START_MONTH) % 12 + 1


def _prev_month(d):
    if d.month == 1:
        return datetime.date(d.year - 1, 12, 1)
    return datetime.date(d.year, d.month - 1, 1)


def detect_layout(ws):
    """Find the first month column (first date in row 7) and the unit
    columns before it, from the workbook itself."""
    global TOTAL_COL, UNIT_COLS
    found = None
    for c in range(FIRST_UNIT_COL, max(ws.max_column, 12) + 1):
        if isinstance(ws.cell(row=DATE_ROW, column=c).value, datetime.datetime):
            found = c
            break
    TOTAL_COL = found or 12
    UNIT_COLS = list(range(FIRST_UNIT_COL, TOTAL_COL))
    return TOTAL_COL


Sheet1Data = namedtuple("Sheet1Data", "month_date rows gross_sales unit_headers sequence")


def read_sheet1(ws):
    """rows: ordered dicts(code, desc, units[8], total) for coded rows.
    gross_sales: Total Balance typed on Sheet1's own "Gross sale(s)" row
    (no GL code), or None. sequence: ordered ("code", code)/("label", text)
    items — labels (TOTAL INCOME, OPENING STOCK ...) mark section
    boundaries used when placing a brand-new row."""
    title = ws.cell(row=4, column=3).value or ws.cell(row=4, column=2).value
    month_date = parse_month_from_title(title)
    unit_headers = [_norm(ws.cell(row=HEADER_ROW, column=c).value) for c in range(4, 12)]
    rows, sequence = [], []
    gross_sales = None
    for row in ws.iter_rows(min_row=7, max_row=ws.max_row):
        code = row[GL_COL - 1].value
        desc = row[DESC_COL - 1].value
        if code:
            units = [row[c - 1].value for c in range(4, 12)]
            total = row[11].value  # column L: Total Balance
            rows.append({"code": code, "desc": desc, "units": units, "total": total})
            sequence.append(("code", code))
        elif isinstance(desc, str) and desc.strip():
            if desc.strip().lower().startswith(GROSS_SALES_LABEL_PREFIX):
                gross_sales = row[11].value
            else:
                sequence.append(("label", desc))
    return Sheet1Data(month_date, rows, gross_sales, unit_headers, sequence)


def build_unit_map(ws, sheet1_headers):
    """[(pnl_column, sheet1_unit_index)] matched by header text."""
    out = []
    for c in UNIT_COLS:
        h = _norm(ws.cell(row=HEADER_ROW, column=c).value)
        if h and h in sheet1_headers:
            out.append((c, sheet1_headers.index(h)))
    return out


def _sum_units(rows):
    total = [0] * 8
    for r in rows:
        for i, v in enumerate(r["units"]):
            if isinstance(v, (int, float)):
                total[i] += v
    return total


# ---------------------------------------------------------------- formulas

def _shift_col_letter(letter, shift_from_col, delta=1):
    idx = column_index_from_string(letter)
    if idx >= shift_from_col:
        return get_column_letter(idx + delta)
    return letter


def remap_formula_columns(formula, shift_from_col, delta=1):
    def _sub(m):
        dollar_col, col, dollar_row, row_s = m.groups()
        return f"{dollar_col}{_shift_col_letter(col, shift_from_col, delta)}{dollar_row}{row_s}"
    return CELL_REF_RE.sub(_sub, formula)


def remap_formula_rows(formula, remap_fn, end_extra_fn=None):
    """Rewrite row numbers after rows were inserted. A range's END row also
    grows by end_extra_fn(end_row) so rows attached directly below a
    section's last row stay inside that section's SUM."""
    stash = {}

    def range_sub(m):
        c1, r1, c2, r2 = m.groups()
        new_r2 = remap_fn(int(r2)) + (end_extra_fn(int(r2)) if end_extra_fn else 0)
        token = f"\x00{len(stash)}\x00"
        stash[token] = f"{c1}{remap_fn(int(r1))}:{c2}{new_r2}"
        return token

    def single_sub(m):
        dollar_col, col, dollar_row, row_s = m.groups()
        return f"{dollar_col}{col}{dollar_row}{remap_fn(int(row_s))}"

    out = CELL_REF_RE.sub(single_sub, RANGE_RE.sub(range_sub, formula))
    for token, text in stash.items():
        out = out.replace(token, text)
    return out


def _retarget_column(formula, from_col, to_col):
    """Replace only references to from_col with to_col's letter."""
    from_letter = get_column_letter(from_col)
    to_letter = get_column_letter(to_col)

    def _sub(m):
        dollar_col, col, dollar_row, row_s = m.groups()
        if col == from_letter:
            col = to_letter
        return f"{dollar_col}{col}{dollar_row}{row_s}"
    return CELL_REF_RE.sub(_sub, formula)


# ------------------------------------------------------------------ layout

def get_existing_gl_rows(ws):
    """Ordered (row, code_cell_value, desc) for every row carrying a GL code."""
    out = []
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        code = ws.cell(row=r, column=GL_COL).value
        if code:
            out.append((r, code, ws.cell(row=r, column=DESC_COL).value))
    return out


def code_to_row_map(existing_rows):
    m = {}
    for r, code_cell, _ in existing_rows:
        for c in parse_codes(code_cell):
            m.setdefault(c, r)
    return m


def get_month_data_columns(ws):
    """Columns holding a month's data, per the date header in row 7."""
    cols = []
    c = TOTAL_COL
    while ws.cell(row=DATE_ROW, column=c).value is not None:
        cols.append(c)
        c += 1
    return cols


def get_current_pnl_month(ws):
    detect_layout(ws)
    d = ws.cell(row=DATE_ROW, column=TOTAL_COL).value
    if isinstance(d, datetime.datetime):
        return d.date()
    return None


def find_adoptions(ws, unmapped_rows):
    """code -> P&L row: a Sheet1 code not in P&L whose description equals
    exactly one code-less manual P&L row is attached to that row."""
    candidates = {}
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        if ws.cell(row=r, column=GL_COL).value:
            continue
        d = ws.cell(row=r, column=DESC_COL).value
        if not d or not str(d).strip():
            continue
        if str(d).strip().upper() in SUBTOTAL_DESCRIPTIONS:
            continue
        if _norm(d).startswith(GROSS_SALES_LABEL_PREFIX):
            continue
        candidates.setdefault(_norm(d), []).append(r)
    adoptions, used = {}, set()
    for s in unmapped_rows:
        rows = [x for x in candidates.get(_norm(s["desc"]), []) if x not in used]
        if len(rows) == 1:
            adoptions[s["code"]] = rows[0]
            used.add(rows[0])
    return adoptions


def plan_placements(sequence, new_codes, code_to_row, last_gl_row):
    """For each new code decide where its row goes, using Sheet1's own
    section labels as boundaries. Returns dicts in Sheet1 order:
    insert_before (original row number) and attach_after (the row it hangs
    below, when it belongs to the END of that section, else None)."""
    plans = []
    for i, (kind, val) in enumerate(sequence):
        if kind != "code" or val not in new_codes:
            continue
        prev_row, boundary_before = None, False
        for j in range(i - 1, -1, -1):
            k2, v2 = sequence[j]
            if k2 == "label":
                boundary_before = True
            elif v2 in code_to_row:
                prev_row = code_to_row[v2]
                break
        next_row, boundary_after = None, False
        for j in range(i + 1, len(sequence)):
            k2, v2 = sequence[j]
            if k2 == "label":
                boundary_after = True
            elif v2 in code_to_row:
                next_row = code_to_row[v2]
                break
        if boundary_after and not boundary_before and prev_row is not None:
            before, after = prev_row + 1, prev_row
        elif next_row is not None:
            before, after = next_row, None
        elif prev_row is not None:
            before, after = prev_row + 1, prev_row
        else:
            before, after = last_gl_row + 1, last_gl_row
        plans.append({"insert_before": before, "attach_after": after, "code": val})
    return plans


def insert_new_gl_rows(ws, new_rows_by_code, plans, unit_map, write_total):
    """Insert the planned rows, rewriting every row-referencing formula.
    write_total=False (new-month flow): the new row is 0 in every existing
    month column — the new month's value is written after the column shift,
    so nothing lands in the previous month. write_total=True (replace
    flow): the newest column gets the real value."""
    if not plans:
        return ws.max_row
    plans = sorted(plans, key=lambda p: p["insert_before"])  # stable
    before_rows = [p["insert_before"] for p in plans]
    after_rows = [p["attach_after"] for p in plans if p["attach_after"] is not None]

    def remap(orig_row):
        return orig_row + sum(1 for b in before_rows if b <= orig_row)

    def end_extra(orig_row):
        return sum(1 for a in after_rows if a == orig_row)

    orig_max_row = ws.max_row
    max_col = ws.max_column
    month_cols = get_month_data_columns(ws)

    snapshot = {}
    for r in range(1, orig_max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                val = remap_formula_rows(val, remap, end_extra)
            snapshot[(r, c)] = (val, copy.copy(cell._style))

    final_max_row = orig_max_row + len(plans)
    for r in range(1, final_max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).value = None
    for orig_r in range(1, orig_max_row + 1):
        final_r = remap(orig_r)
        for c in range(1, max_col + 1):
            val, style = snapshot[(orig_r, c)]
            cell = ws.cell(row=final_r, column=c)
            cell.value = val
            cell._style = copy.copy(style)

    for idx, p in enumerate(plans):
        final_r = p["insert_before"] + idx
        data = new_rows_by_code[p["code"]]
        style_row = final_r - 1 if p["attach_after"] is not None else final_r + 1
        for c in range(1, max_col + 1):
            ws.cell(row=final_r, column=c)._style = copy.copy(ws.cell(row=style_row, column=c)._style)
        ws.cell(row=final_r, column=GL_COL).value = data["code"]
        ws.cell(row=final_r, column=DESC_COL).value = data["desc"]
        for col, i in unit_map:
            v = data["units"][i]
            ws.cell(row=final_r, column=col).value = v if v is not None else 0
        for c in month_cols:
            ws.cell(row=final_r, column=c).value = 0
        if write_total:
            ws.cell(row=final_r, column=TOTAL_COL).value = data["total"] if data["total"] is not None else 0
    return final_max_row


def shift_month_columns(ws, max_row):
    """Shift the real, dated month columns one column right, making room
    for the new month at the first month column. Anything beyond the last
    real month column is left untouched."""
    real_cols = get_month_data_columns(ws)
    last_real_col = max(real_cols) if real_cols else TOTAL_COL - 1

    snapshot = {}
    for r in range(1, max_row + 1):
        for c in range(TOTAL_COL, last_real_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                val = remap_formula_columns(val, TOTAL_COL, delta=1)
            snapshot[(r, c)] = (val, copy.copy(cell._style))
    for r in range(1, max_row + 1):
        for c in range(TOTAL_COL, last_real_col + 1):
            val, style = snapshot[(r, c)]
            dest = ws.cell(row=r, column=c + 1)
            dest.value = val
            dest._style = copy.copy(style)
    for r in range(1, max_row + 1):
        new_cell = ws.cell(row=r, column=TOTAL_COL)
        new_cell.value = None
        new_cell._style = copy.copy(ws.cell(row=r, column=TOTAL_COL + 1)._style)
    ws.column_dimensions[get_column_letter(TOTAL_COL)].width = \
        ws.column_dimensions[get_column_letter(TOTAL_COL + 1)].width
    return last_real_col + 1


# ------------------------------------------------------------ manual rows

def _suggestion_text(v):
    return None if v is None else str(v)


def classify_and_fill_manual_row(ws, r, new_month_date, old_month_cols, gross_sales=None):
    """Fill the newest month cell of a row that Sheet1 does not supply, only
    when a safe repeating pattern exists. Returns a short tag."""
    desc = (ws.cell(row=r, column=DESC_COL).value or "").strip()
    old_l_col = old_month_cols[0]
    old_l_val = ws.cell(row=r, column=old_l_col).value

    if desc.lower().startswith(GROSS_SALES_LABEL_PREFIX):
        if gross_sales is not None:
            ws.cell(row=r, column=TOTAL_COL).value = gross_sales
            return "gross-sales-matched"
        ws.cell(row=r, column=TOTAL_COL).value = None
        return "left-blank"

    if desc.upper() in SUBTOTAL_DESCRIPTIONS:
        if isinstance(old_l_val, str) and old_l_val.startswith('='):
            ws.cell(row=r, column=TOTAL_COL).value = _retarget_column(old_l_val, old_l_col, TOTAL_COL)
            return "template"
        return "no-formula-found"

    if isinstance(old_l_val, str) and old_l_val.startswith('='):
        m = TRAILING_MULT_RE.search(old_l_val)
        if m and int(m.group(2)) == fiscal_month_number(_prev_month(new_month_date)):
            new_formula = old_l_val[:m.start(2)] + str(fiscal_month_number(new_month_date)) + old_l_val[m.end(2):]
            ws.cell(row=r, column=TOTAL_COL).value = _retarget_column(new_formula, old_l_col, TOTAL_COL)
            return "fiscal-accrual"
        return "manual-formula-left-blank"

    vals = [ws.cell(row=r, column=c).value for c in old_month_cols]
    if len(vals) >= 2 and all(isinstance(v, (int, float)) for v in vals) and len(set(vals)) == 1:
        ws.cell(row=r, column=TOTAL_COL).value = vals[0]
        return "flat-value"
    return "left-blank"


# ---------------------------------------------------------------- pipeline

def _match_rows(ws, s1, write_total):
    """Attach codes to matching manual rows, insert genuinely new rows.
    Returns (max_row, new_codes, adopted)."""
    sheet1_by_code = {r["code"]: r for r in s1.rows}
    existing = get_existing_gl_rows(ws)
    code_to_row = code_to_row_map(existing)
    unmapped = [r for r in s1.rows if r["code"] not in code_to_row]

    adoptions = find_adoptions(ws, unmapped)
    adopted = []
    for code, row in adoptions.items():
        ws.cell(row=row, column=GL_COL).value = code
        code_to_row[code] = row
        adopted.append((code, sheet1_by_code[code]["desc"], row))

    new_codes = [r["code"] for r in unmapped if r["code"] not in adoptions]
    last_gl_row = existing[-1][0] if existing else FIRST_DATA_ROW
    plans = plan_placements(s1.sequence, set(new_codes), code_to_row, last_gl_row)
    unit_map = build_unit_map(ws, s1.unit_headers)
    max_row = insert_new_gl_rows(ws, {c: sheet1_by_code[c] for c in new_codes},
                                 plans, unit_map, write_total)
    return max_row, new_codes, adopted, unit_map


def _fill_gl_rows(ws, s1, unit_map, new_month_date, old_month_cols):
    """Write the newest month for every coded row from Sheet1. Rows whose
    codes are all absent from Sheet1 fall back to the manual-row patterns."""
    sheet1_by_code = {r["code"]: r for r in s1.rows}
    partial, manual, suggestions = [], [], {}
    for r, code_cell, desc in get_existing_gl_rows(ws):
        codes = parse_codes(code_cell)
        found = [sheet1_by_code[c] for c in codes if c in sheet1_by_code]
        if found:
            total = sum((f["total"] or 0) for f in found)
            ws.cell(row=r, column=TOTAL_COL).value = total
            units = _sum_units(found)
            for col, i in unit_map:
                ws.cell(row=r, column=col).value = units[i]
            if len(found) < len(codes):
                partial.append((r, code_cell, desc))
        elif old_month_cols:
            old_val = ws.cell(row=r, column=old_month_cols[0]).value
            if old_val in (None, 0):
                ws.cell(row=r, column=TOTAL_COL).value = 0  # inactive account: stays 0
                manual.append((r, desc, "flat-value"))
                continue
            tag = classify_and_fill_manual_row(ws, r, new_month_date, old_month_cols)
            manual.append((r, desc, tag))
            if tag in ("left-blank", "manual-formula-left-blank") and old_val is not None:
                suggestions[r] = _suggestion_text(old_val)
    return partial, manual, suggestions


def _guard_new_month(new_month_date, current_month):
    if current_month is not None and new_month_date <= current_month:
        if new_month_date == current_month:
            raise NotANewMonthError(
                f"The uploaded Sheet1 is for {new_month_date.strftime('%B %Y')}, "
                f"which is already the latest month in the P&L workbook. "
                f"Looks like a duplicate upload — nothing was changed.")
        raise NotANewMonthError(
            f"The uploaded Sheet1 is for {new_month_date.strftime('%B %Y')}, "
            f"but the P&L workbook already has {current_month.strftime('%B %Y')} "
            f"as its newest month. Upload a month after that instead.")


def apply_new_month(wb):
    """Mutates wb in place. Returns a report dict."""
    ws = wb["P&L"]
    detect_layout(ws)
    s1 = read_sheet1(wb["Sheet1"])
    new_month_date = s1.month_date
    _guard_new_month(new_month_date, get_current_pnl_month(ws))

    max_row, new_codes, adopted, unit_map = _match_rows(ws, s1, write_total=False)

    real_cols_before = get_month_data_columns(ws)
    shift_month_columns(ws, max_row)
    old_month_cols = [c + 1 for c in real_cols_before]

    ws.cell(row=DATE_ROW, column=TOTAL_COL).value = datetime.datetime(
        new_month_date.year, new_month_date.month, new_month_date.day)
    ws.cell(row=HEADER_ROW, column=TOTAL_COL).value = "Amount"
    title_cell = ws.cell(row=4, column=3)
    if title_cell.value:
        title_cell.value = re.sub(r'as on\s+[A-Za-z]+[-\s]+\d{4}',
                                  f"as on {new_month_date.strftime('%B-%Y')}", title_cell.value)

    partial, manual_from_gl, suggestions = _fill_gl_rows(ws, s1, unit_map, new_month_date, old_month_cols)

    manual_report = list(manual_from_gl)
    for r in range(FIRST_DATA_ROW, max_row + 1):
        if ws.cell(row=r, column=GL_COL).value:
            continue
        desc = ws.cell(row=r, column=DESC_COL).value
        if not desc:
            continue
        old_val = ws.cell(row=r, column=old_month_cols[0]).value
        tag = classify_and_fill_manual_row(ws, r, new_month_date, old_month_cols, s1.gross_sales)
        manual_report.append((r, desc, tag))
        if tag in ("left-blank", "manual-formula-left-blank") and old_val is not None:
            suggestions[r] = _suggestion_text(old_val)

    sheet1_by_code = {r["code"]: r for r in s1.rows}
    return {
        "new_month": new_month_date.isoformat(),
        "new_gl_rows_inserted": [(c, sheet1_by_code[c]["desc"]) for c in new_codes],
        "adopted_rows": adopted,
        "gl_rows_zeroed": [],
        "gl_rows_partial": partial,
        "manual_rows": manual_report,
        "suggestions": suggestions,
    }


def replace_latest_month(wb):
    """Corrected/final data for the month already merged in: re-populate
    the newest column in place — no new column, no shifting. Non-coded
    manual rows are left untouched."""
    ws = wb["P&L"]
    detect_layout(ws)
    s1 = read_sheet1(wb["Sheet1"])
    new_month_date = s1.month_date

    current_month = get_current_pnl_month(ws)
    if current_month is None:
        raise NotTheLatestMonthError(
            "The P&L workbook doesn't have a latest month yet to replace. "
            "Use the 'New month' option instead.")
    if new_month_date != current_month:
        raise NotTheLatestMonthError(
            f"This file is for {new_month_date.strftime('%B %Y')}, but the "
            f"P&L workbook's latest month is {current_month.strftime('%B %Y')}. "
            f"'Replace latest month' only works when the uploaded file is for "
            f"that same month — use 'New month' if this is meant to move forward.")

    max_row, new_codes, adopted, unit_map = _match_rows(ws, s1, write_total=True)
    old_month_cols = get_month_data_columns(ws)[1:]
    partial, manual, suggestions = _fill_gl_rows(ws, s1, unit_map, new_month_date, old_month_cols)

    sheet1_by_code = {r["code"]: r for r in s1.rows}
    return {
        "new_month": new_month_date.isoformat(),
        "new_gl_rows_inserted": [(c, sheet1_by_code[c]["desc"]) for c in new_codes],
        "adopted_rows": adopted,
        "gl_rows_zeroed": [],
        "gl_rows_partial": partial,
        "manual_rows": manual,
        "suggestions": suggestions,
    }
