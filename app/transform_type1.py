"""
Type 1 transformation: merge a new month's GL trial-balance sheet ("Sheet1")
into the running P&L workbook.

Rules (as confirmed with the user):
- Description is the primary match key between Sheet1 and P&L, for rows
  that carry a GL Code. GL Code is used as a secondary integrity check.
- The new month becomes the first data column (L) in P&L; every existing
  month column shifts one column to the right (formulas' internal column
  letters are rewritten so they keep pointing at the right column).
- If a GL row's description exists in P&L but not in the new Sheet1 ->
  that row's new-month value is 0.
- If a GL row's description exists in the new Sheet1 but not in P&L -> a
  new row is inserted into P&L in the exact same relative position it has
  in Sheet1, with 0 for every historical month column and the real value
  for the new month. Row-referencing formulas below the insertion point
  are rewritten to keep summing the right ranges.
- Rows with no GL Code (subtotals + manual accountant entries) are never
  matched against Sheet1. For the new month column, we only auto-fill
  these when a safe, repeatable pattern is evident from the existing
  columns:
    * "template" rows: the formula shape (cell/range references, all in
      the row's own column) is identical across all existing month
      columns -> copied forward with the column letter updated to L.
    * "fiscal accrual" rows: formula ends in a `* N` (or `/12*N`) style
      cumulative-fiscal-month multiplier that increases by 1 each month
      (reset to 1 at the fiscal year start, April) -> N is bumped.
    * "flat value" rows: the same hardcoded number repeats, unchanged,
      across every existing month column -> copied forward as-is.
    * anything else is left blank and reported back so a human fills it in.
"""
import re
import copy
import datetime
from openpyxl.utils import get_column_letter, column_index_from_string

MONTHS = {m.lower(): i for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July",
     "August", "September", "October", "November", "December"], start=1)}

FISCAL_YEAR_START_MONTH = 4  # April

CELL_REF_RE = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d{1,4})')
TRAILING_MULT_RE = re.compile(r'([*])(\d{1,2})(\s*)$')

GL_COL = 2          # column B
DESC_COL = 3        # column C
UNIT_COLS = list(range(4, 12))   # D..K
TOTAL_COL = 12      # column L is always the "current/new" month slot


def parse_month_from_title(title):
    """'Profit & Loss Account as on July-2026' / 'as on July 2026' -> date(2026,7,1)"""
    m = re.search(r'as on\s+([A-Za-z]+)[-\s]+(\d{4})', title)
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


def read_sheet1(ws):
    """Returns (month_date, ordered list of dict(code, desc, units[8], total))"""
    title = ws.cell(row=4, column=3).value or ws.cell(row=4, column=2).value
    month_date = parse_month_from_title(title)
    rows = []
    for row in ws.iter_rows(min_row=7, max_row=ws.max_row):
        code = row[GL_COL - 1].value
        desc = row[DESC_COL - 1].value
        if not code:
            continue
        units = [row[c - 1].value for c in UNIT_COLS]
        total = row[TOTAL_COL - 1].value
        rows.append({"code": code, "desc": desc, "units": units, "total": total})
    return month_date, rows


def _shift_col_letter(letter, shift_from_col, delta=1):
    idx = column_index_from_string(letter)
    if idx >= shift_from_col:
        return get_column_letter(idx + delta)
    return letter


def remap_formula_columns(formula, shift_from_col, delta=1):
    def _sub(m):
        dollar_col, col, dollar_row, row_s = m.groups()
        new_col = _shift_col_letter(col, shift_from_col, delta)
        return f"{dollar_col}{new_col}{dollar_row}{row_s}"
    return CELL_REF_RE.sub(_sub, formula)


def remap_formula_rows(formula, remap_fn):
    def _sub(m):
        dollar_col, col, dollar_row, row_s = m.groups()
        row = int(row_s)
        return f"{dollar_col}{col}{dollar_row}{remap_fn(row)}"
    return CELL_REF_RE.sub(_sub, formula)


HEADER_ROW = 6
FIRST_DATA_ROW = 8


def get_existing_gl_rows(ws):
    """Ordered list of (row_idx, code, desc) for rows carrying a GL code."""
    out = []
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        code = ws.cell(row=r, column=GL_COL).value
        if code:
            out.append((r, code, ws.cell(row=r, column=DESC_COL).value))
    return out


def get_month_data_columns(ws):
    """Columns (in order) that actually hold a month's Amount data, per the
    date header in row 7 starting at column L."""
    cols = []
    c = TOTAL_COL
    while ws.cell(row=7, column=c).value is not None:
        cols.append(c)
        c += 1
    return cols


def insert_new_gl_rows(ws, new_rows_by_code, existing_gl_rows, sheet1_order):
    """
    new_rows_by_code: dict code -> sheet1 row dict (only codes not already in P&L)
    existing_gl_rows: list of (row_idx, code, desc) currently in P&L (GL rows only)
    sheet1_order: full ordered list of sheet1 row dicts (defines target order)
    Returns updated max_row after insertion, and the ROW-REMAP function to apply
    to every formula elsewhere in the sheet.
    """
    if not new_rows_by_code:
        return ws.max_row, (lambda r: r)

    existing_code_to_row = {code: r for r, code, _ in existing_gl_rows}
    sheet1_codes = [row["code"] for row in sheet1_order]

    # For each new code, find the P&L row of the next sheet1 code that DOES
    # already exist in P&L -> that's the "insert before" anchor.
    breakpoints = []  # (orig_row_to_insert_before, code)
    for i, code in enumerate(sheet1_codes):
        if code not in new_rows_by_code:
            continue
        anchor_row = None
        for j in range(i + 1, len(sheet1_codes)):
            if sheet1_codes[j] in existing_code_to_row:
                anchor_row = existing_code_to_row[sheet1_codes[j]]
                break
        if anchor_row is None:
            # no later anchor -> append after the last existing GL row
            anchor_row = existing_gl_rows[-1][0] + 1
        breakpoints.append((anchor_row, code))

    breakpoints.sort(key=lambda x: x[0])
    bp_rows = [bp for bp, _ in breakpoints]

    def remap(orig_row):
        return orig_row + sum(1 for bp in bp_rows if bp <= orig_row)

    orig_max_row = ws.max_row
    max_col = ws.max_column

    # snapshot with row-remapped formulas
    snapshot = {}
    for r in range(1, orig_max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                val = remap_formula_rows(val, remap)
            snapshot[(r, c)] = {"value": val, "style": copy.copy(cell._style)}

    sorted_bp = sorted(breakpoints, key=lambda x: x[0])
    new_row_final_pos = {code: bp + idx for idx, (bp, code) in enumerate(sorted_bp)}

    final_max_row = orig_max_row + len(breakpoints)

    for r in range(1, final_max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).value = None

    for orig_r in range(1, orig_max_row + 1):
        final_r = remap(orig_r)
        for c in range(1, max_col + 1):
            data = snapshot[(orig_r, c)]
            cell = ws.cell(row=final_r, column=c)
            cell.value = data["value"]
            cell._style = copy.copy(data["style"])

    month_header_cols = [c for c in range(TOTAL_COL, max_col + 1)
                          if ws.cell(row=7, column=c).value is not None
                          or c in (TOTAL_COL, TOTAL_COL + 1, TOTAL_COL + 2)]

    for code, final_r in new_row_final_pos.items():
        data = new_rows_by_code[code]
        style_source_row = final_r + 1
        for c in range(1, max_col + 1):
            src_style = ws.cell(row=style_source_row, column=c)._style
            ws.cell(row=final_r, column=c)._style = copy.copy(src_style)
        ws.cell(row=final_r, column=GL_COL).value = code
        ws.cell(row=final_r, column=DESC_COL).value = data["desc"]
        for i, v in enumerate(data["units"]):
            ws.cell(row=final_r, column=UNIT_COLS[0] + i).value = v if v is not None else 0
        ws.cell(row=final_r, column=TOTAL_COL).value = data["total"] if data["total"] is not None else 0
        for c in range(TOTAL_COL + 1, max_col + 1):
            if c in month_header_cols:
                ws.cell(row=final_r, column=c).value = 0

    return final_max_row, remap


def shift_month_columns(ws, max_row):
    """Shift only the real, dated month columns (L onward, per row 7's
    date header) one column to the right, making room for the new month
    at L. Anything sitting beyond the last real month column (leftover
    debris, unrelated notes, etc.) is deliberately left untouched instead
    of being dragged along by a sheet-wide column insert.

    Rewrites column-letter references inside formulas so each shifted
    formula still refers to its own (new) column, and clones each row's
    style from the column it displaced onto the brand-new L column
    (which openpyxl leaves unstyled by default)."""
    real_cols = get_month_data_columns(ws)
    last_real_col = max(real_cols) if real_cols else TOTAL_COL - 1

    # snapshot the current L..last_real_col range before overwriting it
    snapshot = {}
    for r in range(1, max_row + 1):
        for c in range(TOTAL_COL, last_real_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                val = remap_formula_columns(val, TOTAL_COL, delta=1)
            snapshot[(r, c)] = {"value": val, "style": copy.copy(cell._style)}

    # write the snapshot back one column to the right
    for r in range(1, max_row + 1):
        for c in range(TOTAL_COL, last_real_col + 1):
            data = snapshot[(r, c)]
            dest = ws.cell(row=r, column=c + 1)
            dest.value = data["value"]
            dest._style = copy.copy(data["style"])

    # the fresh L column: blank values, styled like the column it displaced
    for r in range(1, max_row + 1):
        new_l_cell = ws.cell(row=r, column=TOTAL_COL)
        new_l_cell.value = None
        new_l_cell._style = copy.copy(ws.cell(row=r, column=TOTAL_COL + 1)._style)

    ws.column_dimensions[get_column_letter(TOTAL_COL)].width = \
        ws.column_dimensions[get_column_letter(TOTAL_COL + 1)].width

    return last_real_col + 1


SUBTOTAL_DESCRIPTIONS = {
    "NET SALES", "OTHER INCOME", "CLOSING STOCK", "TOTAL INCOME",
    "TOTAL EXPENSES", "PROFIT & LOSS BEFORE DEPRECIATION", "PROFIT & LOSS",
}


def _retarget_column(formula, from_col, to_col):
    """Replace only exact references to from_col (a 1-based column index)
    with to_col's letter; references to any other column are left as-is."""
    from_letter = get_column_letter(from_col)
    to_letter = get_column_letter(to_col)

    def _sub(m):
        dollar_col, col, dollar_row, row_s = m.groups()
        if col == from_letter:
            col = to_letter
        return f"{dollar_col}{col}{dollar_row}{row_s}"
    return CELL_REF_RE.sub(_sub, formula)


def classify_and_fill_manual_row(ws, r, new_month_date, old_month_cols):
    """old_month_cols: the *new* column positions of what used to be
    L, M, N, O ... (after the general column shift already ran, so any
    cross-references inside these formulas already point at the right
    shifted columns; only the row's *own* column letter still needs to
    be retargeted from old_l_col back to L). Fills ws cell at
    (r, TOTAL_COL) if a safe pattern is found. Returns a short tag."""
    desc = (ws.cell(row=r, column=DESC_COL).value or "").strip()
    old_l_col = old_month_cols[0]
    old_l_val = ws.cell(row=r, column=old_l_col).value

    if desc.upper() in SUBTOTAL_DESCRIPTIONS:
        if isinstance(old_l_val, str) and old_l_val.startswith('='):
            new_formula = _retarget_column(old_l_val, old_l_col, TOTAL_COL)
            ws.cell(row=r, column=TOTAL_COL).value = new_formula
            return "template"
        return "no-formula-found"

    if isinstance(old_l_val, str) and old_l_val.startswith('='):
        m = TRAILING_MULT_RE.search(old_l_val)
        if m:
            old_mult = int(m.group(2))
            old_fm = fiscal_month_number(_prev_month(new_month_date))
            if old_mult == old_fm:
                new_mult = fiscal_month_number(new_month_date)
                new_formula = old_l_val[:m.start(2)] + str(new_mult) + old_l_val[m.end(2):]
                new_formula = _retarget_column(new_formula, old_l_col, TOTAL_COL)
                ws.cell(row=r, column=TOTAL_COL).value = new_formula
                return "fiscal-accrual"
        return "manual-formula-left-blank"

    # flat value check: same hardcoded number across all existing month cols?
    vals = [ws.cell(row=r, column=c).value for c in old_month_cols]
    if len(vals) >= 2 and all(isinstance(v, (int, float)) for v in vals) and len(set(vals)) == 1:
        ws.cell(row=r, column=TOTAL_COL).value = vals[0]
        return "flat-value"

    return "left-blank"


def _prev_month(d):
    if d.month == 1:
        return datetime.date(d.year - 1, 12, 1)
    return datetime.date(d.year, d.month - 1, 1)


class NotANewMonthError(Exception):
    """Raised when the uploaded Sheet1's month is not strictly after the
    P&L workbook's current newest month (duplicate upload, wrong file,
    or an attempt to go backwards)."""


def get_current_pnl_month(ws):
    """The date already sitting in the newest (column L) month header,
    or None if the sheet has no month data yet."""
    d = ws.cell(row=7, column=TOTAL_COL).value
    if isinstance(d, datetime.datetime):
        return d.date()
    return None


def apply_new_month(wb):
    """Mutates wb in place. Returns a report dict."""
    ws = wb["P&L"]
    s1 = wb["Sheet1"]

    new_month_date, sheet1_rows = read_sheet1(s1)

    current_month = get_current_pnl_month(ws)
    if current_month is not None and new_month_date <= current_month:
        if new_month_date == current_month:
            raise NotANewMonthError(
                f"The uploaded Sheet1 is for {new_month_date.strftime('%B %Y')}, "
                f"which is already the latest month in the P&L workbook. "
                f"Looks like a duplicate upload — nothing was changed."
            )
        raise NotANewMonthError(
            f"The uploaded Sheet1 is for {new_month_date.strftime('%B %Y')}, "
            f"but the P&L workbook already has {current_month.strftime('%B %Y')} "
            f"as its newest month. Upload a month after that instead."
        )

    sheet1_by_code = {row["code"]: row for row in sheet1_rows}

    existing_gl_rows = get_existing_gl_rows(ws)
    existing_codes = {code for _, code, _ in existing_gl_rows}
    new_codes = [row["code"] for row in sheet1_rows if row["code"] not in existing_codes]
    new_rows_by_code = {c: sheet1_by_code[c] for c in new_codes}

    max_row, _ = insert_new_gl_rows(ws, new_rows_by_code, existing_gl_rows, sheet1_rows)

    real_month_cols_before_shift = get_month_data_columns(ws)

    new_max_col = shift_month_columns(ws, max_row)
    # after shift, the real (dated) month columns each moved one to the right
    shifted_real_month_cols = [c + 1 for c in real_month_cols_before_shift]

    # new month header/date
    ws.cell(row=7, column=TOTAL_COL).value = datetime.datetime(
        new_month_date.year, new_month_date.month, new_month_date.day)
    ws.cell(row=6, column=TOTAL_COL).value = "Amount"

    # refresh title rows to mention the new month name, if present
    title_cell = ws.cell(row=4, column=3)
    if title_cell.value:
        month_str = new_month_date.strftime("%B-%Y")
        title_cell.value = re.sub(r'as on\s+[A-Za-z]+[-\s]+\d{4}',
                                   f'as on {month_str}', title_cell.value)

    manual_report = []
    zero_filled = []
    existing_gl_rows_after = get_existing_gl_rows(ws)

    for r, code, desc in existing_gl_rows_after:
        row = sheet1_by_code.get(code)
        if row is not None:
            for i, c in enumerate(UNIT_COLS):
                ws.cell(row=r, column=c).value = row["units"][i]
            ws.cell(row=r, column=TOTAL_COL).value = row["total"] if row["total"] is not None else 0
        else:
            ws.cell(row=r, column=TOTAL_COL).value = 0
            zero_filled.append((r, code, desc))

    for r in range(FIRST_DATA_ROW, max_row + 1):
        code = ws.cell(row=r, column=GL_COL).value
        if code:
            continue
        desc = ws.cell(row=r, column=DESC_COL).value
        if not desc:
            continue
        tag = classify_and_fill_manual_row(ws, r, new_month_date, shifted_real_month_cols)
        manual_report.append((r, desc, tag))

    return {
        "new_month": new_month_date.isoformat(),
        "new_gl_rows_inserted": [(c, sheet1_by_code[c]["desc"]) for c in new_codes],
        "gl_rows_zeroed": zero_filled,
        "manual_rows": manual_report,
    }


class NotTheLatestMonthError(Exception):
    """Raised when a 'replace latest month' upload's month doesn't match
    the P&L workbook's current newest month."""


def replace_latest_month(wb):
    """For when corrected/final data arrives for a month that's already
    been merged in (e.g. a few more August sales came in after the fact).
    Re-populates the existing latest-month column (L) from the new
    Sheet1, in place — no new column is added, no shifting happens.
    Any GL row in the new Sheet1 that wasn't already in P&L is still
    inserted at the correct position (0 for every older month, the real
    value for the latest month), same as apply_new_month does. Non-GL
    rows (subtotals, manual accrual formulas) are left untouched, since
    they already carry whatever was right for this same month and don't
    need to be recomputed. Mutates wb in place. Returns a report dict."""
    ws = wb["P&L"]
    s1 = wb["Sheet1"]

    new_month_date, sheet1_rows = read_sheet1(s1)

    current_month = get_current_pnl_month(ws)
    if current_month is None:
        raise NotTheLatestMonthError(
            "The P&L workbook doesn't have a latest month yet to replace. "
            "Use the 'New month' option instead."
        )
    if new_month_date != current_month:
        raise NotTheLatestMonthError(
            f"This file is for {new_month_date.strftime('%B %Y')}, but the "
            f"P&L workbook's latest month is {current_month.strftime('%B %Y')}. "
            f"'Replace latest month' only works when the uploaded file is for "
            f"that same month — use 'New month' if this is meant to move forward."
        )

    sheet1_by_code = {row["code"]: row for row in sheet1_rows}

    existing_gl_rows = get_existing_gl_rows(ws)
    existing_codes = {code for _, code, _ in existing_gl_rows}
    new_codes = [row["code"] for row in sheet1_rows if row["code"] not in existing_codes]
    new_rows_by_code = {c: sheet1_by_code[c] for c in new_codes}

    max_row, _ = insert_new_gl_rows(ws, new_rows_by_code, existing_gl_rows, sheet1_rows)

    zero_filled = []
    existing_gl_rows_after = get_existing_gl_rows(ws)
    for r, code, desc in existing_gl_rows_after:
        row = sheet1_by_code.get(code)
        if row is not None:
            for i, c in enumerate(UNIT_COLS):
                ws.cell(row=r, column=c).value = row["units"][i]
            ws.cell(row=r, column=TOTAL_COL).value = row["total"] if row["total"] is not None else 0
        else:
            ws.cell(row=r, column=TOTAL_COL).value = 0
            zero_filled.append((r, code, desc))

    return {
        "new_month": new_month_date.isoformat(),
        "new_gl_rows_inserted": [(c, sheet1_by_code[c]["desc"]) for c in new_codes],
        "gl_rows_zeroed": zero_filled,
        "manual_rows": [],
    }
