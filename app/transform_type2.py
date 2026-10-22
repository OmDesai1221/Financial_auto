"""
Type 2 transformation ("ACPLACT"): merges a new month's GL trial balance
(the same monthly Sheet1 upload used by Type 1) into the two-sheet FIFO
workbook (acplact.xlsx): 'FIFO-sch' (detail) and 'FIFO - P & L' (summary).

Layout, per month, in both sheets:
  - two columns: "movement" (this month's amount minus last month's) then
    "amount" (the actual balance as of that month), newest month leftmost.
  - the single oldest column (fiscal-year start, April) has amount only,
    since there's nothing before it in-sheet to diff against.
  - FIFO-sch's GL Code column can list several codes comma-separated for a
    row that merges multiple GL accounts (e.g. "X2004006,X2004005,X2004007")
    — that row's amount is the sum of every listed code's Sheet1 balance.
  - FIFO-sch's GL Code is the same permanent match key as Type 1's; two
    rows in 'FIFO - P & L' (Directors Remuneration, Perf. Incentive to
    Directors) also carry a GL code directly and get their amount from
    Sheet1 the same way, instead of their old hand-typed formula.
  - Every other formula (subtotals, the Opening/Closing-stock rollforward,
    the cross-sheet % line, etc.) is carried forward for the new month by
    literally reusing the previous newest column's formula text as-is —
    every pattern found in this sheet references only its own row/column
    or a fixed structural neighbor (G/I, or H for "last period's slot"),
    so a verbatim copy stays correct after the shift. The one exception:
    a formula shaped "=G{r}-I{r}-<one-off adjustment>" has that trailing
    adjustment stripped for the new month (it's plainly a one-time reclass
    from whichever month it was first entered for) and is reported back
    so a human can decide whether to re-enter something similar.
"""
import re
import copy
from sheet1_formulas import compose
import datetime
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.cell.cell import MergedCell
from transform_type1 import remap_formula_rows

FIFO_SCH_SHEET = "FIFO-sch"
FIFO_PNL_SHEET = "FIFO - P & L"

SCH_GL_COL = 4        # D
SCH_DESC_COL = 5       # E
SCH_ANCHOR_COL = 6     # F (movement); amount is anchor + 1 (G)
SCH_DATE_ROW = 5
SCH_HEADER_ROW = 4
SCH_FIRST_DATA_ROW = 7
SCH_NET_SALES_ROW = 13  # FIFO-sch's own live Net Sales total row

PNL_GL_COL = 2         # B (only 2 rows use this)
PNL_DESC_COL = 3       # C
PNL_ANCHOR_COL = 5     # E (movement); amount is anchor + 1 (F)
PNL_HEADER_ROW = 4
PNL_FIRST_DATA_ROW = 6
PNL_RATIO_DATE_ROW = 5
PNL_GROSS_SALES_ROW = 6   # from Sheet1's "Gross sales" label row directly
PNL_NET_SALES_ROW = 8    # pulled live from FIFO-sch's Net Sales total

FISCAL_YEAR_START_MONTH = 4  # April

CELL_REF_RE = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d{1,4})')
DIFF_WITH_TAIL_RE = re.compile(r'^=G(\d+)-I\1(.+)$')
NUMERIC_TAIL_RE = re.compile(r'^(=[^0-9]*?-?\+?\$?[A-Z]{1,3}\$?\d{1,4})[+-]\d{4,}(?:\.\d+)?$')
TITLE_DATE_RE = re.compile(r'(year ended\s+)\d{1,2}\.\d{1,2}\.\d{4}', re.I)
TRAILING_MULT_RE = re.compile(r'([*])(\d{1,2})(\s*)$')


def fiscal_month_number(d):
    """1 for April, 2 for May, ... 12 for March."""
    return (d.month - FISCAL_YEAR_START_MONTH) % 12 + 1


def _prev_month(d):
    if d.month == 1:
        return datetime.date(d.year - 1, 12, 1)
    return datetime.date(d.year, d.month - 1, 1)


class NotANewMonthError(Exception):
    pass


def parse_sch_date(text):
    """'31.07.2026' -> date(2026,7,31)"""
    if isinstance(text, datetime.datetime):
        return text.date()
    m = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', str(text).strip())
    if not m:
        raise ValueError(f"Could not parse FIFO-sch date header: {text!r}")
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return datetime.date(y, mo, d)


def last_day_of_month(d):
    if d.month == 12:
        next_first = datetime.date(d.year + 1, 1, 1)
    else:
        next_first = datetime.date(d.year, d.month + 1, 1)
    return next_first - datetime.timedelta(days=1)


def fmt_sch_date(d):
    return last_day_of_month(d).strftime("%d.%m.%Y")


def is_fiscal_year_start(d):
    return d.month == FISCAL_YEAR_START_MONTH


GROSS_SALES_LABEL_PREFIX = "gross sale"   # matches "GROSS SALE" and "GROSS SALES"


def read_sheet1_totals(sheet1_ws):
    """code -> Total Balance (column L), same layout as Type 1's Sheet1.
    Also picks out the 'Gross sales' label row's own Total Balance, if the
    client has typed a value there — that row carries no GL code, so it's
    invisible to the code-keyed totals dict otherwise."""
    totals = {}
    gross_sales = None
    for row in sheet1_ws.iter_rows(min_row=7, max_row=sheet1_ws.max_row):
        code = row[1].value  # column B
        desc = row[2].value  # column C
        if code:
            totals[code] = row[11].value  # column L
        elif isinstance(desc, str) and desc.strip().lower().startswith(GROSS_SALES_LABEL_PREFIX):
            gross_sales = row[11].value

    title = sheet1_ws.cell(row=4, column=2).value or sheet1_ws.cell(row=4, column=3).value
    m = re.search(r'as on\s+([A-Za-z]+)[-\s]+(\d{4})', title or "")
    if not m:
        raise ValueError(f"Could not parse month/year from Sheet1 title: {title!r}")
    months = {name.lower(): i for i, name in enumerate(
        ["January", "February", "March", "April", "May", "June", "July",
         "August", "September", "October", "November", "December"], start=1)}
    month_num = months.get(m.group(1).lower())
    new_month = datetime.date(int(m.group(2)), month_num, 1)
    return totals, new_month, gross_sales


# ----------------------------------------------------------- conflicts

def find_duplicate_codes(wb):
    """A Sheet1 code listed in more than one GL-code cell across this
    workbook gets summed into every one of those rows — silently double
    (or more) counted. Usually this is a slip made while hand-editing a
    merged code cell (e.g. a stamp-duty code accidentally left in both the
    stamp-duty row and a nearby interest row). Returns {code: [(sheet, row,
    description), ...]} for every code found more than once, so it can be
    flagged for a person to fix in the GL-code cell directly."""
    locations = {}
    sch = wb[FIFO_SCH_SHEET]
    for r in range(SCH_FIRST_DATA_ROW, sch.max_row + 1):
        for code in parse_gl_codes(sch.cell(row=r, column=SCH_GL_COL).value):
            locations.setdefault(code, []).append(
                (FIFO_SCH_SHEET, r, sch.cell(row=r, column=SCH_DESC_COL).value))
    pnl = wb[FIFO_PNL_SHEET]
    for r in range(PNL_FIRST_DATA_ROW, pnl.max_row + 1):
        for code in parse_gl_codes(pnl.cell(row=r, column=PNL_GL_COL).value):
            locations.setdefault(code, []).append(
                (FIFO_PNL_SHEET, r, pnl.cell(row=r, column=PNL_DESC_COL).value))
    return {code: locs for code, locs in locations.items() if len(locs) > 1}


# ------------------------------------------------------- new-field sections

SECTION_HEADER_MAX_INDENT = 1
PURE_SUM_RE = re.compile(
    r'^=(?:ROUND\()?SUM\(([A-Z]{1,3})(\d{1,4}):([A-Z]{1,3})(\d{1,4})\)(?:,\s*0\))?$')


def _is_section_header(ws, r):
    code = ws.cell(row=r, column=SCH_GL_COL).value
    desc = ws.cell(row=r, column=SCH_DESC_COL).value
    if code or not isinstance(desc, str) or not desc.strip():
        return False
    indent = len(desc) - len(desc.lstrip(' '))
    if indent > SECTION_HEADER_MAX_INDENT:
        return False
    cell = ws.cell(row=r, column=SCH_DESC_COL)
    return bool(cell.font and cell.font.bold)


def find_sections(ws):
    """FIFO-sch's bold, unindented row labels (Cost of Sales, Administration
    Expenses, Interest & Finance Charges, ...) mark each expense section;
    each one's subtotal is the next plain '=SUM(<col><start>:<col><end>)'
    formula found before the NEXT such header (a header with no subtotal of
    its own in that narrow window — e.g. a repeated/decorative label — is
    skipped). A brand-new GL code picked for a section is inserted directly
    above that subtotal row, and the subtotal's own range is widened to
    include it."""
    headers = [r for r in range(SCH_FIRST_DATA_ROW, ws.max_row + 1) if _is_section_header(ws, r)]
    sections = []
    for i, h in enumerate(headers):
        window_end = (headers[i + 1] - 1) if i + 1 < len(headers) else ws.max_row
        subtotal_row, item_end = None, None
        for r in range(h + 1, window_end + 1):
            formula = ws.cell(row=r, column=SCH_ANCHOR_COL + 1).value  # amount col
            m = PURE_SUM_RE.match(formula) if isinstance(formula, str) else None
            if m:
                subtotal_row = r
                item_end = int(m.group(4))   # the range's own last row — a blank
                break                          # spacer can sit between it and the subtotal
        if subtotal_row is None:
            continue  # no subtotal of its own before the next header — skip
        name = " ".join(str(ws.cell(row=h, column=SCH_DESC_COL).value).split())
        sections.append({"name": name, "header_row": h, "item_start": h + 1,
                          "item_end": item_end, "subtotal_row": subtotal_row})
    return sections


def insert_new_sch_rows(ws, assignments):
    """assignments: [{"code", "desc", "value", "insert_before"}], where
    insert_before is a section's subtotal_row (from find_sections) — the
    new row lands directly above it, everything from there down shifts by
    one, and any formula range ending exactly at that section's last item
    row grows to include the new row. Older months are left blank for a
    brand-new row (there's nothing to carry forward)."""
    if not assignments:
        return []
    before_rows = sorted(a["insert_before"] for a in assignments)

    def remap(orig_row):
        return orig_row + sum(1 for b in before_rows if b <= orig_row)

    def end_extra(orig_row):
        return sum(1 for a in assignments if a["insert_before"] - 1 == orig_row)

    orig_max_row = ws.max_row
    max_col = ws.max_column
    snapshot = {}
    for r in range(1, orig_max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                val = remap_formula_rows(val, remap, end_extra)
            snapshot[(r, c)] = (val, copy.copy(cell._style))

    final_max_row = orig_max_row + len(assignments)
    for r in range(1, final_max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None
    for orig_r in range(1, orig_max_row + 1):
        final_r = remap(orig_r)
        for c in range(1, max_col + 1):
            val, style = snapshot[(orig_r, c)]
            cell = ws.cell(row=final_r, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = val
            cell._style = copy.copy(style)

    new_rows = []
    seen_at = {}
    for a in assignments:
        slot = remap(a["insert_before"] - 1) + 1 + seen_at.get(a["insert_before"], 0)
        seen_at[a["insert_before"]] = seen_at.get(a["insert_before"], 0) + 1
        style_row = slot - 1 if slot > 1 else slot + 1
        for c in range(1, max_col + 1):
            ws.cell(row=slot, column=c)._style = copy.copy(ws.cell(row=style_row, column=c)._style)
        for col, val in ((SCH_GL_COL, a["code"]), (SCH_DESC_COL, a["desc"]), (SCH_ANCHOR_COL + 1, a["value"])):
            cell = ws.cell(row=slot, column=col)
            if not isinstance(cell, MergedCell):
                cell.value = val
        new_rows.append((slot, a["code"], a["desc"]))
    return new_rows


TOTAL_FORMULAS = {}    # code -> Sheet1 arithmetic formula; set per run
GROSS_FORMULA = None


def set_sheet1_formulas(by_code, gross):
    global TOTAL_FORMULAS, GROSS_FORMULA
    TOTAL_FORMULAS, GROSS_FORMULA = (by_code or {}), gross


def _formula_or_value(codes, totals, value, negate=False):
    found = [(c, totals[c]) for c in codes if totals.get(c) is not None]
    return compose(found, TOTAL_FORMULAS, negate) or (-value if negate else value)


def read_sheet1_descriptions(sheet1_ws):
    """code -> GL description, for reporting codes nothing maps."""
    return {row[1].value: row[2].value
            for row in sheet1_ws.iter_rows(min_row=7, max_row=sheet1_ws.max_row) if row[1].value}


def find_unmapped_codes(wb, totals, descriptions):
    """Sheet1 codes with a non-zero balance that no code cell in FIFO-sch or
    FIFO - P & L refers to. Their amounts are in Sheet1 but land nowhere in
    this workbook, so they must be shown to the user rather than dropped."""
    used = set()
    sch = wb[FIFO_SCH_SHEET]
    for r in range(SCH_FIRST_DATA_ROW, sch.max_row + 1):
        used.update(parse_gl_codes(sch.cell(row=r, column=SCH_GL_COL).value))
    pnl = wb[FIFO_PNL_SHEET]
    for r in range(PNL_FIRST_DATA_ROW, pnl.max_row + 1):
        used.update(parse_gl_codes(pnl.cell(row=r, column=PNL_GL_COL).value))
    out = []
    for code, val in totals.items():
        if code in used or not isinstance(val, (int, float)) or val == 0:
            continue
        out.append((code, descriptions.get(code) or "", val))
    return out


def parse_gl_codes(cell_value):
    if not cell_value:
        return []
    return [c.strip() for c in str(cell_value).split(",") if c.strip()]


def sum_for_codes(codes, totals):
    """Returns (sum, any_missing, all_missing). A code missing from this
    month's Sheet1 contributes 0 to the sum; if EVERY listed code is
    missing, the row isn't tracked by Sheet1 at all (a permanently manual
    line, e.g. Audit Fees or the FIFO stock rollforward) and the caller
    should fall back to treating it like any other manual row rather than
    zeroing it."""
    total = 0
    any_missing = False
    found_count = 0
    for c in codes:
        v = totals.get(c)
        if v is None:
            any_missing = True
        else:
            total += v
            found_count += 1
    all_missing = found_count == 0
    return total, any_missing, all_missing


def _retarget_column(formula, from_col, to_col):
    from_letter = get_column_letter(from_col)
    to_letter = get_column_letter(to_col)

    def _sub(m):
        dollar_col, col, dollar_row, row_s = m.groups()
        if col == from_letter:
            col = to_letter
        return f"{dollar_col}{col}{dollar_row}{row_s}"
    return CELL_REF_RE.sub(_sub, formula)


SHEET_REF_RE = re.compile(
    r"((?:'[^']+'|[A-Za-z_][A-Za-z0-9_]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d{1,4})(?::(\$?)([A-Z]{1,3})(\$?)(\d{1,4}))?")

# Both sheets grow by the same number of columns each month, but from
# different anchors, so a reference has to be shifted by the anchor of the
# sheet it POINTS AT, not the sheet the formula sits in.
SHEET_ANCHORS = {FIFO_SCH_SHEET: SCH_ANCHOR_COL, FIFO_PNL_SHEET: PNL_ANCHOR_COL}


def remap_formula_columns(formula, shift_from_col, delta, own_sheet=None):
    def bump(col, anchor):
        idx = column_index_from_string(col)
        return get_column_letter(idx + delta) if idx >= anchor else col

    def _sub(m):
        prefix, d1, c1, d2, r1, d3, c2, d4, r2 = m.groups()
        if prefix:
            sheet = prefix[:-1].strip("'")
            anchor = SHEET_ANCHORS.get(sheet)
            if anchor is None:
                return m.group(0)   # some other sheet: not ours to shift
        else:
            anchor = shift_from_col
        out = f"{prefix or ''}{d1}{bump(c1, anchor)}{d2}{r1}"
        if c2:
            out += f":{d3}{bump(c2, anchor)}{d4}{r2}"
        return out
    return SHEET_REF_RE.sub(_sub, formula)


def remap_formula_columns_bounded(formula, from_col, to_col_inclusive, delta):
    """Like remap_formula_columns, but only touches column letters strictly
    within [from_col, to_col_inclusive] — for rewriting a cross-block
    reference (e.g. a ratio-block cell pointing back at a main-block
    column) without re-touching that block's own already-shifted refs."""
    def _sub(m):
        dollar_col, col, dollar_row, row_s = m.groups()
        idx = column_index_from_string(col)
        if from_col <= idx <= to_col_inclusive:
            col = get_column_letter(idx + delta)
        return f"{dollar_col}{col}{dollar_row}{row_s}"
    return CELL_REF_RE.sub(_sub, formula)


def get_real_month_columns(ws, date_row, anchor_col):
    """Columns (in order, left to right = newest to oldest) that hold a
    parseable month-date label in `date_row`, starting at anchor_col."""
    cols = []
    c = anchor_col
    while True:
        v = ws.cell(row=date_row, column=c).value
        if v is None or str(v).strip() == "":
            break
        cols.append(c)
        c += 1
    return cols


def shift_columns_and_style(ws, anchor_col, insert_count, max_row, real_cols):
    """Insert `insert_count` fresh columns at anchor_col, shifting the real
    month columns right by that many; anything past the last real column
    is left untouched. Formulas in the shifted columns get their embedded
    column letters bumped by insert_count. Returns each row's captured
    pre-shift anchor-pair formulas/values (for building the new columns)."""
    last_real_col = max(real_cols) if real_cols else anchor_col - 1

    captured = {}
    for r in range(1, max_row + 1):
        captured[r] = {}
        for off in range(insert_count):
            cell = ws.cell(row=r, column=anchor_col + off)
            captured[r][off] = {"value": cell.value, "style": copy.copy(cell._style)}

    snapshot = {}
    for r in range(1, max_row + 1):
        for c in range(anchor_col, last_real_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                val = remap_formula_columns(val, anchor_col, insert_count, ws.title)
            snapshot[(r, c)] = {"value": val, "style": copy.copy(cell._style)}

    for r in range(1, max_row + 1):
        for c in range(anchor_col, last_real_col + 1):
            data = snapshot[(r, c)]
            dest = ws.cell(row=r, column=c + insert_count)
            dest.value = data["value"]
            dest._style = copy.copy(data["style"])

    for r in range(1, max_row + 1):
        for off in range(insert_count):
            new_cell = ws.cell(row=r, column=anchor_col + off)
            new_cell.value = None
            style_source = ws.cell(row=r, column=anchor_col + insert_count + off)
            new_cell._style = copy.copy(style_source._style)

    return captured


def classify_movement_formula(formula, row):
    """Returns (kind, formula_to_use_for_new_month)."""
    if formula is None:
        return "blank", None
    if not isinstance(formula, str) or not formula.startswith("="):
        return "value", formula
    exact_diff = f"=G{row}-I{row}"
    exact_mirror = f"=G{row}"
    if formula == exact_diff or formula == exact_mirror:
        return "template", formula
    m = DIFF_WITH_TAIL_RE.match(formula)
    if m and int(m.group(1)) == row:
        return "adjustment_dropped", exact_diff
    m = NUMERIC_TAIL_RE.match(formula)
    if m:   # e.g. "=-H36-26198033": a one-off number bolted on for one month
        return "adjustment_dropped", m.group(1)
    return "template", formula  # subtotal / cross-reference — copy verbatim


def classify_amount_formula(formula, prev_month, new_month):
    """Classifies a non-GL-sourced *amount* cell's old formula/value so we
    never silently repeat a stale one-off number as if it were live data:
      - ends in a fiscal-YTD "*N" multiplier matching last month's elapsed
        count (e.g. Audit Fees' "=600000/12*4") -> bump N for the new month.
      - contains a cell reference at all (a subtotal SUM, the Opening/
        Closing-stock rollforward's "-H31" style cross-reference, etc.) ->
        it's structural, copy it forward verbatim.
      - a formula with no cell references and no recognized multiplier
        pattern is pure one-off arithmetic typed for whichever month it was
        first entered (e.g. "=10786172+1492831") -> leave blank rather than
        repeat a number that has no reason to still be right.
      - a plain hardcoded number, same story -> leave blank.
    For the "left-blank" case, the old value/formula is still returned as
    a `suggestion` — not written to the cell (we don't trust it), but
    worth showing on the review page as a starting point to edit rather
    than leaving the box empty with no context at all.
    Returns (kind, value_or_formula_to_use, suggestion_or_None).
    """
    if formula is None:
        return "blank", None, None
    if not isinstance(formula, str) or not formula.startswith("="):
        return "left-blank", None, str(formula)  # plain hardcoded number — don't repeat it, but suggest it

    m = TRAILING_MULT_RE.search(formula)
    if m:
        old_mult = int(m.group(2))
        if old_mult == fiscal_month_number(prev_month):
            new_mult = fiscal_month_number(new_month)
            new_formula = formula[:m.start(2)] + str(new_mult) + formula[m.end(2):]
            return "fiscal-accrual", new_formula, new_formula

    if CELL_REF_RE.search(formula):
        return "template", formula, formula

    return "left-blank", None, formula


def process_fifo_sch(ws, totals, new_month, is_new_year_start):
    # Shift everything from the anchor to the sheet's last used column, not
    # just the dated month columns: the client keeps side notes to the right
    # of the months, and a month-only shift wrote straight over them.
    real_cols = list(range(SCH_ANCHOR_COL, max(ws.max_column, SCH_ANCHOR_COL) + 1))
    max_row = ws.max_row
    insert_count = 1 if is_new_year_start else 2

    captured = shift_columns_and_style(ws, SCH_ANCHOR_COL, insert_count, max_row, real_cols)
    old_amount_col = SCH_ANCHOR_COL + 1 + insert_count   # where last month's amount now sits

    amount_col = SCH_ANCHOR_COL if is_new_year_start else SCH_ANCHOR_COL + 1
    movement_col = None if is_new_year_start else SCH_ANCHOR_COL

    ws.cell(row=SCH_DATE_ROW, column=amount_col).value = fmt_sch_date(new_month)
    if movement_col:
        ws.cell(row=SCH_DATE_ROW, column=movement_col).value = fmt_sch_date(new_month)
    ws.cell(row=SCH_HEADER_ROW, column=amount_col).value = "Amt as on"
    if movement_col:
        ws.cell(row=SCH_HEADER_ROW, column=movement_col).value = "only"

    gl_rows_partial = []
    manual_rows = []
    gl_rows_manual = []
    suggestions = {}

    prev_month = _prev_month(new_month)

    def apply_manual_amount(r, desc):
        old_amount_off = 1 if 1 in captured[r] else 0
        old_amount = captured[r][old_amount_off]["value"]
        kind, new_value, suggestion = classify_amount_formula(old_amount, prev_month, new_month)
        if kind in ("template", "fiscal-accrual"):
            formula = _retarget_column(new_value, SCH_ANCHOR_COL + old_amount_off, amount_col)
            ws.cell(row=r, column=amount_col).value = formula
            if desc:
                manual_rows.append((r, desc, kind))
        else:
            ws.cell(row=r, column=amount_col).value = None
            if desc:
                manual_rows.append((r, desc, "left-blank"))
            if suggestion is not None:
                suggestions[r] = suggestion

    sign_flipped = []
    opening_carried = []

    for r in range(SCH_FIRST_DATA_ROW, max_row + 1):
        code_cell = ws.cell(row=r, column=SCH_GL_COL).value
        desc = ws.cell(row=r, column=SCH_DESC_COL).value
        codes = parse_gl_codes(code_cell)

        old_amt = ws.cell(row=r, column=old_amount_col).value
        is_opening = bool(codes) and all(c.startswith("X1") for c in codes)
        is_closing = bool(codes) and all(c.startswith("R3") for c in codes)

        if is_opening:
            # Opening stock is fixed for the whole fiscal year. Sheet1's
            # X1xxx balances are book figures that differ from the costed
            # opening stock kept here, and (having been overwritten by them)
            # this used to pick up the closing-stock figure by mistake.
            has_refs = isinstance(old_amt, str) and old_amt.startswith("=") and CELL_REF_RE.search(old_amt)
            if is_new_year_start or old_amt is None or has_refs:
                apply_manual_amount(r, desc)
            else:
                ws.cell(row=r, column=amount_col).value = old_amt
                opening_carried.append((r, code_cell, desc))
        elif codes:
            value, any_missing, all_missing = sum_for_codes(codes, totals)
            if not all_missing and is_closing:
                # this sheet keeps closing stock as a deduction (negative);
                # Sheet1 shows it as a positive balance. Follow whatever sign
                # the previous month used.
                was_negative = (isinstance(old_amt, (int, float)) and old_amt < 0) or \
                    (isinstance(old_amt, str) and old_amt.startswith("=-"))
                if was_negative:
                    sign_flipped.append((r, code_cell, desc))
                value = _formula_or_value(codes, totals, value, negate=was_negative)
            if all_missing:
                # none of these codes are tracked by Sheet1 at all (e.g. the
                # FIFO stock rollforward, Audit Fees) — treat as manual
                gl_rows_manual.append((r, code_cell, desc))
                apply_manual_amount(r, desc)
            else:
                if not is_closing:
                    value = _formula_or_value(codes, totals, value)
                ws.cell(row=r, column=amount_col).value = value
                if any_missing:
                    gl_rows_partial.append((r, code_cell, desc))
        else:
            apply_manual_amount(r, desc)

        if movement_col is not None:
            old_mvt = captured[r].get(0, {}).get("value")
            kind, new_formula = classify_movement_formula(old_mvt, r)
            if kind in ("template", "adjustment_dropped"):
                ws.cell(row=r, column=movement_col).value = new_formula
                if kind == "adjustment_dropped":
                    manual_rows.append((r, desc, "adjustment_dropped"))
            elif kind == "value" and old_mvt is not None:
                ws.cell(row=r, column=movement_col).value = old_mvt

    return {
        "gl_rows_partial": gl_rows_partial,
        "gl_rows_manual": gl_rows_manual,
        "manual_rows": manual_rows,
        "amount_col": amount_col,
        "suggestions": suggestions,
        "sign_flipped": sign_flipped,
        "opening_carried": opening_carried,
    }


def process_fifo_pnl(ws, totals, new_month, is_new_year_start, gross_sales, sch_amount_col):
    real_cols = get_real_month_columns(ws, PNL_HEADER_ROW, PNL_ANCHOR_COL)
    max_row = ws.max_row
    insert_count = 1 if is_new_year_start else 2
    prev_month = _prev_month(new_month)

    # The main block and the '% of Net Sales' ratio block (one spacer
    # column apart) both grow by `insert_count` every month, but the ratio
    # block's total displacement is 2x that: once because the main block's
    # own growth pushes everything after it over, and again for the ratio
    # block's own new-month insertion. Treating both as a plain
    # `insert_count` shift (as an earlier version of this code did) made
    # the main block's growth land exactly on the ratio block's freshly
    # vacated column, silently overwriting it — this is computed fresh
    # every run (never a fixed offset) so it keeps working as both blocks
    # keep growing month after month.
    last_main_col_before_shift = max(real_cols) if real_cols else PNL_ANCHOR_COL - 1
    ratio_anchor = last_main_col_before_shift + 2  # current ratio block position
    real_ratio_cols = get_real_month_columns(ws, PNL_RATIO_DATE_ROW, ratio_anchor)
    total_ratio_shift = insert_count * 2

    # Snapshot the ratio block's OLD content BEFORE step 1 touches anything —
    # step 1's own destination range reaches into the ratio block's current
    # front column(s) (that's the whole point: the main block's growth
    # pushes the ratio block over), so reading it after step 1 would already
    # see step 1's overwrite instead of the original ratio data.
    ratio_snapshot = {}
    if real_ratio_cols:
        last_ratio_col = max(real_ratio_cols)
        for r in range(1, max_row + 1):
            for c in range(ratio_anchor, last_ratio_col + 1):
                cell = ws.cell(row=r, column=c)
                val = cell.value
                if isinstance(val, str) and val.startswith('='):
                    val = remap_formula_columns(val, ratio_anchor, total_ratio_shift, ws.title)
                ratio_snapshot[(r, c)] = {"value": val, "style": copy.copy(cell._style)}

    # Step 1: shift the main block AND its one spacer column by insert_count.
    main_and_spacer_cols = list(real_cols) + ([last_main_col_before_shift + 1] if real_ratio_cols else [])
    captured = shift_columns_and_style(ws, PNL_ANCHOR_COL, insert_count, max_row, main_and_spacer_cols)

    # Step 2: relocate the ratio block's OLD content by 2x insert_count
    # (its own displacement, doubled), while creating fresh blank columns
    # only at its new front — NOT at [ratio_anchor, ratio_anchor+insert_count)
    # which step 1 already correctly filled with the main block's own
    # (shifted) rightmost column and spacer.
    if real_ratio_cols:
        for r in range(1, max_row + 1):
            for c in range(ratio_anchor, last_ratio_col + 1):
                data = ratio_snapshot[(r, c)]
                dest = ws.cell(row=r, column=c + total_ratio_shift)
                dest.value = data["value"]
                dest._style = copy.copy(data["style"])

        new_ratio_front = ratio_anchor + insert_count
        for r in range(1, max_row + 1):
            for off in range(insert_count):
                new_cell = ws.cell(row=r, column=new_ratio_front + off)
                new_cell.value = None
                style_source = ws.cell(row=r, column=new_ratio_front + insert_count + off)
                new_cell._style = copy.copy(style_source._style)

        # the ratio block's cells also reference the main block's columns
        # directly (e.g. "=+F10/F$8%"), which needs the main block's own
        # (smaller) shift amount applied, not the ratio block's doubled one
        last_ratio_col_after_shift = last_ratio_col + total_ratio_shift
        for r in range(1, max_row + 1):
            for c in range(new_ratio_front, last_ratio_col_after_shift + 1):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.value = remap_formula_columns_bounded(
                        cell.value, PNL_ANCHOR_COL, last_main_col_before_shift, insert_count)

    amount_col = PNL_ANCHOR_COL if is_new_year_start else PNL_ANCHOR_COL + 1
    movement_col = None if is_new_year_start else PNL_ANCHOR_COL

    for tr in range(1, PNL_HEADER_ROW):
        for tc in range(1, 8):
            v = ws.cell(row=tr, column=tc).value
            if isinstance(v, str) and TITLE_DATE_RE.search(v):
                ws.cell(row=tr, column=tc).value = TITLE_DATE_RE.sub(
                    lambda m: m.group(1) + fmt_sch_date(new_month), v)

    month_str = new_month.strftime("%B-%Y")
    ws.cell(row=PNL_HEADER_ROW, column=amount_col).value = f"As on                    {month_str}"
    if movement_col:
        ws.cell(row=PNL_HEADER_ROW, column=movement_col).value = f"Only                      {month_str}"

    # Row 6 (Gross Sales & Other Incomes) and row 8 (Net Sales & Other
    # Incomes) used to carry one-off hardcoded arithmetic that only
    # happened to be right for whichever month it was typed for — copying
    # it forward silently repeated a stale number every month after.
    # Row 6 now comes from the client typing it directly into Sheet1's
    # "Gross sales" label row; row 8 is pulled live from FIFO-sch's own
    # Net Sales total (row 13), so it's always correct without any manual
    # re-entry. Row 7 (Less: GST & Others) needs no special handling — its
    # existing "=F6-F8" style formula already just works once 6 and 8 are
    # right, and gets carried forward by the generic logic below like any
    # other self-referential formula.
    gross_sales_status = "missing"
    if gross_sales is not None:
        ws.cell(row=PNL_GROSS_SALES_ROW, column=amount_col).value = GROSS_FORMULA or gross_sales
        gross_sales_status = "provided"
    else:
        ws.cell(row=PNL_GROSS_SALES_ROW, column=amount_col).value = None

    sch_amount_letter = get_column_letter(sch_amount_col)
    ws.cell(row=PNL_NET_SALES_ROW, column=amount_col).value = \
        f"=+'{FIFO_SCH_SHEET}'!{sch_amount_letter}{SCH_NET_SALES_ROW}"

    gl_rows = []
    for r in range(PNL_FIRST_DATA_ROW, max_row + 1):
        if r in (PNL_GROSS_SALES_ROW, PNL_NET_SALES_ROW):
            skip_amount = True
        else:
            skip_amount = False
        code_cell = ws.cell(row=r, column=PNL_GL_COL).value
        codes = parse_gl_codes(code_cell)
        desc = ws.cell(row=r, column=PNL_DESC_COL).value

        if not skip_amount:
            if codes:
                value, any_missing, all_missing = sum_for_codes(codes, totals)
                if all_missing:
                    codes = []  # fall through to manual handling below
                else:
                    ws.cell(row=r, column=amount_col).value = _formula_or_value(codes, totals, value)
                    gl_rows.append((r, code_cell, desc))
            if not codes:
                old_amount_off = 1 if 1 in captured[r] else 0
                old_amount = captured[r][old_amount_off]["value"]
                kind, new_value, _suggestion = classify_amount_formula(old_amount, prev_month, new_month)
                if kind in ("template", "fiscal-accrual"):
                    ws.cell(row=r, column=amount_col).value = _retarget_column(
                        new_value, PNL_ANCHOR_COL + old_amount_off, amount_col)
                else:
                    ws.cell(row=r, column=amount_col).value = None

        if movement_col is not None:
            old_mvt = captured[r].get(0, {}).get("value")
            if isinstance(old_mvt, str) and old_mvt.startswith("="):
                ws.cell(row=r, column=movement_col).value = old_mvt
            elif old_mvt is not None:
                ws.cell(row=r, column=movement_col).value = old_mvt

    if real_ratio_cols:
        process_pnl_ratio_block(ws, new_month, is_new_year_start, amount_col, movement_col,
                                 max_row, new_ratio_front)

    return {"gl_rows": gl_rows, "gross_sales_status": gross_sales_status}


def process_pnl_ratio_block(ws, new_month, is_new_year_start, main_amount_col, main_movement_col,
                             max_row, ratio_anchor):
    """The '% of Net Sales' block mirrors every main column. Its columns
    were already shifted/relocated by process_fifo_pnl (which computes
    where its fresh columns land, since that depends on the main block's
    current size, not a fixed offset) — this call only populates them."""
    insert_count = 1 if is_new_year_start else 2

    new_month_dt = datetime.datetime(new_month.year, new_month.month, new_month.day)
    new_ratio_amount_col = ratio_anchor + (0 if is_new_year_start else 1)
    new_ratio_movement_col = None if is_new_year_start else ratio_anchor

    ws.cell(row=PNL_RATIO_DATE_ROW, column=new_ratio_amount_col).value = new_month_dt
    if new_ratio_movement_col:
        ws.cell(row=PNL_RATIO_DATE_ROW, column=new_ratio_movement_col).value = new_month_dt
        ws.cell(row=6, column=new_ratio_movement_col).value = "only"

    main_amount_letter = get_column_letter(main_amount_col)
    for r in range(1, max_row + 1):
        # only rows following the standard "=+<col><row>/<col>$8%" shape get
        # a fresh one — a handful of one-off rows (e.g. a delta between two
        # specific historical ratio columns) don't fit that mold and are
        # left alone rather than guessed at
        row_formula_re = re.compile(rf'^=\+?[A-Z]{{1,3}}{r}/[A-Z]{{1,3}}\$8%$')
        has_standard_ratio = any(
            isinstance(ws.cell(row=r, column=c).value, str) and row_formula_re.match(ws.cell(row=r, column=c).value)
            for c in range(ratio_anchor + insert_count, ws.max_column + 1)
        )
        if not has_standard_ratio:
            continue
        ws.cell(row=r, column=new_ratio_amount_col).value = f"=+{main_amount_letter}{r}/{main_amount_letter}$8%"
        if new_ratio_movement_col and main_movement_col:
            main_movement_letter = get_column_letter(main_movement_col)
            ws.cell(row=r, column=new_ratio_movement_col).value = \
                f"=+{main_movement_letter}{r}/{main_movement_letter}$8%"


def get_current_month(wb):
    ws = wb[FIFO_SCH_SHEET]
    d = ws.cell(row=SCH_DATE_ROW, column=SCH_ANCHOR_COL).value
    if d is None:
        return None
    return parse_sch_date(d)


def _year_month(d):
    return (d.year, d.month)


def apply_new_month(wb, sheet1_ws):
    totals, new_month, gross_sales = read_sheet1_totals(sheet1_ws)
    unmapped = find_unmapped_codes(wb, totals, read_sheet1_descriptions(sheet1_ws))

    current_month = get_current_month(wb)
    if current_month is not None and _year_month(new_month) <= _year_month(current_month):
        if _year_month(new_month) == _year_month(current_month):
            raise NotANewMonthError(
                f"The uploaded Sheet1 is for {new_month.strftime('%B %Y')}, which is "
                f"already the latest month in the ACPLACT workbook. Looks like a "
                f"duplicate upload — nothing was changed."
            )
        raise NotANewMonthError(
            f"The uploaded Sheet1 is for {new_month.strftime('%B %Y')}, but the "
            f"ACPLACT workbook already has {current_month.strftime('%B %Y')} as its "
            f"newest month. Upload a month after that instead."
        )

    is_new_year_start = is_fiscal_year_start(new_month)

    sch_ws = wb[FIFO_SCH_SHEET]
    pnl_ws = wb[FIFO_PNL_SHEET]

    sch_report = process_fifo_sch(sch_ws, totals, new_month, is_new_year_start)
    pnl_report = process_fifo_pnl(pnl_ws, totals, new_month, is_new_year_start,
                                   gross_sales, sch_report["amount_col"])

    return {
        "new_month": new_month.isoformat(),
        "gl_rows_partial": sch_report["gl_rows_partial"],
        "gl_rows_manual": sch_report["gl_rows_manual"],
        "manual_rows": sch_report["manual_rows"],
        "pnl_gl_rows": pnl_report["gl_rows"],
        "gross_sales_status": pnl_report["gross_sales_status"],
        "suggestions": sch_report["suggestions"],
        "sign_flipped": sch_report["sign_flipped"],
        "opening_carried": sch_report["opening_carried"],
        "unmapped_codes": unmapped,
    }


class NotTheLatestMonthError(Exception):
    pass


def _current_amount_col(ws, header_row, anchor_col):
    """Whether the current newest month is a single (amount-only, fiscal
    year start) or paired (movement+amount) column, by reading its own
    header label: the anchor column holds 'only'/'Only ...' when paired,
    or the amount label directly when it's a single (fiscal-year-start)
    column."""
    label = str(ws.cell(row=header_row, column=anchor_col).value or "")
    is_paired = label.strip().lower().startswith("only")
    return anchor_col + 1 if is_paired else anchor_col


def replace_latest_month(wb, sheet1_ws):
    """For corrected/final data arriving after a month's already been
    merged in. Re-populates the existing latest-month amount column(s)
    from the new Sheet1, in place — no shifting, no row insertion (adding
    a brand-new GL row isn't supported for ACPLACT; only pre-existing
    rows get updated). Movement columns and non-GL rows are untouched,
    since they already carry whatever was right for this same month."""
    totals, new_month, gross_sales = read_sheet1_totals(sheet1_ws)

    current_month = get_current_month(wb)
    if current_month is None:
        raise NotTheLatestMonthError(
            "The ACPLACT workbook doesn't have a latest month yet to replace. "
            "Use 'New month' instead."
        )
    if _year_month(new_month) != _year_month(current_month):
        raise NotTheLatestMonthError(
            f"This file is for {new_month.strftime('%B %Y')}, but the ACPLACT "
            f"workbook's latest month is {current_month.strftime('%B %Y')}. "
            f"'Replace latest month' only works when the uploaded file is for "
            f"that same month — use 'New month' if this is meant to move forward."
        )

    sch_ws = wb[FIFO_SCH_SHEET]
    pnl_ws = wb[FIFO_PNL_SHEET]

    sch_amount_col = _current_amount_col(sch_ws, SCH_HEADER_ROW, SCH_ANCHOR_COL)
    pnl_amount_col = _current_amount_col(pnl_ws, PNL_HEADER_ROW, PNL_ANCHOR_COL)

    gl_rows_partial = []
    gl_rows_updated = []
    for r in range(SCH_FIRST_DATA_ROW, sch_ws.max_row + 1):
        codes = parse_gl_codes(sch_ws.cell(row=r, column=SCH_GL_COL).value)
        if not codes:
            continue
        if all(c.startswith("X1") for c in codes):
            continue  # opening stock is carried forward, never taken from Sheet1
        value, any_missing, all_missing = sum_for_codes(codes, totals)
        if all_missing:
            continue  # permanently-manual row — leave whatever is already there
        if all(c.startswith("R3") for c in codes):
            here = sch_ws.cell(row=r, column=sch_amount_col).value
            before = sch_ws.cell(row=r, column=sch_amount_col + 2).value
            if any((isinstance(v, (int, float)) and v < 0) or (isinstance(v, str) and v.startswith("=-"))
                   for v in (here, before)):
                value = -value
                negate = True
            else:
                negate = False
        else:
            negate = False
        sch_ws.cell(row=r, column=sch_amount_col).value = _formula_or_value(
            codes, totals, abs(value) if negate else value, negate)
        gl_rows_updated.append((r, sch_ws.cell(row=r, column=SCH_GL_COL).value,
                                 sch_ws.cell(row=r, column=SCH_DESC_COL).value))
        if any_missing:
            gl_rows_partial.append((r, sch_ws.cell(row=r, column=SCH_GL_COL).value,
                                     sch_ws.cell(row=r, column=SCH_DESC_COL).value))

    pnl_gl_rows_updated = []
    for r in range(PNL_FIRST_DATA_ROW, pnl_ws.max_row + 1):
        if r in (PNL_GROSS_SALES_ROW, PNL_NET_SALES_ROW):
            continue
        codes = parse_gl_codes(pnl_ws.cell(row=r, column=PNL_GL_COL).value)
        if not codes:
            continue
        value, any_missing, all_missing = sum_for_codes(codes, totals)
        if all_missing:
            continue
        pnl_ws.cell(row=r, column=pnl_amount_col).value = _formula_or_value(codes, totals, value)
        pnl_gl_rows_updated.append((r, pnl_ws.cell(row=r, column=PNL_GL_COL).value,
                                     pnl_ws.cell(row=r, column=PNL_DESC_COL).value))

    gross_sales_status = "missing"
    if gross_sales is not None:
        pnl_ws.cell(row=PNL_GROSS_SALES_ROW, column=pnl_amount_col).value = GROSS_FORMULA or gross_sales
        gross_sales_status = "provided"
    sch_amount_letter = get_column_letter(sch_amount_col)
    pnl_ws.cell(row=PNL_NET_SALES_ROW, column=pnl_amount_col).value = \
        f"=+'{FIFO_SCH_SHEET}'!{sch_amount_letter}{SCH_NET_SALES_ROW}"

    return {
        "new_month": new_month.isoformat(),
        "gl_rows_partial": gl_rows_partial,
        "gl_rows_manual": [],
        "manual_rows": [],
        "pnl_gl_rows": pnl_gl_rows_updated,
        "gl_rows_updated": gl_rows_updated,
        "gross_sales_status": gross_sales_status,
        "suggestions": {},
    }
