import os
import re
import shutil
import datetime
import traceback
from flask import Flask, request, render_template, send_file, redirect, url_for, flash, session

import openpyxl
from openpyxl.cell.cell import MergedCell

from sheet1_formulas import read_total_formulas
import transform_type1 as t1
import transform_type2 as t2
import auth
import json
import io
from history import HistoryLog, earlier_month_changes

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(BASE_DIR)
PNL_PATH = os.path.join(PROJECT_DIR, "P&L STATEMENT.xlsx")
ACPLACT_PATH = os.path.join(PROJECT_DIR, "acplact.xlsx")

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
HISTORY_DIR = os.path.join(BASE_DIR, "history")
PENDING_DIR = os.path.join(BASE_DIR, "pending")
PENDING_PNL_PATH = os.path.join(PENDING_DIR, "pending_pnl.xlsx")
PENDING_ACPLACT_PATH = os.path.join(PENDING_DIR, "pending_acplact.xlsx")

HISTORY = HistoryLog(HISTORY_DIR)
META_PATH = os.path.join(PENDING_DIR, "meta.json")

for d in (UPLOAD_DIR, OUTPUT_DIR, HISTORY_DIR, PENDING_DIR):
    os.makedirs(d, exist_ok=True)

app = Flask(__name__)
app.secret_key = "astra-chemtech-pnl-tool"  # signs the login session cookie; kept stable across restarts on purpose

PUBLIC_ROUTES = {"login"}  # accessible without being logged in


@app.before_request
def require_login():
    if request.endpoint is None:
        return  # e.g. a 404 on a non-matching path — let it 404 normally
    if request.endpoint in PUBLIC_ROUTES or request.endpoint == "static":
        return
    if not session.get("logged_in"):
        return redirect(url_for("login", next=request.path))


def _safe_next_path(candidate):
    """Only ever redirect to a path within this app — a `next` value from
    a query string or form field is user-controlled, so an unrestricted
    redirect there would be an open-redirect vector."""
    if candidate and candidate.startswith("/") and not candidate.startswith("//"):
        return candidate
    return url_for("index")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if auth.verify_login(username, password):
            session["logged_in"] = True
            session["username"] = username
            return redirect(_safe_next_path(request.form.get("next")))
        flash("Incorrect username or password.", "error")
    next_path = request.args.get("next", "")
    return render_template("login.html", next_path=next_path)


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/change-password", methods=["GET", "POST"])
def change_password():
    if request.method == "POST":
        old_password = request.form.get("old_password", "")
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")
        if new_password != confirm_password:
            flash("New password and confirmation don't match.", "error")
        else:
            ok, error = auth.change_password(old_password, new_password)
            if ok:
                flash("Password changed.", "ok")
                return redirect(url_for("index"))
            flash(error, "error")
    return render_template("change_password.html")

TRANSFORMATIONS = {
    "pnl": {
        "label": "PNL",
        "description": "The monthly P&L STATEMENT.xlsx merge — matches Sheet1 by GL Code, adds the new month as a column.",
    },
    "acplact": {
        "label": "ACPLACT",
        "description": "The FIFO-sch / FIFO-PNL merge in acplact.xlsx — GL Code (incl. merged multi-code rows) summed from the same Sheet1.",
    },
}

MODES = {
    "new_month": {
        "label": "New month",
        "description": "Add this month as a brand-new column, shifting older months right.",
    },
    "replace_latest": {
        "label": "Replace latest month",
        "description": "Overwrite the current latest month's data with corrected/final figures — use this when more data arrived after you already processed that month.",
    },
}

SOURCE_LABELS = {
    "matched": "From Sheet1 (matched by GL Code)",
    "zeroed": "Set to 0 — not found in this Sheet1",
    "new_gl": "NEW row inserted — check it belongs here or add its code to an existing row",
    "template": "Auto-filled — subtotal/movement formula",
    "fiscal-accrual": "Auto-filled — fiscal-YTD formula, multiplier bumped",
    "flat-value": "Auto-filled — unchanged from prior months",
    "left-blank": "Left blank — needs manual entry",
    "manual-formula-left-blank": "Left blank — old formula had no safe pattern",
    "no-formula-found": "Left blank — no formula to carry forward",
    "adjustment_dropped": "One-off adjustment dropped — review if one applies this month",
    "gross-sales-matched": "From Sheet1's Gross Sales row",
    "adopted": "GL Code attached to your existing row (no duplicate added)",
    "partial": "Some of this row's GL codes are missing from Sheet1",
}


def cell_display_value(v):
    if v is None:
        return ""
    return str(v)


def parse_edited_value(text):
    """Turn a review-page text field back into a cell value: blank -> None,
    '=...' -> kept as a formula string, else numeric if possible, else raw text."""
    text = text.strip()
    if text == "":
        return None
    if text.startswith("="):
        return t2.fix_comma_numbers(text)
    try:
        # values are shown with thousand separators (1,234,567.89); Excel
        # formulas can't contain those, so strip them before parsing
        f = float(text.replace(",", "").replace(" ", ""))
        return int(f) if f.is_integer() else f
    except ValueError:
        return text


def bad_inputs(form, prefixes, current=None):
    """Review-page boxes holding something that is neither blank, a number
    (commas allowed) nor a formula — e.g. "abc" or "10 lakh". Written into
    an amount cell, text turns every total that uses it into #VALUE!."""
    bad = []
    for key, text in form.items():
        if not key.startswith(prefixes):
            continue
        if current is not None and current(key) == text.strip():
            continue   # left as it was — only what the user changed is checked
        v = parse_edited_value(text)
        if isinstance(v, str) and not v.startswith("="):
            bad.append(f"'{text.strip()}' typed in {key.replace('cell_acplact_', '').replace('cell_pnl_', 'row ').replace('nf_value_', 'new GL code #')}"
                       f" is not a number or a formula (starting with =). Nothing was saved — go back and correct it.")
        elif isinstance(v, str) and t2._stray_comma_positions(v):
            bad.append(f"'{text.strip()}' typed in {key} has a comma Excel can't read. Nothing was saved.")
    return bad


def find_previous_sheet1_totals(prev_month, exclude_path=None):
    """Last month's Sheet1 {code: Total Balance}, from the newest earlier
    upload whose title is for prev_month (every upload is kept in
    UPLOAD_DIR). Used to spot rows the client adjusted by hand last month.
    None if no such upload exists (e.g. first run on a new machine)."""
    want = (prev_month.year, prev_month.month)
    try:
        names = sorted((f for f in os.listdir(UPLOAD_DIR) if f.endswith(".xlsx")), reverse=True)
    except OSError:
        return None
    for name in names:
        path = os.path.join(UPLOAD_DIR, name)
        if exclude_path and os.path.abspath(path) == os.path.abspath(exclude_path):
            continue
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.worksheets[0]
            title = next((v for row in ws.iter_rows(min_row=4, max_row=4, max_col=3, values_only=True)
                          for v in row if v), "")
            m = re.search(r'as on\s+([A-Za-z]+)[-\s]+(\d{4})', str(title))
            if not m:
                continue
            month = datetime.datetime.strptime(m.group(1)[:3], "%b").month
            if (int(m.group(2)), month) != want:
                continue
            wb.close()
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.worksheets[0]
            return t2.read_sheet1_totals(ws)[0]
        except Exception:
            continue
    return None


def copy_sheet_values(src_ws, dest_ws):
    """Copy values (not formulas' side-effects) from an uploaded sheet into
    the persisted workbook's Sheet1, skipping merged-cell followers."""
    for row in src_ws.iter_rows():
        for cell in row:
            target = dest_ws.cell(row=cell.row, column=cell.column)
            if isinstance(target, MergedCell):
                continue
            target.value = cell.value


NEEDS_ATTENTION_TAGS = {
    "zeroed", "left-blank", "manual-formula-left-blank", "no-formula-found", "adjustment_dropped",
    "new_gl", "partial",
}


def build_pnl_review_rows(ws, report):
    """One row per GL-linked or manual/subtotal row touched by the merge,
    editable value/formula for the new/latest month's column L."""
    t1.detect_layout(ws)
    new_codes = {code for code, _ in report.get("new_gl_rows_inserted", [])}
    adopted_codes = {code for code, _, _ in report.get("adopted_rows", [])}
    partial_rows = {r for r, _, _ in report.get("gl_rows_partial", [])}
    manual_tags = {r: tag for r, _, tag in report.get("manual_rows", [])}
    suggestions = report.get("suggestions", {})

    rows = []
    for r in range(1, ws.max_row + 1):
        code = ws.cell(row=r, column=t1.GL_COL).value
        desc = ws.cell(row=r, column=t1.DESC_COL).value
        codes = t1.parse_codes(code)
        if r in manual_tags:
            source = manual_tags[r]
        elif code and r > t1.HEADER_ROW:
            if any(c in new_codes for c in codes):
                source = "new_gl"
            elif any(c in adopted_codes for c in codes):
                source = "adopted"
            elif r in partial_rows:
                source = "partial"
            else:
                source = "matched"
        else:
            continue

        cell = ws.cell(row=r, column=t1.TOTAL_COL)
        value = cell_display_value(cell.value)
        if value == "" and r in suggestions:
            value = suggestions[r]
        rows.append({
            "row": r,
            "code": code or "",
            "desc": desc or "",
            "value": value,
            "source": SOURCE_LABELS.get(source, source),
            "tag": source,
            "needs_attention": source in NEEDS_ATTENTION_TAGS,
        })
    return rows


def build_acplact_editable_rows(wb, report):
    """Only the fields ACPLACT can never get from Sheet1 at all — GL codes
    with no match anywhere in the trial balance, non-GL rows left blank
    with no safe pattern, and Gross Sales (which the client types directly
    into Sheet1's label row). Calculated rows (subtotals, GST, the
    Opening/Closing-stock rollforward) are never made editable here."""
    sch_ws = wb[t2.FIFO_SCH_SHEET]
    pnl_ws = wb[t2.FIFO_PNL_SHEET]
    sch_amount_col = t2._current_amount_col(sch_ws, t2.SCH_HEADER_ROW, t2.SCH_ANCHOR_COL)
    pnl_amount_col = t2._current_amount_col(pnl_ws, t2.PNL_HEADER_ROW, t2.PNL_ANCHOR_COL)

    manual_row_nums = {r for r, _code, _desc in report.get("gl_rows_manual", [])}
    left_blank_rows = {r for r, _desc, tag in report.get("manual_rows", []) if tag == "left-blank"}
    editable_sch_rows = manual_row_nums | left_blank_rows
    suggestions = report.get("suggestions", {})

    rows = []
    for r in sorted(editable_sch_rows):
        cell = sch_ws.cell(row=r, column=sch_amount_col)
        # the cell itself is deliberately left blank (we don't trust the old
        # formula enough to auto-apply it) — but show it as a starting point
        # to edit rather than an empty box with no context
        display_value = cell_display_value(cell.value)
        if display_value == "" and r in suggestions:
            display_value = str(suggestions[r])
        rows.append({
            "sheet": "sch",
            "row": r,
            "code": sch_ws.cell(row=r, column=t2.SCH_GL_COL).value or "",
            "desc": sch_ws.cell(row=r, column=t2.SCH_DESC_COL).value or "",
            "value": display_value,
            "needs_attention": True,
        })

    # rows the client changed by hand LAST month (figure differs from last
    # month's Sheet1): this month got the raw ledger figure — show last
    # month's formula so the same kind of adjustment can be re-applied
    for r, code, desc, used, ledger, old_formula in report.get("hand_adjusted", []):
        if r in editable_sch_rows:
            continue
        cell = sch_ws.cell(row=r, column=sch_amount_col)
        rows.append({
            "sheet": "sch", "row": r, "code": code or "", "desc": desc,
            "value": cell_display_value(cell.value),
            "needs_attention": True,
            "note": (f"Adjusted by hand last month: ledger said {ledger:,.2f}, the sheet used "
                     f"{used:,.2f} ({old_formula}). This month has the raw ledger figure — "
                     f"apply the same kind of adjustment if it still holds."),
        })

    # P&L lines that had a figure last month but have no source this month
    # (Wadi sales, the foreign-travel memo line, ...)
    for r, desc, suggestion, is_wadi in report.get("pnl_left_blank", []):
        rows.append({
            "sheet": "pnl", "row": r, "code": "", "desc": desc,
            "value": "",
            "needs_attention": True,
            "note": ((f"Wadi sales as on this month — Gross Sales and Net Sales add it automatically. "
                      if is_wadi else "Had a figure last month; enter this month's. ")
                     + (f"Last month: {suggestion}" if suggestion else "")),
        })

    gross_cell = pnl_ws.cell(row=t2.PNL_GROSS_SALES_ROW, column=pnl_amount_col)
    rows.append({
        "sheet": "pnl",
        "row": t2.PNL_GROSS_SALES_ROW,
        "code": "",
        "desc": pnl_ws.cell(row=t2.PNL_GROSS_SALES_ROW, column=t2.PNL_DESC_COL).value or "Gross Sales",
        "value": cell_display_value(gross_cell.value),
        "needs_attention": True,
    })
    return rows


def get_pnl_current_month():
    if not os.path.isfile(PNL_PATH):
        return None
    try:
        wb = openpyxl.load_workbook(PNL_PATH, data_only=True)
        d = t1.get_current_pnl_month(wb["P&L"])
        return d.strftime("%B %Y") if d else None
    except Exception:
        return None


class MasterDamagedError(Exception):
    pass


class UploadNotCalculatedError(Exception):
    pass


def file_fingerprint(path):
    """mtime + size: enough to notice the master was swapped or changed
    between the upload (review page) and Confirm."""
    try:
        st = os.stat(path)
        return [st.st_mtime, st.st_size]
    except OSError:
        return None


def acplact_master_status():
    """What the tool will actually read as the ACPLACT master, shown on the
    home page so nobody has to guess: full path, its latest month, when it
    was last saved, whether it's healthy, and any OTHER ACPLACT workbook
    lying in the project or app folder that's newer than it — the usual
    sign a replacement was copied to the wrong place or got a double
    extension (Windows hides ".xlsx", so "acplact.xlsx" typed as a new name
    becomes "acplact.xlsx.xlsx")."""
    st = {"path": ACPLACT_PATH, "folder": PROJECT_DIR, "exists": os.path.isfile(ACPLACT_PATH),
          "month": None, "saved": None, "errors": [], "lookalikes": []}
    if st["exists"]:
        st["saved"] = datetime.datetime.fromtimestamp(
            os.path.getmtime(ACPLACT_PATH)).strftime("%d %b %Y, %H:%M")
        try:
            wb = openpyxl.load_workbook(ACPLACT_PATH)
            d = t2.get_current_month(wb)
            st["month"] = d.strftime("%B %Y") if d else None
            t2.repair_comma_formulas(wb)
            st["errors"] = t2.blocking_problems(t2.validate_workbook(wb))
        except Exception as e:
            st["errors"] = [f"Could not read it: {e}"]
    for folder in (PROJECT_DIR, BASE_DIR):
        try:
            names = os.listdir(folder)
        except OSError:
            continue
        for name in names:
            path = os.path.join(folder, name)
            if (not name.lower().endswith(".xlsx") or name.startswith("~$")
                    or os.path.abspath(path) == os.path.abspath(ACPLACT_PATH) or not os.path.isfile(path)):
                continue
            # every other ACPLACT workbook, whatever its date: a copied file
            # keeps its ORIGINAL modified date on Windows, so "newer than the
            # master" would hide exactly the replacement someone just dropped in
            month = None
            try:
                wbo = openpyxl.load_workbook(path, read_only=True)
                ok = t2.is_acplact_workbook(wbo.sheetnames)
                if ok:
                    try:
                        d = t2.parse_sch_date(wbo[t2.FIFO_SCH_SHEET].cell(
                            row=t2.SCH_DATE_ROW, column=t2.SCH_ANCHOR_COL).value)
                        month = d.strftime("%B %Y")
                    except Exception:
                        pass
                wbo.close()
            except Exception:
                ok = False
            if ok:
                st["lookalikes"].append({
                    "name": name, "folder": folder, "month": month,
                    "double_ext": name.lower().endswith(".xlsx.xlsx"),
                    "when": datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d %b %Y, %H:%M")})
    st["lookalikes"].sort(key=lambda x: (not x["double_ext"], x["name"]))
    return st


def get_acplact_current_month():
    if not os.path.isfile(ACPLACT_PATH):
        return None
    try:
        wb = openpyxl.load_workbook(ACPLACT_PATH, data_only=True)
        d = t2.get_current_month(wb)
        return d.strftime("%B %Y") if d else None
    except Exception:
        return None


BACKUP_TS_RE = re.compile(r'_(\d{8}_\d{6})\.xlsx$')


def latest_backup_batch():
    """PNL and ACPLACT backups made in the same /confirm call share the
    same timestamp suffix (e.g. '20260917_193900') — group by that so
    Undo can restore every file from the most recent run together,
    instead of only whichever single file happens to be newest by mtime."""
    names = [f for f in os.listdir(HISTORY_DIR) if f.endswith(".xlsx")]
    with_ts = [(BACKUP_TS_RE.search(f).group(1), f) for f in names if BACKUP_TS_RE.search(f)]
    if not with_ts:
        return []
    latest_ts = max(ts for ts, _ in with_ts)
    return [os.path.join(HISTORY_DIR, f) for ts, f in with_ts if ts == latest_ts]


def recent_outputs():
    """Newest finished file of each kind, so the client can re-download last
    time's result on a fresh visit."""
    found = []
    for label, prefix in (("PNL", "P&L STATEMENT ("), ("ACPLACT", "acplact (")):
        files = [f for f in os.listdir(OUTPUT_DIR) if f.startswith(prefix) and f.endswith(".xlsx")]
        if not files:
            continue
        newest = max(files, key=lambda f: os.path.getmtime(os.path.join(OUTPUT_DIR, f)))
        when = datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(OUTPUT_DIR, newest)))
        found.append({"label": label, "filename": newest, "when": when.strftime("%d %b %Y, %H:%M")})
    return found


@app.route("/", methods=["GET"])
def index():
    pnl_month = get_pnl_current_month()
    acplact_month = get_acplact_current_month()
    can_undo = len(latest_backup_batch()) > 0
    return render_template(
        "index.html",
        transformations=TRANSFORMATIONS,
        modes=MODES,
        pnl_month=pnl_month,
        acplact_month=acplact_month,
        acplact_status=acplact_master_status(),
        can_undo=can_undo,
        undo_lines=undo_preview() if can_undo else [],
        recent=recent_outputs(),
    )


@app.route("/process", methods=["POST"])
def process():
    file = request.files.get("sheet1_file")
    if not file or file.filename == "":
        flash("Please choose an Excel file to upload.", "error")
        return redirect(url_for("index"))

    selected = request.form.getlist("transforms")
    if not selected:
        flash("Please select at least one transformation to run.", "error")
        return redirect(url_for("index"))

    mode_key = request.form.get("mode", "new_month")
    if mode_key not in MODES:
        flash("Unknown mode selected.", "error")
        return redirect(url_for("index"))

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    upload_path = os.path.join(UPLOAD_DIR, f"sheet1_{ts}.xlsx")
    file.save(upload_path)
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump({"source": file.filename}, f)

    sections = []

    try:
        uploaded_wb = openpyxl.load_workbook(upload_path, data_only=True)
        uploaded_ws = uploaded_wb["Sheet1"] if "Sheet1" in uploaded_wb.sheetnames \
            else uploaded_wb.worksheets[0]
        # same file again with formulas intact, so arithmetic typed in Sheet1
        # (e.g. "=600000/12*5") stays visible in the output cells
        formulas_wb = openpyxl.load_workbook(upload_path)
        formulas_ws = formulas_wb["Sheet1"] if "Sheet1" in formulas_wb.sheetnames \
            else formulas_wb.worksheets[0]
        # A Sheet1 whose formula cells carry no saved result (exported by
        # another program, or saved by a script rather than Excel) reads as
        # EMPTY for every such cell: those codes would silently fall back to
        # last month's figures. Refuse it instead.
        uncached = [f"{formulas_ws.cell(row=r, column=12).coordinate}"
                    for r in range(7, formulas_ws.max_row + 1)
                    if isinstance(formulas_ws.cell(row=r, column=12).value, str)
                    and formulas_ws.cell(row=r, column=12).value.startswith("=")
                    and uploaded_ws.cell(row=r, column=12).value is None]
        if uncached:
            raise UploadNotCalculatedError(
                f"Nothing was changed. {len(uncached)} formula cell(s) in the uploaded Sheet1's Amount "
                f"column have no calculated value saved (e.g. {', '.join(uncached[:5])}) — the file was "
                f"not last saved by Excel. Open it in Excel, press Save, and upload it again.")
        by_code, gross_formula = read_total_formulas(formulas_ws)
        t1.set_sheet1_formulas(by_code, gross_formula)
        t2.set_sheet1_formulas(by_code, gross_formula)

        if "pnl" in selected:
            if not os.path.isfile(PNL_PATH):
                raise FileNotFoundError("No existing P&L STATEMENT.xlsx found in the project folder.")
            pnl_wb = openpyxl.load_workbook(PNL_PATH)  # keep formulas
            if "Sheet1" not in pnl_wb.sheetnames:
                pnl_wb.create_sheet("Sheet1")
            copy_sheet_values(uploaded_ws, pnl_wb["Sheet1"])
            run_fn = t1.apply_new_month if mode_key == "new_month" else t1.replace_latest_month
            report = run_fn(pnl_wb)
            pnl_wb.save(PENDING_PNL_PATH)
            sections.append({
                "key": "pnl",
                "label": TRANSFORMATIONS["pnl"]["label"],
                "report": report,
                "rows": build_pnl_review_rows(pnl_wb["P&L"], report),
                "editable": True,
            })

        if "acplact" in selected:
            if not os.path.isfile(ACPLACT_PATH):
                raise FileNotFoundError("No existing acplact.xlsx found in the project folder.")
            acplact_fp = file_fingerprint(ACPLACT_PATH)
            acplact_wb = openpyxl.load_workbook(ACPLACT_PATH)
            t2.repair_comma_formulas(acplact_wb)
            # Refuse to build on a damaged master: every month made from it
            # would inherit the broken links (this is exactly what happened
            # when an old, damaged acplact.xlsx was still the one being read)
            master_errors = t2.blocking_problems(t2.validate_workbook(acplact_wb))
            if master_errors:
                month = t2.get_current_month(acplact_wb)
                raise MasterDamagedError(
                    f"Nothing was changed. The ACPLACT file the tool is reading — {ACPLACT_PATH} "
                    f"(latest month {month.strftime('%B %Y') if month else 'unknown'}, saved "
                    f"{datetime.datetime.fromtimestamp(os.path.getmtime(ACPLACT_PATH)).strftime('%d %b %Y %H:%M')}) "
                    f"— is damaged, so it is not the corrected file: "
                    + "; ".join(master_errors[:4])
                    + (f"; and {len(master_errors) - 4} more" if len(master_errors) > 4 else "")
                    + ". Replace that exact file with the correct workbook (named exactly acplact.xlsx, "
                      "in that folder) and upload again. The home page shows which file is in use.")
            with open(META_PATH, "r+", encoding="utf-8") as f:
                meta = json.load(f)
                meta["acplact_fp"] = acplact_fp
                f.seek(0); f.truncate(); json.dump(meta, f)
            prev_totals = None
            if mode_key == "new_month":
                _t, upload_month, _g = t2.read_sheet1_totals(uploaded_ws)
                prev_totals = find_previous_sheet1_totals(t2._prev_month(upload_month), upload_path)
            t2.set_previous_totals(prev_totals)
            run_fn = t2.apply_new_month if mode_key == "new_month" else t2.replace_latest_month
            report = run_fn(acplact_wb, uploaded_ws)
            acplact_wb.save(PENDING_ACPLACT_PATH)

            sch_ws_now = acplact_wb[t2.FIFO_SCH_SHEET]
            sch_sections = t2.find_sections(sch_ws_now)
            sch_merge_targets = t2.list_sch_rows(sch_ws_now)
            prev_amt_col = t2._current_amount_col(sch_ws_now, t2.SCH_HEADER_ROW, t2.SCH_ANCHOR_COL) + 2
            new_field_candidates = []
            for code, desc, amt in report.get("unmapped_codes", []):
                formula = t2.TOTAL_FORMULAS.get(code)
                value_text = formula if formula else str(amt)
                # if last month this code's ledger figure was typed by hand
                # into some row (e.g. Solid Waste into "Solid Waste Testing
                # Charges"), suggest folding it into that same row
                default = ""
                last = (prev_totals or {}).get(code)
                if isinstance(last, (int, float)) and last:
                    for m in sch_merge_targets:
                        used = t2.eval_arithmetic(sch_ws_now.cell(row=m["row"], column=prev_amt_col).value)
                        if used is not None and abs(used - last) <= 1:
                            default = f"row::{m['row']}"
                            break
                new_field_candidates.append({"code": code, "desc": (desc or "").strip(),
                                             "value": value_text, "default": default})
            dup_conflicts = t2.find_duplicate_codes(acplact_wb)
            problems = t2.validate_workbook(acplact_wb)

            sections.append({
                "key": "acplact",
                "label": TRANSFORMATIONS["acplact"]["label"],
                "report": report,
                "rows": [],
                "editable": False,
                "editable_rows": build_acplact_editable_rows(acplact_wb, report),
                "sch_sections": sch_sections,
                "sch_merge_targets": sch_merge_targets,
                "new_field_candidates": new_field_candidates,
                "dup_conflicts": dup_conflicts,
                "problems": problems,
            })

        return render_template("review.html", sections=sections, mode_key=mode_key,
                                mode_label=MODES[mode_key]["label"],
                                source_labels=SOURCE_LABELS)

    except (t1.NotANewMonthError, t1.NotTheLatestMonthError,
            t2.NotANewMonthError, t2.NotTheLatestMonthError, MasterDamagedError,
            UploadNotCalculatedError) as e:
        for p in (PENDING_PNL_PATH, PENDING_ACPLACT_PATH):
            if os.path.isfile(p):
                os.remove(p)
        flash(str(e), "error")
        return redirect(url_for("index"))

    except Exception as e:
        for p in (PENDING_PNL_PATH, PENDING_ACPLACT_PATH):
            if os.path.isfile(p):
                os.remove(p)
        flash(f"Processing failed: {e}", "error")
        traceback.print_exc()
        return redirect(url_for("index"))


@app.route("/confirm", methods=["POST"])
def confirm():
    mode_key = request.form.get("mode_key", "new_month")
    ran_pnl = os.path.isfile(PENDING_PNL_PATH)
    ran_acplact = os.path.isfile(PENDING_ACPLACT_PATH)

    if not ran_pnl and not ran_acplact:
        flash("Nothing pending to confirm — please upload again.", "error")
        return redirect(url_for("index"))

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    downloads = []
    try:
        with open(META_PATH, encoding="utf-8") as f:
            meta = json.load(f)
    except (OSError, ValueError):
        meta = {}
    source_name = meta.get("source", "")

    if ran_pnl:
        pnl_wb = openpyxl.load_workbook(PENDING_PNL_PATH)
        ws = pnl_wb["P&L"]
        t1.detect_layout(ws)
        edited = 0
        pnl_bad = bad_inputs(request.form, ("cell_pnl_",), lambda k: cell_display_value(
            ws.cell(row=int(k[len("cell_pnl_"):]), column=t1.TOTAL_COL).value))
        if pnl_bad:
            os.remove(PENDING_PNL_PATH)
            ran_pnl = False
            downloads.append({"label": "PNL", "failed": True, "problems": pnl_bad})
    if ran_pnl:
        for key, text in request.form.items():
            if not key.startswith("cell_pnl_"):
                continue
            row_no = int(key[len("cell_pnl_"):])
            new_value = parse_edited_value(text)
            cell = ws.cell(row=row_no, column=t1.TOTAL_COL)
            if cell_display_value(cell.value) != text.strip():
                cell.value = new_value
                edited += 1

        new_month = ws.cell(row=7, column=t1.TOTAL_COL).value.date().isoformat()
        backup_path = os.path.join(HISTORY_DIR, f"P&L_before_{mode_key}_{ts}.xlsx")
        shutil.copy(PNL_PATH, backup_path)
        ws.column_dimensions["B"].hidden = False  # GL Code always visible
        pnl_wb.save(PNL_PATH)

        if "Sheet1" in pnl_wb.sheetnames:
            pnl_wb.remove(pnl_wb["Sheet1"])
        output_name = f"P&L STATEMENT ({new_month}).xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_name)
        pnl_wb.save(output_path)
        changes = earlier_month_changes("PNL", backup_path, PNL_PATH, mode_key)
        HISTORY.add("PNL", MODES[mode_key]["label"], new_month, PNL_PATH, ts, edited, source_name, changes)
        downloads.append({"label": "PNL", "filename": output_name, "edited": edited,
                          "changes": changes[:50], "changed_count": len(changes)})
        os.remove(PENDING_PNL_PATH)

    if ran_acplact and meta.get("acplact_fp") and file_fingerprint(ACPLACT_PATH) != meta.get("acplact_fp"):
        # the master was replaced/changed AFTER this review page was built:
        # the pending result is based on the OLD file — never save it
        os.remove(PENDING_ACPLACT_PATH)
        ran_acplact = False
        downloads.append({"label": "ACPLACT", "failed": True, "problems": [
            f"{ACPLACT_PATH} was changed after you uploaded (it was replaced or edited while the "
            "review page was open), so this review was built from the OLD file. Nothing was "
            "saved — please upload Sheet1 again so it is built from the current file."]})

    if ran_acplact:
        _pend = openpyxl.load_workbook(PENDING_ACPLACT_PATH)

        def _current(k):
            if not k.startswith("cell_acplact_"):
                return None
            sheet, row = k[len("cell_acplact_"):].split("_")
            ws_ = _pend[t2.FIFO_SCH_SHEET if sheet == "sch" else t2.FIFO_PNL_SHEET]
            col = t2._current_amount_col(ws_, t2.SCH_HEADER_ROW if sheet == "sch" else t2.PNL_HEADER_ROW,
                                         t2.SCH_ANCHOR_COL if sheet == "sch" else t2.PNL_ANCHOR_COL)
            return cell_display_value(ws_.cell(row=int(row), column=col).value)
        acplact_bad = bad_inputs(request.form, ("cell_acplact_", "nf_value_"), _current)
        if acplact_bad:
            os.remove(PENDING_ACPLACT_PATH)
            ran_acplact = False
            downloads.append({"label": "ACPLACT", "failed": True, "problems": acplact_bad})

    if ran_acplact:
        acplact_wb = openpyxl.load_workbook(PENDING_ACPLACT_PATH)
        sch_ws = acplact_wb[t2.FIFO_SCH_SHEET]
        pnl_ws = acplact_wb[t2.FIFO_PNL_SHEET]
        sch_amount_col = t2._current_amount_col(sch_ws, t2.SCH_HEADER_ROW, t2.SCH_ANCHOR_COL)
        pnl_amount_col = t2._current_amount_col(pnl_ws, t2.PNL_HEADER_ROW, t2.PNL_ANCHOR_COL)

        acplact_edited = 0
        for key, text in request.form.items():
            if key.startswith("cell_acplact_sch_"):
                ws, col = sch_ws, sch_amount_col
                row_no = int(key[len("cell_acplact_sch_"):])
            elif key.startswith("cell_acplact_pnl_"):
                ws, col = pnl_ws, pnl_amount_col
                row_no = int(key[len("cell_acplact_pnl_"):])
            else:
                continue
            new_value = parse_edited_value(text)
            cell = ws.cell(row=row_no, column=col)
            if cell_display_value(cell.value) != text.strip():
                cell.value = new_value
                acplact_edited += 1

        # new GL codes the client assigned to a section on the review page
        # (row numbers from build above are still valid — only columns were
        # touched since the review page was built, never rows)
        section_by_name = {s["name"]: s for s in t2.find_sections(sch_ws)}
        nf_indices = {k[len("nf_section_"):] for k in request.form if k.startswith("nf_section_")}
        new_field_assignments = []
        merged_fields = []
        for idx in nf_indices:
            dest = request.form.get(f"nf_section_{idx}", "").strip()
            code = request.form.get(f"nf_code_{idx}", "").strip()
            desc = request.form.get(f"nf_desc_{idx}", "").strip()
            if not dest or not code:
                continue
            raw_value = request.form.get(f"nf_value_{idx}", "")
            value = parse_edited_value(raw_value)
            if dest.startswith("row::"):
                # merges don't change row numbers, so these can be applied
                # right away, using the original (still-valid) row number —
                # do this BEFORE any new-row insertion below, which does shift rows
                target_row = int(dest[len("row::"):])
                t2.merge_code_into_row(sch_ws, target_row, code, value, sch_amount_col)
                merged_fields.append((target_row, code, desc))
            elif dest.startswith("sec::"):
                section = section_by_name.get(dest[len("sec::"):])
                if section:
                    new_field_assignments.append({
                        "code": code, "desc": f"    {desc}" if desc else f"    {code}",
                        "value": value, "insert_before": section["item_end"] + 1,
                    })
        new_rows = t2.insert_new_sch_rows(sch_ws, new_field_assignments)
        t2.repair_comma_formulas(acplact_wb)   # last line of defence before saving
        acplact_problems = t2.validate_workbook(acplact_wb)
        acplact_errors = t2.blocking_problems(acplact_problems)

    if ran_acplact and acplact_errors:
        # never let a broken result become the master or a download
        os.remove(PENDING_ACPLACT_PATH)
        ran_acplact = False
        downloads.append({"label": "ACPLACT", "failed": True, "problems": acplact_errors})

    if ran_acplact:
        new_month = t2.get_current_month(acplact_wb).isoformat()
        backup_path = os.path.join(HISTORY_DIR, f"acplact_before_{mode_key}_{ts}.xlsx")
        shutil.copy(ACPLACT_PATH, backup_path)
        acplact_wb.save(ACPLACT_PATH)

        output_name = f"acplact ({new_month}).xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_name)
        acplact_wb.save(output_path)
        changes = earlier_month_changes("ACPLACT", backup_path, ACPLACT_PATH, mode_key)
        HISTORY.add("ACPLACT", MODES[mode_key]["label"], new_month, ACPLACT_PATH, ts, acplact_edited,
                    source_name, changes)
        downloads.append({"label": "ACPLACT", "filename": output_name, "edited": acplact_edited,
                          "changes": changes[:50], "changed_count": len(changes),
                          "new_fields_added": new_rows, "fields_merged": merged_fields,
                          "problems": acplact_problems})
        os.remove(PENDING_ACPLACT_PATH)

    return render_template("result.html", downloads=downloads, mode_label=MODES[mode_key]["label"])


@app.route("/cancel", methods=["POST"])
def cancel():
    for p in (PENDING_PNL_PATH, PENDING_ACPLACT_PATH):
        if os.path.isfile(p):
            os.remove(p)
    flash("Discarded — nothing was saved.", "ok")
    return redirect(url_for("index"))


def _file_month(path, label):
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        d = t2.get_current_month(wb) if label == "ACPLACT" else t1.get_current_pnl_month(wb["P&L"])
        return d.strftime("%B %Y") if d else "no data"
    except Exception:
        return "unknown"


def _backup_target(backup):
    return (ACPLACT_PATH, "ACPLACT") if os.path.basename(backup).startswith("acplact_before_") \
        else (PNL_PATH, "PNL")


def undo_preview():
    """Plain-English list of what Undo would do right now."""
    lines = []
    for backup in latest_backup_batch():
        target, label = _backup_target(backup)
        now, then = _file_month(target, label), _file_month(backup, label)
        if now == then:
            lines.append(f"{label}: stays at {now}, but its figures go back to how they were before the last replace")
        else:
            lines.append(f"{label}: goes back from {now} to {then}")
    return lines


@app.route("/undo", methods=["POST"])
def undo():
    backups = latest_backup_batch()
    if not backups:
        flash("No previous version to restore to.", "error")
        return redirect(url_for("index"))

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    reverted = []
    for backup in backups:
        target, label = _backup_target(backup)
        shutil.copy(backup, target)
        os.remove(backup)
        HISTORY.add(label, "Undo", _file_month(target, label), target, ts,
                    note="Reverted to the version before the previous change")
        reverted.append(f"{label} (now {_file_month(target, label)})")

    flash(f"Reverted {' and '.join(reverted)} to their previous version.", "ok")
    return redirect(url_for("index"))


@app.route("/history")
def history_page():
    entries = list(reversed(HISTORY.entries()))
    return render_template("history.html", entries=entries,
                           pnl_month=get_pnl_current_month(), acplact_month=get_acplact_current_month())


@app.route("/history/<entry_id>")
def history_detail(entry_id):
    entry = HISTORY.get(entry_id)
    if not entry:
        flash("That history entry was not found.", "error")
        return redirect(url_for("history_page"))
    return render_template("history_detail.html", e=entry)


PREVIEW_MAX_ROWS, PREVIEW_MAX_COLS = 400, 40


@app.route("/history/<entry_id>/preview")
def history_preview(entry_id):
    entry = HISTORY.get(entry_id)
    if not entry or not os.path.isfile(HISTORY.version_path(entry)):
        flash("Saved copy not found.", "error")
        return redirect(url_for("history_page"))
    wb = openpyxl.load_workbook(HISTORY.version_path(entry))
    names = [n for n in wb.sheetnames if not (entry["kind"] == "PNL" and n == "Sheet1")]
    sheet = request.args.get("sheet", names[0])
    if sheet not in names:
        sheet = names[0]
    ws = wb[sheet]
    hidden_cols = {k for k, d in ws.column_dimensions.items() if d.hidden and k != "B"}
    rows = []
    for r in range(1, min(ws.max_row, PREVIEW_MAX_ROWS) + 1):
        cells = []
        for c in range(1, min(ws.max_column, PREVIEW_MAX_COLS) + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, datetime.datetime):
                v = v.strftime("%d %b %Y")
            elif isinstance(v, float):
                v = f"{v:,.2f}"
            elif isinstance(v, int) and not isinstance(v, bool) and abs(v) >= 1000:
                v = f"{v:,}"
            cells.append("" if v is None else str(v))
        if any(cells):
            rows.append((r, cells))
    ncols = min(ws.max_column, PREVIEW_MAX_COLS)
    return render_template("history_preview.html", e=entry, names=names, sheet=sheet, rows=rows,
                           ncols=ncols, truncated=ws.max_row > PREVIEW_MAX_ROWS or ws.max_column > PREVIEW_MAX_COLS,
                           col_letters=[openpyxl.utils.get_column_letter(i) for i in range(1, ncols + 1)])


@app.route("/history/<entry_id>/download")
def history_download(entry_id):
    entry = HISTORY.get(entry_id)
    if not entry or not os.path.isfile(HISTORY.version_path(entry)):
        flash("Saved copy not found.", "error")
        return redirect(url_for("history_page"))
    buf = io.BytesIO()
    wb = openpyxl.load_workbook(HISTORY.version_path(entry))
    if entry["kind"] == "PNL" and "Sheet1" in wb.sheetnames:
        wb.remove(wb["Sheet1"])
    wb.save(buf)
    buf.seek(0)
    base = "P&L STATEMENT" if entry["kind"] == "PNL" else "acplact"
    return send_file(buf, as_attachment=True,
                     download_name=f"{base} ({entry['month']}, saved {entry['id'][:8]}-{entry['id'][9:15]}).xlsx")


@app.route("/history/<entry_id>/restore", methods=["POST"])
def history_restore(entry_id):
    entry = HISTORY.get(entry_id)
    if not entry or not os.path.isfile(HISTORY.version_path(entry)):
        flash("Saved copy not found.", "error")
        return redirect(url_for("history_page"))
    target, name = (PNL_PATH, "P&L") if entry["kind"] == "PNL" else (ACPLACT_PATH, "acplact")
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    if os.path.isfile(target):   # keep what is there now, so Undo can reverse this too
        shutil.copy(target, os.path.join(HISTORY_DIR, f"{name}_before_restore_{ts}.xlsx"))
    shutil.copy(HISTORY.version_path(entry), target)
    HISTORY.add(entry["kind"], "Restore", entry["month"], target, ts,
                note=f"Restored the version saved {entry['when']}")
    flash(f"Restored {entry['kind']} to the version saved {entry['when']} ({entry['month']}).", "ok")
    return redirect(url_for("index"))


@app.route("/download/<path:filename>")
def download(filename):
    filename = os.path.basename(filename)
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.isfile(path):
        flash("File not found.", "error")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    app.run(debug=False, port=5050)
