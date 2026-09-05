"""
ALLOC AUTOMATION REMASTERD
==========================
Automation 1: FOR CALL OUTS

Pipeline:
  1. Read NOTES key (hierarchy lookup tables A:F and P:Q)
  2. Clean raw DRR  -> fix swapped dates (Date column only), fix account
     numbers (pad to full length, leading zeros preserved as text),
     set Card No. + Service No. = fixed Account No.
  3. Build VOLARE STATUS internally (CONCA/EFFORT/SORT/STATUS/SKIP STAT/
     REMARKS/RFD computed from NOTES) -- no upload needed
  4. Sort by SORT smallest-to-largest (ties -> earliest Date + Time)
  5. First-match lookup ("Call outs" & Account #) fills S/T/U in the
     CALL OUTS file as STATIC VALUES, original formatting preserved
  6. Output: CALL OUTS file only

Run:  streamlit run ALLOC_AUTOMATION_REMASTERD.py
"""

import io
import re
from collections import Counter
from datetime import datetime, time as dtime

import openpyxl
import streamlit as st

NA = "#N/A"
NO_EFFORT = "NO EFFORT"

# The 52 DRR headers expected in VOLARE STATUS H2:BG2, in canonical order
DRR_HEADERS = [
    "Date", "Time", "Debtor", "Account No.", "Card No.", "Service No.",
    "DPD", "Call Status", "Status", "Remark", "Remark By", "Remark Type",
    "Field Visit Date", "Collector", "Client", "Product Description",
    "Product Type", "Batch No", "Account Type", "Relation", "PTP Amount",
    "Next Call", "PTP Date", "Claim Paid Amount", "Claim Paid Date",
    "Dialed Number", "Days Past Write Off", "Balance", "Contact Type",
    "Black Case No.", "Red Case No.", "Court Name", "Lawyer", "Legal Stage",
    "Legal Status", "Next Legal Follow up", "Call Duration",
    "Talk Time Duration", "Cycle", "Old IC", "I.C Issue Date", "Bank Code",
    "Over Limit Amount", "Min Payment", "Due Date", "Monthly Installment",
    "30 Days", "MIA", "Area", "Debtor ID", "Last Pay Date", "Last Pay Amount",
]

IDX_DATE = 0
IDX_TIME = 1
IDX_ACCT = 3
IDX_CARD = 4
IDX_SVC = 5
IDX_STATUS = 8


def norm_header(h):
    return re.sub(r"\s+", " ", str(h)).strip().lower() if h is not None else ""


# ----------------------------------------------------------------------
# Universal Excel reader -- xlsx / xlsm / xlsb / xls
# (alloc files stay xlsx/xlsm because they are edited in place)
# ----------------------------------------------------------------------
DRR_DATE_HEADERS = {"date", "field visit date", "ptp date", "claim paid date",
                    "next legal follow up", "i.c issue date", "due date",
                    "last pay date"}
FIELD_DATE_HEADERS = {"date", "ptp-date"}
TIME_HEADERS = {"time"}


def _file_ext(filename):
    return (filename or "").lower().rsplit(".", 1)[-1]


DRR_NUMERIC_HEADERS = {"dpd", "ptp amount", "claim paid amount", "balance",
                       "over limit amount", "min payment",
                       "monthly installment", "30 days", "mia",
                       "last pay amount"}
_DMY_RE = re.compile(r"^\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})\s*$")
_ISO_RE = re.compile(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?\s*$")
_TIME_RE = re.compile(r"^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?\s*$", re.I)
_NUM_RE = re.compile(r"^\s*-?[\d,]+(?:\.\d+)?\s*$")


def _csv_detect_dayfirst(samples):
    """Infers whether d/m or m/d from a column's values. If any first part
    exceeds 12 -> day-first; if any second part exceeds 12 -> month-first;
    fully ambiguous columns default to day-first (source system standard)."""
    for s in samples:
        m = _DMY_RE.match(s)
        if m and int(m.group(1)) > 12:
            return True
    for s in samples:
        m = _DMY_RE.match(s)
        if m and int(m.group(2)) > 12:
            return False
    return True


def _csv_parse_date(s, dayfirst):
    m = _ISO_RE.match(s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return s
    m = _DMY_RE.match(s)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        dd, mm = (a, b) if dayfirst else (b, a)
        try:
            return datetime(y, mm, dd)
        except ValueError:
            return s
    return s


def _csv_parse_time(s):
    m = _TIME_RE.match(s)
    if not m:
        return s
    h, mi = int(m.group(1)), int(m.group(2))
    se = int(m.group(3) or 0)
    ap = (m.group(4) or "").upper()
    if ap == "PM" and h < 12:
        h += 12
    elif ap == "AM" and h == 12:
        h = 0
    try:
        return dtime(h, mi, se)
    except ValueError:
        return s


def _iter_csv_rows(file_bytes, date_headers, time_headers, numeric_headers):
    import csv as _csv
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("Could not decode the CSV file.")
    # detect delimiter from the header line (Sniffer is unreliable with
    # quoted remark fields); quoting stays standard Excel-style
    first_line = text.split("\n", 1)[0]
    counts = {d: first_line.count(d) for d in (",", ";", "\t", "|")}
    delim = max(counts, key=counts.get) if max(counts.values()) > 0 else ","
    rows = list(_csv.reader(io.StringIO(text), delimiter=delim))
    if not rows:
        return
    header = rows[0]
    date_idx, time_idx, num_idx = set(), set(), set()
    for i, h in enumerate(header):
        nh = norm_header(h)
        if nh in date_headers:
            date_idx.add(i)
        elif nh in time_headers:
            time_idx.add(i)
        elif nh in numeric_headers:
            num_idx.add(i)
    dayfirst = {}
    for i in date_idx:
        samples = []
        for r in rows[1:2001]:
            if i < len(r) and _DMY_RE.match(r[i] or ""):
                samples.append(r[i])
        dayfirst[i] = _csv_detect_dayfirst(samples)
    yield tuple((h.strip() if isinstance(h, str) else h) or None for h in header)
    for r in rows[1:]:
        out = []
        for i in range(len(header)):
            v = r[i] if i < len(r) else None
            if v is None or (isinstance(v, str) and v.strip() == ""):
                out.append(None)
                continue
            if i in date_idx:
                out.append(_csv_parse_date(v, dayfirst[i]))
            elif i in time_idx:
                out.append(_csv_parse_time(v))
            elif i in num_idx and isinstance(v, str) and _NUM_RE.match(v):
                n = float(v.replace(",", ""))
                out.append(int(n) if n.is_integer() else n)
            else:
                out.append(v)
        yield tuple(out)


def list_sheets_any(file_bytes, filename):
    ext = _file_ext(filename)
    if ext == "csv":
        return ["CSV"]
    try:
        if ext in ("xlsx", "xlsm"):
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
            names = list(wb.sheetnames); wb.close(); return names
        if ext == "xlsb":
            from pyxlsb import open_workbook
            with open_workbook(io.BytesIO(file_bytes)) as wb:
                return list(wb.sheets)
        if ext == "xls":
            import xlrd
            book = xlrd.open_workbook(file_contents=file_bytes)
            return book.sheet_names()
    except Exception:
        return []
    return []


def _xlsb_fix(v, is_date, is_time):
    from pyxlsb import convert_date
    if isinstance(v, float):
        if is_time:
            dt = convert_date(v % 1)
            return dt.time() if dt else v
        if is_date and v >= 1:
            return convert_date(v)
    return v


def iter_rows_any(file_bytes, filename, sheet_name=None,
                  date_headers=frozenset(), time_headers=frozenset(),
                  numeric_headers=frozenset()):
    """Yields value rows (header first) from any supported Excel format.
    xlsx/xlsm: openpyxl (types already correct). xls: xlrd (dates detected
    from cell type). xlsb: pyxlsb (date/time columns converted by HEADER
    NAME using date_headers/time_headers)."""
    ext = _file_ext(filename)
    if ext == "csv":
        yield from _iter_csv_rows(file_bytes, date_headers, time_headers,
                                  numeric_headers)
        return
    if ext in ("xlsx", "xlsm"):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        ws = None
        if sheet_name:
            for n in wb.sheetnames:
                if norm_header(n) == norm_header(sheet_name):
                    ws = wb[n]; break
        if ws is None:
            ws = wb.worksheets[0]
        try:
            for row in ws.iter_rows(values_only=True):
                yield row
        finally:
            wb.close()
    elif ext == "xlsb":
        from pyxlsb import open_workbook
        with open_workbook(io.BytesIO(file_bytes)) as wb:
            target = None
            if sheet_name:
                for n in wb.sheets:
                    if norm_header(n) == norm_header(sheet_name):
                        target = n; break
            if target is None:
                target = wb.sheets[0]
            with wb.get_sheet(target) as ws:
                date_idx, time_idx = set(), set()
                first = True
                for row in ws.rows(sparse=False):
                    vals = [c.v for c in row]
                    if first:
                        for i, h in enumerate(vals):
                            nh = norm_header(h)
                            if nh in date_headers: date_idx.add(i)
                            elif nh in time_headers: time_idx.add(i)
                        first = False
                        yield tuple(vals)
                        continue
                    yield tuple(
                        _xlsb_fix(v, i in date_idx, i in time_idx)
                        for i, v in enumerate(vals))
    elif ext == "xls":
        import xlrd
        book = xlrd.open_workbook(file_contents=file_bytes)
        sheet = None
        if sheet_name:
            for n in book.sheet_names():
                if norm_header(n) == norm_header(sheet_name):
                    sheet = book.sheet_by_name(n); break
        if sheet is None:
            sheet = book.sheet_by_index(0)
        for r in range(sheet.nrows):
            out = []
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    dt = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                    v = dt.time() if dt.year < 1901 else dt
                elif cell.ctype == xlrd.XL_CELL_EMPTY or v == "":
                    v = None
                out.append(v)
            yield tuple(out)
    else:
        raise ValueError(
            f"Unsupported file type '.{ext}'. Supported: .xlsx, .xlsm, "
            ".xlsb, .xls, .csv."
        )


# ----------------------------------------------------------------------
# Cleaning rules
# ----------------------------------------------------------------------
def fix_account(v):
    """Replicates:
    =IF(LEN(K)=8,"000000"&K, IF(LEN(K)=7,"0000000"&K, IF(LEN(K)=15,"0"&K)))
    Result is always stored as TEXT so leading zeros survive."""
    if v is None:
        return None
    s = str(int(v)).strip() if isinstance(v, float) and v.is_integer() else str(v).strip()
    if s.endswith(".0"):  # numeric leak from Excel
        s = s[:-2]
    n = len(s)
    if n == 8:
        return "000000" + s
    if n == 7:
        return "0000000" + s
    if n == 15:
        return "0" + s
    return s


def fix_drr_date(d):
    """Excel silently swaps day/month for dd/mm dates with day <= 12.
    Applies ONLY to the DRR 'Date' column (Due Date etc. are exempt)."""
    if isinstance(d, datetime) and d.day <= 12:
        try:
            return d.replace(month=d.day, day=d.month)
        except ValueError:
            return d
    return d


def peek_sheetnames(file_bytes, filename="file.xlsx"):
    """Cheap look at a workbook's sheet names (no data loaded)."""
    return list_sheets_any(file_bytes, filename)


def _precheck_alloc(alloc_bytes, sheet_keywords):
    """Cheap read-only validation that the alloc upload really contains an
    ALLOC sheet ('ch code' header) BEFORE the expensive editable load.
    Prevents out-of-memory crashes when a huge wrong file (e.g. the raw
    FIELD RESULT or DRR) is placed in the alloc slot."""
    wb = openpyxl.load_workbook(io.BytesIO(alloc_bytes), read_only=True)
    ws = None
    for name in wb.sheetnames:
        n = norm_header(name)
        if any(k in n for k in sheet_keywords):
            ws = wb[name]
            break
    if ws is None:
        ws = wb.worksheets[0]
    found = False
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=1, values_only=True):
        if norm_header(row[0] if row else None) == "ch code":
            found = True
            break
    names = list(wb.sheetnames)
    title = ws.title
    wb.close()
    if not found:
        raise ValueError(
            f"'{title}' doesn't look like an ALLOC sheet (no 'ch code' "
            f"header found). Sheets in this file: {', '.join(names)}. "
            "Check that the correct alloc file is in the second upload slot "
            "— the uploads may be swapped."
        )


# ----------------------------------------------------------------------
# Stage 1 -- NOTES key
# ----------------------------------------------------------------------
def load_notes(file_bytes, filename="notes.xlsx", sheet_hint="NOTES"):
    notes_af, notes_pq, notes_ij, notes_ux = {}, {}, {}, {}
    sheets = list_sheets_any(file_bytes, filename)
    target = None
    for n in sheets:
        if norm_header(n) == norm_header(sheet_hint):
            target = n
            break
    for row in iter_rows_any(file_bytes, filename, sheet_name=target):
        a = row[0] if len(row) > 0 else None
        if a is not None and a not in notes_af:
            # (B TouchPoint, C Pos/Neg, D Hierarchy, E SkipStat, F Remarks)
            hier = row[3] if len(row) > 3 else None
            if isinstance(hier, str) and hier.strip().isdigit():
                hier = int(hier.strip())     # CSV text -> numeric SORT
            notes_af[a] = (
                row[1] if len(row) > 1 else None,
                row[2] if len(row) > 2 else None,
                hier,
                row[4] if len(row) > 4 else None,
                row[5] if len(row) > 5 else None,
            )
        p = row[15] if len(row) > 15 else None
        if p is not None and p not in notes_pq:
            notes_pq[p] = row[16] if len(row) > 16 else None
        i9 = row[8] if len(row) > 8 else None
        if i9 is not None and i9 not in notes_ij:
            notes_ij[i9] = row[9] if len(row) > 9 else None
        u21 = row[20] if len(row) > 20 else None
        if u21 is not None and u21 not in notes_ux:
            notes_ux[u21] = (
                row[21] if len(row) > 21 else None,   # V -> PAYMENT TYPE
                row[22] if len(row) > 22 else None,   # W -> SOURCE OF PYT
                row[23] if len(row) > 23 else None,   # X -> COMPLYING/DEFAULTED
            )
    return notes_af, notes_pq, notes_ij, notes_ux


# ----------------------------------------------------------------------
# Unmapped-status fixer -- suggest NOTES entries for missing statuses,
# apply them in-session, and produce an updated NOTES file
# ----------------------------------------------------------------------
import difflib


def suggest_notes_mappings(missing_statuses, notes_af, notes_pq, notes_ux):
    """For each missing status, prefill NOTES columns from the most similar
    existing entry (these are usually spelling/format variants)."""
    af_keys = list(notes_af.keys())
    ux_keys = list(notes_ux.keys())
    rows = []
    for stt in missing_statuses:
        stt_s = str(stt)
        if stt in notes_af:
            tpl, af = "(already in NOTES A:F)", notes_af[stt]
            rfd = notes_pq.get(stt)
        else:
            m = difflib.get_close_matches(stt_s, af_keys, n=1, cutoff=0.4)
            tpl = m[0] if m else ""
            af = notes_af.get(tpl, (None, None, None, None, None))
            rfd = notes_pq.get(tpl)
        mu = difflib.get_close_matches(stt_s, ux_keys, n=1, cutoff=0.4)
        ux = notes_ux.get(mu[0], (None, None, None)) if mu else (None, None, None)
        rows.append({
            "ADD": True,
            "STATUS": stt_s,
            "COPIED FROM": str(tpl),
            "EFFORT (B)": af[0],
            "POS/NEG (C)": af[1],
            "SORT (D)": "" if af[2] is None else str(af[2]),
            "SKIP STAT (E)": af[3],
            "REMARKS (F)": af[4],
            "RFD (Q)": rfd,
            "PAYMENT TYPE (V)": ux[0],
            "SOURCE OF PYT (W)": ux[1],
            "COMPLIED/DEFAULTED (X)": ux[2],
        })
    return rows


def _rows_to_patch(rows):
    patch = {}
    for r in rows:
        if not r.get("ADD"):
            continue
        stt = r.get("STATUS")
        if not stt:
            continue
        sortv = r.get("SORT (D)")
        if isinstance(sortv, str):
            sv = sortv.strip()
            sortv = int(sv) if sv.isdigit() else (sv or None)
        af = (r.get("EFFORT (B)") or None, r.get("POS/NEG (C)") or None,
              sortv, r.get("SKIP STAT (E)") or None,
              r.get("REMARKS (F)") or None)
        ux = (r.get("PAYMENT TYPE (V)") or None,
              r.get("SOURCE OF PYT (W)") or None,
              r.get("COMPLIED/DEFAULTED (X)") or None)
        patch[stt] = {
            "af": af if any(v is not None for v in af) else None,
            "rfd": r.get("RFD (Q)") or None,
            "ux": ux if any(v is not None for v in ux) else None,
        }
    return patch


def _apply_notes_patch(notes_af, notes_pq, notes_ij, notes_ux, patch=None):
    """Merges session-saved status mappings into the loaded NOTES tables."""
    if patch is None:
        patch = st.session_state.get("notes_patch", {})
    for stt, p in patch.items():
        if p.get("af"):
            notes_af[stt] = tuple(p["af"])
        if p.get("rfd"):
            notes_pq[stt] = p["rfd"]
        if p.get("ux"):
            notes_ux[stt] = tuple(p["ux"])
    return notes_af, notes_pq, notes_ij, notes_ux


def build_updated_notes(notes_bytes, notes_fn, patch):
    """Returns a standalone NOTES workbook with the new statuses appended
    (A:F key + P:Q RFD + U:X CONFIRMED table on the same new rows --
    whole-column VLOOKUPs pick them up wherever they sit)."""
    ext = _file_ext(notes_fn)
    if ext in ("xlsx", "xlsm"):
        wb = openpyxl.load_workbook(io.BytesIO(notes_bytes))
        ws = None
        for name in wb.sheetnames:
            if norm_header(name) == "notes":
                ws = wb[name]
                break
        if ws is None:
            ws = wb.worksheets[0]
        for name in list(wb.sheetnames):
            if wb[name] is not ws:
                del wb[name]
        ws.title = "NOTES"
    else:   # xlsb / xls / csv source -> rebuild values-only
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "NOTES"
        sheets = list_sheets_any(notes_bytes, notes_fn)
        target = next((n for n in sheets if norm_header(n) == "notes"), None)
        for row in iter_rows_any(notes_bytes, notes_fn, sheet_name=target):
            ws.append(list(row))
    r = ws.max_row + 1
    for stt, p in patch.items():
        af = p.get("af") or (None, None, None, None, None)
        ws.cell(row=r, column=1, value=stt)
        for i, v in enumerate(af, start=2):
            ws.cell(row=r, column=i, value=v)
        if p.get("rfd"):
            ws.cell(row=r, column=16, value=stt)
            ws.cell(row=r, column=17, value=p["rfd"])
        if p.get("ux"):
            ws.cell(row=r, column=21, value=stt)
            for i, v in enumerate(p["ux"], start=22):
                ws.cell(row=r, column=i, value=v)
        r += 1
    out = io.BytesIO()
    wb.save(out)
    wb.close()
    out.seek(0)
    return out.getvalue()


# ----------------------------------------------------------------------
# Stage 2+3 -- clean raw DRR and build VOLARE STATUS in memory
# ----------------------------------------------------------------------
_MONTH_NAMES = {name: i for i, name in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"], start=1)}
for _n, _i in list(_MONTH_NAMES.items()):
    _MONTH_NAMES[_n[:3]] = _i          # jan, feb, ... aliases


def _month_from_title(title):
    """Extracts a month from a sheet name like 'June Reco 1' or 'APR 2026'."""
    for tok in re.split(r"[^a-z]+", str(title or "").lower()):
        if tok in _MONTH_NAMES:
            return _MONTH_NAMES[tok]
    return None


def _detect_swap(day_le12_pairs, months_hi, sheet_title=None):
    """Decides whether a sheet's dates were Excel-swapped (dd/mm misread).

    Primary evidence -- months_hi: months of rows with day >= 13
    (unambiguous, Excel cannot misread those). Under the swap bug an
    ambiguous row's DAY equals the true month; on a clean export its
    MONTH does.

    When a sheet has NO day>=13 rows (e.g. an early-month export):
    1. dispersion -- swapped rows share one DAY with many months, clean
       rows share one MONTH with many days;
    2. the sheet title's month name ('June Reco 1') as a tie-break;
    3. legacy assumption (swap) as the last resort."""
    if not day_le12_pairs:
        return False, 0, 0
    if months_hi:
        clean_hits = sum(1 for m, d in day_le12_pairs if m in months_hi)
        bug_hits = sum(1 for m, d in day_le12_pairs if d in months_hi)
        return (bug_hits >= clean_hits), clean_hits, bug_hits
    uniq_months = len({m for m, d in day_le12_pairs})
    uniq_days = len({d for m, d in day_le12_pairs})
    if uniq_months < uniq_days:
        return False, uniq_days, uniq_months     # one month, many days -> clean
    if uniq_days < uniq_months:
        return True, uniq_days, uniq_months      # one day, many months -> swapped
    tm = _month_from_title(sheet_title)
    if tm:
        m_hits = sum(1 for m, d in day_le12_pairs if m == tm)
        d_hits = sum(1 for m, d in day_le12_pairs if d == tm)
        if m_hits != d_hits:
            return (d_hits > m_hits), m_hits, d_hits
    return True, 0, 0                            # legacy assumption


def build_volare(drr_bytes, notes_af, notes_pq, progress_cb=None,
                 drr_name="drr.xlsx", swap_mode="auto"):
    """Reads EVERY qualifying sheet of the raw DRR (Sheet 1 / Sheet 2 /
    monthly tabs etc.), aligns each sheet's columns by header NAME, and
    builds one combined VOLARE STATUS.

    swap_mode: "auto"  -> per-sheet detection of the Excel day/month swap
               "force" -> always apply the swap (legacy dd/mm misread)
               "off"   -> never swap (dates already correct)"""
    is_csv = _file_ext(drr_name) == "csv"
    sheet_names = list_sheets_any(drr_bytes, drr_name) or [None]

    vol_rows = []
    unmapped_status = Counter()
    date_swapped = 0
    acct_lens = Counter()
    month_dist = Counter()
    total_rows = 0
    _acct_cache = {}       # raw acct -> fixed string (shared object)
    _conca_cache = {}      # (effort, acct) -> conca string (shared object)
    missing_union, extra_union = set(), set()
    sheets_used, sheets_skipped = [], []
    any_qualified = False
    first_header_sample = []

    for s_name in sheet_names:
        rows_iter = iter_rows_any(drr_bytes, drr_name, sheet_name=s_name,
                                  date_headers=DRR_DATE_HEADERS,
                                  time_headers=TIME_HEADERS,
                                  numeric_headers=DRR_NUMERIC_HEADERS)
        raw_header = None
        for row in rows_iter:
            raw_header = list(row)
            break
        if raw_header is None:
            sheets_skipped.append(s_name or "(first sheet)")
            continue
        if not first_header_sample:
            first_header_sample = [str(h) for h in raw_header if h is not None][:8]

        raw_pos = {norm_header(h): i for i, h in enumerate(raw_header)
                   if h is not None}
        col_map, missing = [], []
        for h in DRR_HEADERS:
            key = norm_header(h)
            if key in raw_pos:
                col_map.append(raw_pos[key])
            else:
                col_map.append(None)
                missing.append(h)
        if len(missing) > 20:          # not a DRR sheet -> skip it
            sheets_skipped.append(s_name or "(first sheet)")
            continue
        any_qualified = True
        missing_union.update(missing)
        known = {norm_header(x) for x in DRR_HEADERS}
        extra_union.update(
            str(raw_header[i]) for i, h in enumerate(raw_header)
            if h is not None and norm_header(h) not in known)

        sheet_start = len(vol_rows)
        sheet_rows = 0
        day_le12_pairs = []
        months_hi = set()

        for row in rows_iter:
            total_rows += 1
            sheet_rows += 1
            aligned = [row[i] if (i is not None and i < len(row)) else None
                       for i in col_map]

            if all(v is None or str(v).strip() == "" for v in aligned):
                total_rows -= 1
                sheet_rows -= 1
                continue

            d0 = aligned[IDX_DATE]
            if isinstance(d0, datetime):
                if d0.day <= 12:
                    day_le12_pairs.append((d0.month, d0.day))
                else:
                    months_hi.add(d0.month)

            raw_acct = aligned[IDX_ACCT]
            if raw_acct is not None:
                acct_lens[len(str(raw_acct).strip())] += 1
            acct = _acct_cache.get(raw_acct)
            if acct is None and raw_acct is not None:
                acct = fix_account(raw_acct)
                _acct_cache[raw_acct] = acct
            aligned[IDX_ACCT] = acct
            aligned[IDX_CARD] = acct   # Card No.    = fixed Account No.
            aligned[IDX_SVC] = acct    # Service No. = fixed Account No.

            status = aligned[IDX_STATUS]
            n = notes_af.get(status)
            if n:
                effort, posneg, hier, skipstat, remarks = n
            else:
                effort = posneg = hier = skipstat = remarks = NA
                if status is not None:
                    unmapped_status[status] += 1
            rfd = notes_pq.get(status, NA)
            if effort == NA:
                conca = NA
            else:
                ck = (effort, acct)
                conca = _conca_cache.get(ck)
                if conca is None:
                    conca = f"{effort}{acct or ''}"
                    _conca_cache[ck] = conca

            vol_rows.append([conca, effort, hier, posneg, skipstat, remarks,
                             rfd] + aligned)

            if progress_cb and total_rows % 25000 == 0:
                progress_cb(total_rows)

        # ---- per-sheet swap decision, applied after the sheet is read ----
        if is_csv or swap_mode == "off":
            do_swap = False
        elif swap_mode == "force":
            do_swap = True
        else:
            do_swap, _, _ = _detect_swap(day_le12_pairs, months_hi, s_name)

        sheet_swapped = 0
        if do_swap:
            for r in vol_rows[sheet_start:]:
                d = r[7 + IDX_DATE]
                if isinstance(d, datetime) and d.day <= 12:
                    r[7 + IDX_DATE] = d.replace(month=d.day, day=d.month)
                    sheet_swapped += 1
        elif is_csv:
            sheet_swapped = len(day_le12_pairs)   # comparable metric

        date_swapped += sheet_swapped
        for r in vol_rows[sheet_start:]:
            d = r[7 + IDX_DATE]
            if isinstance(d, datetime):
                month_dist[d.month] += 1

        sheets_used.append({
            "name": s_name or "(first sheet)",
            "rows": sheet_rows,
            "swap": ("swapped " + f"{sheet_swapped:,}" if do_swap
                     else "dates already correct — no swap"),
        })

    if not any_qualified:
        raise ValueError(
            "This file doesn't look like a RAW DRR — no sheet has the "
            "expected DRR headers. First headers found: "
            + ", ".join(first_header_sample) + ". "
            "Check that the raw DRR export is in the first upload slot."
        )

    stats = {
        "total_rows": total_rows,
        "date_swapped": date_swapped,
        "acct_lens": dict(acct_lens),
        "unmapped_status": unmapped_status,
        "missing_headers": sorted(missing_union),
        "extra_headers": sorted(extra_union),
        "month_dist": dict(month_dist),
        "sheets_used": sheets_used,
        "sheets_skipped": sheets_skipped,
    }
    return vol_rows, stats


# ----------------------------------------------------------------------
# Stage 4 -- sort by SORT smallest-to-largest (Excel semantics, stable),
#            then Date (column H) NEWEST to OLDEST, then Time newest first
#            -> the most recent effort wins the first-match lookup
# ----------------------------------------------------------------------
def _date_key(r, newest_first=True):
    """Date+time tie-break key. newest_first=True -> newer sorts first
    (Excel: Sort Newest to Oldest, then SORT smallest to largest).
    Rows without a date sort last either way."""
    d, t = r[7], r[8]
    if isinstance(d, datetime):
        dk = -d.toordinal() if newest_first else d.toordinal()
    else:
        dk = float("inf")
    if isinstance(t, dtime):
        s = t.hour * 3600 + t.minute * 60 + t.second
        tk = -s if newest_first else s
    else:
        tk = float("inf")
    return dk, tk


def sort_volare(vol_rows, newest_first=True):
    """Replicates the manual two-step Excel process:
    Step 1 - sort Date (column H) (newest->oldest by default)
    Step 2 - sort SORT (column C) smallest->largest (stable)
    => hierarchy primary, date order as tie-break, so the FIRST match per
    account is the best effort. POS additionally beats NEG within a tier."""
    def key(r):
        v = r[2]                                   # SORT
        pk = 0 if r[3] == "POS" else 1             # POS beats NEG within a tier
        dk, tk = _date_key(r, newest_first)        # Date/Time tie-break
        if isinstance(v, bool):                    # bools are not "numbers" in Excel sort
            return (1, 0, str(v).upper(), pk, dk, tk)
        if isinstance(v, (int, float)):
            return (0, v, "", pk, dk, tk)          # numbers first, ascending
        if isinstance(v, str):
            if v == NA:
                return (2, 0, "", pk, dk, tk)      # errors after text
            return (1, 0, v.upper(), pk, dk, tk)   # text A-Z after numbers
        return (3, 0, "", pk, dk, tk)              # blanks last
    return sorted(vol_rows, key=key)


def build_lookup_index(vol_sorted):
    """VLOOKUP first-match, case-insensitive on CONCA."""
    index = {}
    for r in vol_sorted:
        k = str(r[0]).lower()
        if k not in index:
            index[k] = r
    return index


# ----------------------------------------------------------------------
# Stage 5 -- fill S/T/U in the CALL OUTS file (static values,
#            original formatting preserved)
# ----------------------------------------------------------------------
def fill_alloc(alloc_bytes, index, effort_key, sheet_keywords, progress_cb=None, vba=False):
    """Generic S/T/U filler. effort_key: VLOOKUP prefix ("call outs" /
    "skiptrace"). sheet_keywords: substrings to locate the target sheet."""
    _precheck_alloc(alloc_bytes, sheet_keywords)
    wb = openpyxl.load_workbook(io.BytesIO(alloc_bytes), keep_vba=vba)

    ws = None
    for name in wb.sheetnames:
        n = norm_header(name)
        if any(k in n for k in sheet_keywords):
            ws = wb[name]
            break
    if ws is None:
        ws = wb.worksheets[0]

    # locate header row (the row where column A == 'ch code')
    header_row = None
    for r in range(1, 11):
        if norm_header(ws.cell(row=r, column=1).value) == "ch code":
            header_row = r
            break
    if header_row is None:
        names = list(wb.sheetnames)
        wb.close()
        raise ValueError(
            f"'{ws.title}' doesn't look like an ALLOC sheet (no 'ch code' "
            f"header found). Sheets in this file: {', '.join(names)}. "
            "Check that the correct alloc file is in the second upload slot."
        )

    filled = 0
    no_effort = 0
    pos = neg = 0
    result_breakdown = Counter()   # report only -- does not affect output
    status_breakdown = Counter()   # report only -- does not affect output
    min_date = max_date = None     # report only -- does not affect output

    # clear stray formulas above the header band (e.g. S1:U1 template refs)
    for r in range(1, header_row):
        for c in (19, 20, 21):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None

    for r in range(header_row + 1, ws.max_row + 1):
        acct = ws.cell(row=r, column=4).value  # D = ACCOUNT NUMBER
        s_cell = ws.cell(row=r, column=19)     # S = CALL DATE
        t_cell = ws.cell(row=r, column=20)     # T = CALL OUTS (POS/NEG)
        u_cell = ws.cell(row=r, column=21)     # U = CALL RESULT (Remarks)

        if acct is None or str(acct).strip() == "":
            for cell in (s_cell, t_cell, u_cell):
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = None
            continue

        m = index.get(f"{effort_key}{acct}".lower())
        if m:
            call_date = m[7]                   # VOLARE H = Date
            s_cell.value = call_date
            if isinstance(call_date, datetime):
                s_cell.number_format = "MM/DD/YYYY"
                if min_date is None or call_date < min_date:
                    min_date = call_date
                if max_date is None or call_date > max_date:
                    max_date = call_date
            t_cell.value = m[3]                # VOLARE D = STATUS
            u_cell.value = m[5]                # VOLARE F = REMARKS
            status_breakdown[str(m[3])] += 1
            result_breakdown[str(m[5]).strip()] += 1
            if m[3] == "POS":
                pos += 1
            elif m[3] == "NEG":
                neg += 1
        else:
            s_cell.value = NO_EFFORT
            t_cell.value = NO_EFFORT
            u_cell.value = NO_EFFORT
            no_effort += 1
        filled += 1

        if progress_cb and filled % 2000 == 0:
            progress_cb(filled)

    # output: alloc sheet + NOTES (template kept)
    for name in list(wb.sheetnames):
        if wb[name] is not ws and norm_header(name) != "notes":
            del wb[name]

    out = io.BytesIO()
    wb.save(out)
    wb.close()
    out.seek(0)
    return out.getvalue(), {
        "filled": filled, "no_effort": no_effort, "pos": pos, "neg": neg,
        "result_breakdown": result_breakdown,
        "status_breakdown": status_breakdown,
        "min_date": min_date, "max_date": max_date,
    }


def fill_call_outs(callouts_bytes, index, progress_cb=None, vba=False):
    return fill_alloc(callouts_bytes, index, "call outs", ("call outs",), progress_cb, vba)


VOLARE_SHEET_HEADERS = ["CONCA", "EFFORT", "SORT", "STATUS", "SKIP STAT",
                        "REMARKS", "RFD"] + DRR_HEADERS

# Excel hard limit is 1,048,576 rows per sheet; use a round safe cap.
VOLARE_MAX_ROWS_PER_SHEET = 1_000_000


def build_volare_file(vol_sorted, progress_cb=None):
    """Writes the full VOLARE STATUS (A:G computed + H:BG cleaned DRR) to its
    own workbook using write-only mode, so the huge sheet never has to live
    inside the formatted alloc workbook (memory-safe).

    The FILE is written in the exact working order the lookups use:
    Date sorted first, then SORT smallest -> largest (with POS over NEG)
    -- so a manual VLOOKUP on this sheet reproduces the alloc results.

    When the data exceeds Excel's 1,048,576-row sheet limit, it continues
    onto numbered sheets -- VOLARE STATUS (1), (2), ... -- cut sequentially
    so the working order is preserved ACROSS the sheets (an account's first
    match is always in the earliest sheet it appears in)."""
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    total = len(vol_sorted)
    multi = total > VOLARE_MAX_ROWS_PER_SHEET
    wb = Workbook(write_only=True)
    ws = None
    part = 0
    rows_in_sheet = VOLARE_MAX_ROWS_PER_SHEET   # force sheet creation
    n = 0
    for r in vol_sorted:
        if rows_in_sheet >= VOLARE_MAX_ROWS_PER_SHEET:
            part += 1
            title = f"VOLARE STATUS ({part})" if multi else "VOLARE STATUS"
            ws = wb.create_sheet(title)
            ws.append(VOLARE_SHEET_HEADERS)
            rows_in_sheet = 0
        row = list(r)
        if isinstance(row[7], datetime):
            c = WriteOnlyCell(ws, value=row[7])
            c.number_format = "MM/DD/YYYY"
            row[7] = c
        ws.append(row)
        rows_in_sheet += 1
        n += 1
        if progress_cb and n % 50000 == 0:
            progress_cb(n)
    if ws is None:                               # empty input edge case
        ws = wb.create_sheet("VOLARE STATUS")
        ws.append(VOLARE_SHEET_HEADERS)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# ----------------------------------------------------------------------
# FIELD VISITATION pipeline (no DRR / VOLARE needed)
# ----------------------------------------------------------------------
# The 65 headers of the RESULT sheet block C:BO, which map 1:1 onto the
# FIELD reference sheet columns C:BO
FIELD_HEADERS = [
    "chcode", "status", "SUB STATUS", "EMAIL ADDRESS", "AUTOFIELD AREA",
    "CH NAME", "ADDRESS", "MUNICIPALITY", "BARANGAY", "SUBDIVISION",
    "LANDMARK", "AVAILABILITY", "ADDTYPE", "VISITEDADDTYPE", "DLType",
    "INFORMANT", "3RD PARTY LIST", "NAME OF 3RD PARTY", "CLIENT NUMBER",
    "INFORMANT NUMBER", "DL RECEIVED/UNRECEIVED", "Message", "PTP-Date",
    "PTP AMOUNT", "field_name", "DATE", "TIME", "REPORT TO",
    "OFFICER IN TAG", "ID NUMBER", "bank", "BRANCH", "CONCA", "SORT",
    "MONTH", "SUMMARY", "DUPPLICATES", "NUMBER CHECKING", "DAILY PROD",
    "reference_code", "TYPE", "AREA", "Repo Client Status",
    "Repo Client Substatus", "Repo Client Substatus 2",
    "Repo Client Substatus 3", "Unit Status", "Unit Substatus",
    "Unit Substatus 2", "Unit Substatus 3", "PSTR Classification", "RFD",
    "POSITION", "MC", "NEW POSITION", "PLACEMENT", "DL RECEIVED BY",
    "DL RECEIVER NAME", "CLIENT STATUS", "3RD PARTY NUMBER", "SBC RFD",
    "SBC RPC SUBSTATUS", "OB", "NEW BANK NAME", "ACCOUNT NUMBER",
]
F_STATUS = 1      # status
F_SUBSTAT = 2     # SUB STATUS        -> W VISITATION REMARKS
F_3RDPARTY = 16   # 3RD PARTY LIST    -> V INFORMANT
F_DATE = 25       # DATE              -> S VISIT DATE
F_MC = 53         # MC
F_PLACEMENT = 55  # PLACEMENT
F_NEWBANK = 63    # NEW BANK NAME
F_ACCT = 64       # ACCOUNT NUMBER

NO_VISIT = "NO VISITATION"


def fix_account_excel(v):
    """Faithful Excel semantics of the account-fix formula: lengths other
    than 7/8/15 return FALSE (never matches a real account)."""
    if v is None:
        return False
    s = str(int(v)).strip() if isinstance(v, float) and v.is_integer() else str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    n = len(s)
    if n == 8:
        return "000000" + s
    if n == 7:
        return "0000000" + s
    if n == 15:
        return "0" + s
    return False


def build_field_index(field_bytes, notes_ij, mc_value, bank_names,
                      placement_keyword, progress_cb=None,
                      field_name="field.xlsx"):
    """Reads the raw FIELD RESULT file's RESULT sheet, filters rows
    (MC + NEW BANK NAME + PLACEMENT contains keyword), computes
    A (fixed account) and B (STAT via NOTES I:J), and returns the
    first-match lookup index -- the in-memory equivalent of the FIELD
    reference sheet."""
    sheets = list_sheets_any(field_bytes, field_name)
    target = None
    for name in sheets:
        if norm_header(name) == "result":
            target = name
            break
    if target is None and _file_ext(field_name) == "csv":
        target = "CSV"              # a CSV *is* the RESULT table
    if target is None:
        names = sheets
        hint = ""
        if any("drr" in norm_header(n) or norm_header(n) == "sheet1" for n in names) and len(names) <= 2:
            hint = (" This file looks like a RAW DRR — reminder: the Field "
                    "Visitation module does NOT use the raw DRR.")
        raise ValueError(
            "No 'RESULT' sheet found in the FIELD RESULT file. "
            f"Sheets found: {', '.join(names) if names else '(none)'}.{hint} "
            "Upload the raw FIELD RESULT file (the one with WRONG STATUS / "
            "Sheet1 / RESULT sheets) in the first slot."
        )

    rows_iter = iter_rows_any(field_bytes, field_name, sheet_name=target,
                              date_headers=FIELD_DATE_HEADERS,
                              time_headers=TIME_HEADERS)
    raw_header = None
    for row in rows_iter:
        raw_header = list(row)
        break
    if raw_header is None:
        raise ValueError("RESULT sheet appears to be empty.")

    raw_pos = {norm_header(h): i for i, h in enumerate(raw_header) if h is not None}
    col_map, missing = [], []
    for h in FIELD_HEADERS:
        key = norm_header(h)
        if key in raw_pos:
            col_map.append(raw_pos[key])
        else:
            col_map.append(None)
            missing.append(h)

    if len(missing) > 30:
        found = [str(h) for h in raw_header if h is not None][:8]
        raise ValueError(
            "This file doesn't look like a FIELD RESULT export — most "
            f"expected headers are missing. First headers found: "
            f"{', '.join(found)}."
        )
    banks = {b.strip().upper() for b in bank_names if b.strip()}
    kw = placement_keyword.strip().upper()
    mc_norm = mc_value.strip().upper()

    index = {}
    kept_rows = []
    total = kept = 0
    bank_counts = Counter()
    unmapped_status = Counter()
    dup_skipped = 0

    for row in rows_iter:
        total += 1
        aligned = [row[i] if (i is not None and i < len(row)) else None for i in col_map]
        if all(v is None or str(v).strip() == "" for v in aligned):
            total -= 1
            continue

        mc = str(aligned[F_MC]).strip().upper() if aligned[F_MC] is not None else ""
        nb = str(aligned[F_NEWBANK]).strip().upper() if aligned[F_NEWBANK] is not None else ""
        pl = str(aligned[F_PLACEMENT]).strip().upper() if aligned[F_PLACEMENT] is not None else ""
        if mc != mc_norm or nb not in banks or kw not in pl:
            continue
        kept += 1
        bank_counts[aligned[F_NEWBANK]] += 1

        status = aligned[F_STATUS]
        stat = notes_ij.get(status)
        if stat is None:
            stat = NA
            if status is not None:
                unmapped_status[status] += 1

        acct_key = str(fix_account_excel(aligned[F_ACCT]))
        kept_rows.append((acct_key, stat, aligned))
        if acct_key in index:
            dup_skipped += 1          # VLOOKUP keeps the FIRST match
            continue
        index[acct_key] = {
            "S": aligned[F_DATE],      # VISIT DATE
            "U": stat,                 # VISITATION RESULT (POS/NEG)
            "V": aligned[F_3RDPARTY],  # INFORMANT
            "W": aligned[F_SUBSTAT],   # VISITATION REMARKS
        }
        if progress_cb and total % 10000 == 0:
            progress_cb(total)

    stats = {
        "total_rows": total,
        "kept_rows": kept,
        "bank_counts": bank_counts,
        "dup_skipped": dup_skipped,
        "unmapped_status": unmapped_status,
        "missing_headers": missing,
        "filters": {"MC": mc_value, "banks": sorted(banks),
                    "placement": placement_keyword},
    }
    return index, stats, kept_rows


def fill_visitation(alloc_bytes, index, sheet_keywords, field_rows=None, progress_cb=None, vba=False):
    """Fills S/U/V/W in the ALLOC FIELD VISITATION sheet as static values.
    T (VISITATION THRU) is left untouched. Lookup key = column C (ACCT #).
    Output keeps the alloc sheet only; formatting preserved."""
    _precheck_alloc(alloc_bytes, sheet_keywords)
    wb = openpyxl.load_workbook(io.BytesIO(alloc_bytes), keep_vba=vba)
    ws = None
    for name in wb.sheetnames:
        n = norm_header(name)
        if any(k in n for k in sheet_keywords):
            ws = wb[name]
            break
    if ws is None:
        ws = wb.worksheets[0]

    header_row = None
    for r in range(1, 11):
        if norm_header(ws.cell(row=r, column=1).value) == "ch code":
            header_row = r
            break
    if header_row is None:
        names = list(wb.sheetnames)
        wb.close()
        raise ValueError(
            f"'{ws.title}' doesn't look like an ALLOC sheet (no 'ch code' "
            f"header found). Sheets in this file: {', '.join(names)}. "
            "Make sure the ALLOC FIELD VISITATION file is in the second "
            "slot — it looks like the uploads may be swapped."
        )

    TARGET_COLS = (19, 21, 22, 23)   # S, U, V, W  (T=20 untouched)
    filled = no_visit = pos = neg = 0
    result_breakdown = Counter()
    status_breakdown = Counter()
    min_date = max_date = None

    for r in range(1, header_row):
        for c in TARGET_COLS:
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None

    for r in range(header_row + 1, ws.max_row + 1):
        acct = ws.cell(row=r, column=3).value   # C = ACCT #
        cells = {c: ws.cell(row=r, column=c) for c in TARGET_COLS}

        if acct is None or str(acct).strip() == "":
            for cell in cells.values():
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = None
            continue

        m = index.get(str(acct))
        if m:
            visit_date = m["S"]
            cells[19].value = visit_date
            if isinstance(visit_date, datetime):
                cells[19].number_format = "MM/DD/YYYY"
                if min_date is None or visit_date < min_date:
                    min_date = visit_date
                if max_date is None or visit_date > max_date:
                    max_date = visit_date
            cells[21].value = m["U"]
            cells[22].value = m["V"]
            cells[23].value = m["W"]
            status_breakdown[str(m["U"])] += 1
            result_breakdown[str(m["W"]).strip()] += 1
            if m["U"] == "POS":
                pos += 1
            elif m["U"] == "NEG":
                neg += 1
        else:
            for c in TARGET_COLS:
                cells[c].value = NO_VISIT
            no_visit += 1
        filled += 1
        if progress_cb and filled % 2000 == 0:
            progress_cb(filled)

    for name in list(wb.sheetnames):
        if wb[name] is not ws and norm_header(name) != "notes":
            del wb[name]

    # include the rebuilt FIELD reference sheet (template kept)
    if field_rows:
        wsf = wb.create_sheet("FIELD")
        wsf.append(["chcode", "STAT"] + FIELD_HEADERS)
        for (a_key, stat, aligned) in field_rows:
            wsf.append([a_key, stat] + list(aligned))
            dcell = wsf.cell(row=wsf.max_row, column=2 + F_DATE + 1)
            if isinstance(dcell.value, datetime):
                dcell.number_format = "MM/DD/YYYY"

    out = io.BytesIO()
    wb.save(out)
    wb.close()
    out.seek(0)
    return out.getvalue(), {
        "filled": filled, "no_effort": no_visit, "pos": pos, "neg": neg,
        "result_breakdown": result_breakdown,
        "status_breakdown": status_breakdown,
        "min_date": min_date, "max_date": max_date,
    }


def build_visit_report(stats, fill_stats, run_ts, field_name, alloc_name,
                       notes_name, elapsed, mod):
    lines = []
    w = lines.append
    bar = "=" * 62
    w(bar)
    w("  ALLOC AUTOMATION REMASTERD  ·  STATUS REPORT")
    w(f"  Module: {mod['name']}")
    w(bar)
    w(f"  Run date/time    : {run_ts:%m/%d/%Y %I:%M %p}")
    w(f"  Processing time  : {elapsed:.1f} seconds")
    w(f"  Field Result file: {field_name}")
    w(f"  Alloc file       : {alloc_name}")
    w(f"  NOTES key source : {notes_name}")
    w("")
    w("  STAGE 1 · FIELD RESULT FILTER (RESULT sheet)")
    w("  " + "-" * 50)
    w(f"  Rows in RESULT sheet        : {stats['total_rows']:,}")
    f = stats["filters"]
    w(f"  Filter MC                   : {f['MC']}")
    w(f"  Filter NEW BANK NAME        : {', '.join(f['banks'])}")
    w(f"  Filter PLACEMENT contains   : {f['placement']}")
    w(f"  Rows kept after filter      : {stats['kept_rows']:,}")
    for b, c in stats["bank_counts"].most_common():
        w(f"      - {b}  ({c:,} rows)")
    w(f"  Duplicate accounts skipped  : {stats['dup_skipped']:,} "
      f"(first visit kept, same as VLOOKUP)")
    w(f"  Missing RESULT headers      : "
      + (", ".join(stats["missing_headers"]) if stats["missing_headers"] else "None"))
    unmapped_total = sum(stats["unmapped_status"].values())
    w(f"  Statuses missing from NOTES : {unmapped_total:,} rows "
      f"({len(stats['unmapped_status'])} distinct)")
    if stats["unmapped_status"]:
        for s, c in stats["unmapped_status"].most_common(10):
            w(f"      - {s}  ({c:,} rows)")
    w("")
    w(f"  STAGE 2 · {mod['name']} FILL (S/U/V/W — T untouched)")
    w("  " + "-" * 50)
    filled = fill_stats["filled"] or 1
    w(f"  Accounts filled             : {fill_stats['filled']:,}")
    w(f"  POS                         : {fill_stats['pos']:,}"
      f"  ({fill_stats['pos'] / filled * 100:.1f}%)")
    w(f"  NEG                         : {fill_stats['neg']:,}"
      f"  ({fill_stats['neg'] / filled * 100:.1f}%)")
    w(f"  NO VISITATION               : {fill_stats['no_effort']:,}"
      f"  ({fill_stats['no_effort'] / filled * 100:.1f}%)")
    if fill_stats["min_date"] and fill_stats["max_date"]:
        w(f"  {mod['date_label']:<28}: "
          f"{fill_stats['min_date']:%m/%d/%Y} - {fill_stats['max_date']:%m/%d/%Y}")
    w("")
    w("  TOP 10 RESULTS (VISITATION REMARKS)")
    w("  " + "-" * 50)
    for remark, c in fill_stats["result_breakdown"].most_common(10):
        w(f"  {c:>7,}  ·  {remark}")
    w("")
    w(bar)
    w("  Output: alloc sheet only · S/U/V/W written as static values")
    w("  T (VISITATION THRU) untouched · VISIT DATE formatted MM/DD/YYYY")
    w(bar)
    return "\n".join(lines)


# ----------------------------------------------------------------------
# PAYMENTS ON NEW ENDO & PTP pipeline
# ----------------------------------------------------------------------
PTP_PREFIXES = ("PTP NEW NEGO", "PTP OLD", "PTP EPA")
CONFIRMED_LIKE = ("CONFIRMED", "PAYMENT", "PAID")

# VOLARE row indices (7 computed cols + 52 DRR cols)
V_CONCA, V_RFD = 0, 6
V_DATE, V_ACCT, V_CARD = 7, 10, 11
V_STATUS, V_REMARK, V_REMBY = 15, 16, 17
V_PTP_AMT, V_PTP_DATE = 27, 29
V_CLAIM_AMT, V_CLAIM_DATE = 30, 31


def _nonzero(v):
    return v is not None and v != 0 and str(v).strip() not in ("", "0")


def fix_claim_date(v):
    """CONFIRMED Claim Paid Date fix -> real MM/DD/YYYY dates:
    - datetimes with day <= 12 were Excel-swapped -> swap back
    - 'dd/mm/yyyy' text (days 13-31 Excel could not misread) -> parse"""
    if isinstance(v, datetime):
        return fix_drr_date(v)
    if isinstance(v, str):
        m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*", v)
        if m:
            dd, mm, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
            try:
                return datetime(yy, mm, dd)
            except ValueError:
                return v
    return v


def extract_payments(vol_sorted, notes_af, notes_ux):
    """Walks the SORTED VOLARE and produces:
    - CONFIRMED rows (Claim Paid Amount <> 0 AND status in NOTES U)
    - PTP rows (PTP Amount <> 0 AND status starts with PTP NEW NEGO /
      PTP OLD / PTP EPA AND status in NOTES A)
    - total payment per account (SUMIF of Claim Paid Amount by Card No.)
    - first-match maps for the ALLOC lookups
    - flags for statuses missing from NOTES (surfaced in the report)"""
    conf_rows, ptp_rows = [], []
    conf_first, ptp_first = {}, {}
    total_pay = Counter()
    conf_missing = Counter()
    ptp_missing = Counter()

    for r in vol_sorted:
        card = r[V_CARD]
        amt = r[V_CLAIM_AMT]
        if isinstance(amt, (int, float)):
            total_pay[card] += amt

        acct, status = r[V_ACCT], r[V_STATUS]
        s = str(status) if status is not None else ""

        if _nonzero(r[V_CLAIM_AMT]):
            if status in notes_ux:
                conf_rows.append((acct, r[V_CONCA], status, r[V_REMARK],
                                  r[V_REMBY], r[V_CLAIM_AMT],
                                  fix_claim_date(r[V_CLAIM_DATE])))
                if acct not in conf_first:
                    conf_first[acct] = status
            elif s.upper().startswith(CONFIRMED_LIKE):
                conf_missing[status] += 1

        if _nonzero(r[V_PTP_AMT]) and s.startswith(PTP_PREFIXES):
            if status in notes_af:
                ptp_rows.append((acct, status, r[V_REMARK], r[V_REMBY],
                                 r[V_PTP_AMT], r[V_PTP_DATE], r[V_DATE]))
                if acct not in ptp_first:
                    ptp_first[acct] = (status, r[V_REMARK], r[V_REMBY],
                                       r[V_PTP_AMT], r[V_PTP_DATE], r[V_DATE])
            else:
                ptp_missing[status] += 1

    return {
        "conf_rows": conf_rows, "ptp_rows": ptp_rows,
        "conf_first": conf_first, "ptp_first": ptp_first,
        "total_pay": total_pay,
        "conf_missing": conf_missing, "ptp_missing": ptp_missing,
    }


def fill_payments(alloc_bytes, pay, volare_index, notes_ux, sheet_keywords,
                  include_ref_sheets=True, progress_cb=None, vba=False):
    """Fills S..AA in the ALLOC PAYMENTS AND PTP sheet as static values.
    S/T/U from CONFIRMED, V = RFD (Call outs first match), W..AA from PTP.
    Optionally writes generated CONFIRMED and PTP sheets into the output."""
    _precheck_alloc(alloc_bytes, sheet_keywords)
    wb = openpyxl.load_workbook(io.BytesIO(alloc_bytes), keep_vba=vba)
    ws = None
    for name in wb.sheetnames:
        n = norm_header(name)
        if any(k in n for k in sheet_keywords):
            ws = wb[name]
            break
    if ws is None:
        ws = wb.worksheets[0]

    header_row = None
    for r in range(1, 11):
        if norm_header(ws.cell(row=r, column=1).value) == "ch code":
            header_row = r
            break
    if header_row is None:
        names = list(wb.sheetnames)
        wb.close()
        raise ValueError(
            f"'{ws.title}' doesn't look like an ALLOC sheet (no 'ch code' "
            f"header found). Sheets in this file: {', '.join(names)}. "
            "Check that the correct alloc file is in the second upload slot."
        )

    COLS = tuple(range(19, 28))  # S..AA
    for r in range(1, header_row):
        for c in COLS:
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = None

    conf_first, ptp_first = pay["conf_first"], pay["ptp_first"]
    total_pay = pay["total_pay"]

    filled = with_pay = with_ptp = with_both = 0
    total_collected = 0.0
    ptp_amount_sum = 0.0
    paytype_breakdown = Counter()
    ptpstat_breakdown = Counter()
    min_date = max_date = None

    for r in range(header_row + 1, ws.max_row + 1):
        acct = ws.cell(row=r, column=4).value  # D = ACCOUNT NUMBER
        cells = {c: ws.cell(row=r, column=c) for c in COLS}

        if acct is None or str(acct).strip() == "":
            for cell in cells.values():
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = None
            continue

        filled += 1
        has_pay = acct in conf_first
        has_ptp = acct in ptp_first

        if has_pay:
            status = conf_first[acct]
            ux = notes_ux.get(status, (None, None, None))
            tp = total_pay.get(acct, 0)
            cells[19].value = tp                       # S TOTAL PAYMENT
            cells[20].value = ux[0]                    # T PAYMENT TYPE
            cells[21].value = ux[1]                    # U RECEIVED THRU
            if isinstance(tp, (int, float)):
                total_collected += tp
            paytype_breakdown[str(ux[0])] += 1
            with_pay += 1
        else:
            cells[19].value = None
            cells[20].value = None
            cells[21].value = None

        m = volare_index.get(f"call outs{acct}".lower())
        rfd = m[V_RFD] if m else None
        cells[22].value = None if (rfd is None or rfd == NA) else rfd  # V RFD

        if has_ptp:
            st2, rem2, rb2, amt2, paid2, sdate2 = ptp_first[acct]
            cells[23].value = sdate2                   # W NEGO DATE
            if isinstance(sdate2, datetime):
                cells[23].number_format = "MM/DD/YYYY"
                if min_date is None or sdate2 < min_date:
                    min_date = sdate2
                if max_date is None or sdate2 > max_date:
                    max_date = sdate2
            cells[24].value = st2                      # X STATUS
            cells[25].value = amt2                     # Y AMOUNT
            cells[26].value = paid2                    # Z PTP DATE (as-is)
            if isinstance(paid2, datetime):
                cells[26].number_format = "MM/DD/YYYY"
            cells[27].value = rb2                      # AA AGENT
            if isinstance(amt2, (int, float)):
                ptp_amount_sum += amt2
            ptpstat_breakdown[str(st2)] += 1
            with_ptp += 1
        else:
            cells[23].value = None
            cells[24].value = "-"                      # X fallback
            cells[25].value = None
            cells[26].value = None
            cells[27].value = None

        if has_pay and has_ptp:
            with_both += 1
        if progress_cb and filled % 2000 == 0:
            progress_cb(filled)

    # keep the alloc sheet + NOTES, drop everything else
    for name in list(wb.sheetnames):
        if wb[name] is not ws and norm_header(name) != "notes":
            del wb[name]

    # optionally add generated CONFIRMED / PTP reference sheets
    if include_ref_sheets:
        wsc = wb.create_sheet("CONFIRMED")
        wsc.append(["Account No.", "CONCA", "Status", "Remark", "Remark By",
                    "Claim Paid Amount", "Claim Paid Date", "total payment",
                    "FOR EPA \n(COMPLYING OR DEFAULTED)",
                    "FOR EPA \n(COMPLYING OR DEFAULTED)",
                    "SOURCE OF PYT \n(CALL OUTS,SKIPTRACE,LS,SMS,FIELD)"])
        for (acct, conca, status, remark, remby, amt, cdate) in pay["conf_rows"]:
            ux = notes_ux.get(status, (None, None, None))
            j = ux[2] if ux[2] is not None else 0   # Excel VLOOKUP: blank -> 0
            wsc.append([acct, conca, status, remark, remby, amt, cdate,
                        total_pay.get(acct, 0), ux[0], j, ux[1]])
            if isinstance(cdate, datetime):
                wsc.cell(row=wsc.max_row, column=7).number_format = "MM/DD/YYYY"

        wsp = wb.create_sheet("PTP")
        wsp.append(["Account No.", "Status", "Remark", "Remark By",
                    "PTP Amount", "PTP Paid Date", "STATUS DATE"])
        for (acct, status, remark, remby, amt, paid, sdate) in pay["ptp_rows"]:
            wsp.append([acct, status, remark, remby, amt, paid, sdate])
            if isinstance(sdate, datetime):
                wsp.cell(row=wsp.max_row, column=7).number_format = "MM/DD/YYYY"
            if isinstance(paid, datetime):
                wsp.cell(row=wsp.max_row, column=6).number_format = "MM/DD/YYYY"

    out = io.BytesIO()
    wb.save(out)
    wb.close()
    out.seek(0)
    return out.getvalue(), {
        "filled": filled,
        "with_pay": with_pay, "with_ptp": with_ptp, "with_both": with_both,
        "no_effort": filled - with_pay - with_ptp + with_both,
        "pos": with_pay, "neg": with_ptp,
        "total_collected": total_collected, "ptp_amount_sum": ptp_amount_sum,
        "confirmed_rows": len(pay["conf_rows"]), "ptp_rows": len(pay["ptp_rows"]),
        "result_breakdown": ptpstat_breakdown,
        "status_breakdown": paytype_breakdown,
        "min_date": min_date, "max_date": max_date,
    }


def build_payments_report(stats, pay, fill_stats, run_ts, drr_name, alloc_name,
                          notes_name, elapsed, mod, include_ref_sheets):
    lines = []
    w = lines.append
    bar = "=" * 62
    w(bar)
    w("  ALLOC AUTOMATION REMASTERD  ·  STATUS REPORT")
    w(f"  Module: {mod['name']}")
    w(bar)
    w(f"  Run date/time    : {run_ts:%m/%d/%Y %I:%M %p}")
    w(f"  Processing time  : {elapsed:.1f} seconds")
    w(f"  Raw DRR file     : {drr_name}")
    w(f"  Alloc file       : {alloc_name}")
    w(f"  NOTES key source : {notes_name}")
    w("")
    w("  STAGE 1 · DRR CLEANING / VOLARE STATUS BUILD")
    w("  " + "-" * 50)
    w(f"  DRR rows processed          : {stats['total_rows']:,}")
    for sh in stats.get("sheets_used", []):
        w(f"      - {sh['name']}: {sh['rows']:,} rows · {sh['swap']}")
    w(f"  Dates fixed (day/month swap): {stats['date_swapped']:,}")
    w(f"  Missing DRR headers         : "
      + (", ".join(stats['missing_headers']) if stats['missing_headers'] else "None"))
    unmapped_total = sum(stats["unmapped_status"].values())
    w(f"  Statuses missing from NOTES : {unmapped_total:,} rows "
      f"({len(stats['unmapped_status'])} distinct)")
    w("")
    w("  STAGE 2 · CONFIRMED EXTRACTION (Claim Paid Amount <> 0)")
    w("  " + "-" * 50)
    w(f"  CONFIRMED rows extracted    : {len(pay['conf_rows']):,}")
    w(f"  Unique accounts w/ payment  : {len(pay['conf_first']):,}")
    if pay["conf_missing"]:
        w("")
        w("  ⚠⚠  ATTENTION — CONFIRMED-TYPE STATUSES *NOT* IN NOTES (col U)")
        w("  ⚠⚠  These rows have a Claim Paid Amount but were EXCLUDED.")
        w("  ⚠⚠  Add them to the NOTES sheet and re-run:")
        for s, c in pay["conf_missing"].most_common():
            w(f"      >>> {s}  ({c:,} rows)")
    else:
        w("  Confirmed statuses missing from NOTES : None ✔")
    w("")
    w("  STAGE 3 · PTP EXTRACTION (PTP Amount <> 0; PTP NEW NEGO /")
    w("            PTP OLD / PTP EPA)")
    w("  " + "-" * 50)
    w(f"  PTP rows extracted          : {len(pay['ptp_rows']):,}")
    w(f"  Unique accounts w/ PTP      : {len(pay['ptp_first']):,}")
    if pay["ptp_missing"]:
        w("")
        w("  ⚠⚠  ATTENTION — PTP-TYPE STATUSES *NOT* IN NOTES (col A)")
        w("  ⚠⚠  These rows have a PTP Amount but were EXCLUDED.")
        w("  ⚠⚠  Add them to the NOTES sheet and re-run:")
        for s, c in pay["ptp_missing"].most_common():
            w(f"      >>> {s}  ({c:,} rows)")
    else:
        w("  PTP statuses missing from NOTES : None ✔")
    w("")
    w(f"  STAGE 4 · {mod['name']} FILL (S..AA)")
    w("  " + "-" * 50)
    w(f"  Accounts filled             : {fill_stats['filled']:,}")
    w(f"  Accounts with PAYMENT       : {fill_stats['with_pay']:,}")
    w(f"  Accounts with PTP           : {fill_stats['with_ptp']:,}")
    w(f"  Accounts with BOTH          : {fill_stats['with_both']:,}")
    w(f"  Total collected (S column)  : {fill_stats['total_collected']:,.2f}")
    w(f"  Total PTP amount (Y column) : {fill_stats['ptp_amount_sum']:,.2f}")
    if fill_stats["min_date"] and fill_stats["max_date"]:
        w(f"  {mod['date_label']:<28}: "
          f"{fill_stats['min_date']:%m/%d/%Y} - {fill_stats['max_date']:%m/%d/%Y}")
    w("")
    w("  PAYMENT TYPE MIX (T column)")
    w("  " + "-" * 50)
    for k, c in fill_stats["status_breakdown"].most_common():
        w(f"  {c:>7,}  ·  {k}")
    w("")
    w("  TOP 10 PTP STATUSES (X column)")
    w("  " + "-" * 50)
    for k, c in fill_stats["result_breakdown"].most_common(10):
        w(f"  {c:>7,}  ·  {k}")
    w("")
    w(bar)
    w("  Output: alloc sheet with S..AA as static values"
      + (" + generated CONFIRMED & PTP sheets" if include_ref_sheets else ""))
    w("  NEGO DATE / dates formatted MM/DD/YYYY · PTP Paid Date kept as-is")
    w(bar)
    return "\n".join(lines)


# ----------------------------------------------------------------------
# Status report (text) -- generated from run stats, logic untouched
# ----------------------------------------------------------------------
def build_status_report(stats, fill_stats, run_ts, drr_name, co_name, notes_name, elapsed, mod=None):
    lines = []
    w = lines.append
    bar = "=" * 62
    w(bar)
    w("  ALLOC AUTOMATION REMASTERD  ·  STATUS REPORT")
    w(f"  Module: {mod['name'] if mod else 'FOR CALL OUTS'}")
    w(bar)
    w(f"  Run date/time    : {run_ts:%m/%d/%Y %I:%M %p}")
    w(f"  Processing time  : {elapsed:.1f} seconds")
    w(f"  Raw DRR file     : {drr_name}")
    w(f"  Alloc file       : {co_name}")
    w(f"  NOTES key source : {notes_name}")
    w("")
    w("  STAGE 1 · DRR CLEANING / VOLARE STATUS BUILD")
    w("  " + "-" * 50)
    w(f"  DRR rows processed          : {stats['total_rows']:,}")
    for sh in stats.get("sheets_used", []):
        w(f"      - {sh['name']}: {sh['rows']:,} rows · {sh['swap']}")
    if stats.get("sheets_skipped"):
        w(f"      - skipped (not DRR): {', '.join(map(str, stats['sheets_skipped']))}")
    w(f"  Dates fixed (day/month swap): {stats['date_swapped']:,}")
    w(f"  Account length distribution : "
      + ", ".join(f"{k}-digit x {v:,}" for k, v in sorted(stats['acct_lens'].items())))
    w(f"  Missing DRR headers         : "
      + (", ".join(stats['missing_headers']) if stats['missing_headers'] else "None"))
    w(f"  Extra DRR headers ignored   : "
      + (", ".join(map(str, stats['extra_headers'])) if stats['extra_headers'] else "None"))
    unmapped_total = sum(stats["unmapped_status"].values())
    w(f"  Statuses missing from NOTES : {unmapped_total:,} rows "
      f"({len(stats['unmapped_status'])} distinct)")
    if stats["unmapped_status"]:
        for s, c in stats["unmapped_status"].most_common(10):
            w(f"      - {s}  ({c:,} rows)")
    md = stats.get("month_dist") or {}
    if md:
        w("  Month distribution          : "
          + ", ".join(f"{m}: {c:,}" for m, c in sorted(md.items())))
    w("")
    w(f"  STAGE 2 · {mod['name'] if mod else 'CALL OUTS'} FILL (S/T/U)")
    w("  " + "-" * 50)
    filled = fill_stats["filled"] or 1
    w(f"  Accounts filled             : {fill_stats['filled']:,}")
    w(f"  POS                         : {fill_stats['pos']:,}"
      f"  ({fill_stats['pos'] / filled * 100:.1f}%)")
    w(f"  NEG                         : {fill_stats['neg']:,}"
      f"  ({fill_stats['neg'] / filled * 100:.1f}%)")
    w(f"  NO EFFORT                   : {fill_stats['no_effort']:,}"
      f"  ({fill_stats['no_effort'] / filled * 100:.1f}%)")
    other = fill_stats["filled"] - fill_stats["pos"] - fill_stats["neg"] - fill_stats["no_effort"]
    if other:
        w(f"  Other statuses              : {other:,}")
    if fill_stats["min_date"] and fill_stats["max_date"]:
        w(f"  {(mod['date_label'] if mod else 'Call date range'):<28}: "
          f"{fill_stats['min_date']:%m/%d/%Y} - {fill_stats['max_date']:%m/%d/%Y}")
    w("")
    w("  TOP 10 RESULTS (REMARKS)")
    w("  " + "-" * 50)
    for remark, c in fill_stats["result_breakdown"].most_common(10):
        w(f"  {c:>7,}  ·  {remark}")
    w("")
    w(bar)
    w("  Output: CALL OUTS sheet only · S/T/U written as static values")
    w("  Formatting preserved · CALL DATE formatted MM/DD/YYYY")
    w(bar)
    return "\n".join(lines)


# ----------------------------------------------------------------------
# Module registry -- one pipeline, different lookup keys / labels
# ----------------------------------------------------------------------
MODULES = {
    "🚀 ALL-IN-ONE (Run everything)": {
        "id": "all", "pipeline": "all", "live": True, "icon": "🚀",
        "name": "ALL-IN-ONE",
        "desc": ("one run for the whole ALLOC review — the raw DRR is parsed "
                 "ONCE and shared by Call Outs, Skip Trace and Payments; the "
                 "FIELD RESULT feeds Visitation. Upload whichever files you "
                 "have; only those modules run. Outputs come individually and "
                 "as a single ZIP with a combined status report."),
    },
    "📞 For Call Outs": {
        "id": "call_outs",
        "icon": "📞",
        "name": "FOR CALL OUTS",
        "effort_key": "call outs",
        "sheet_keywords": ("call outs",),
        "alloc_label": "CALL OUTS file",
        "cols": "S/T/U — CALL DATE / CALL OUTS (POS & NEG) / CALL RESULT",
        "date_label": "Call date range",
        "out_file": "ALLOC FOR CALL OUTS - FINAL.xlsx",
        "report_file": "CALL OUTS STATUS REPORT",
        "desc": ("rebuilds VOLARE STATUS from the raw DRR + NOTES key, sorts by "
                 "hierarchy (smallest → largest, dates newest → oldest), and fills CALL DATE / CALL OUTS / "
                 "CALL RESULT as static values."),
        "live": True,
    },
    "🔎 For Skiptrace": {
        "id": "skiptrace",
        "icon": "🔎",
        "name": "FOR SKIP TRACE",
        "effort_key": "skiptrace",
        "sheet_keywords": ("skip trace", "skiptrace"),
        "alloc_label": "ALLOC SKIP TRACE file",
        "cols": "S/T/U — SKIPTRACING DATE / SKIPTRACE (POS, NEG & UNCONFIRMED) / SKIPTRACE REMARKS",
        "date_label": "Skiptrace date range",
        "out_file": "ALLOC FOR SKIP TRACE - FINAL.xlsx",
        "report_file": "SKIP TRACE STATUS REPORT",
        "desc": ("rebuilds VOLARE STATUS from the raw DRR + NOTES key, sorts by "
                 "hierarchy, and fills SKIPTRACING DATE / SKIPTRACE STATUS / "
                 "SKIPTRACE REMARKS as static values. NEWLY GATHERED INFO is left "
                 "untouched."),
        "live": True,
    },
    "🏠 For Visitation": {
        "id": "visitation",
        "icon": "🏠",
        "name": "FOR FIELD VISITATION",
        "pipeline": "field",
        "sheet_keywords": ("field visitation", "visitation", "visit"),
        "alloc_label": "ALLOC FIELD VISITATION file",
        "cols": "S/U/V/W — VISIT DATE / VISITATION RESULT / INFORMANT / VISITATION REMARKS (T untouched)",
        "date_label": "Visit date range",
        "out_file": "ALLOC FOR FIELD VISITATION - FINAL.xlsx",
        "report_file": "FIELD VISITATION STATUS REPORT",
        "desc": ("filters the raw FIELD RESULT (RESULT sheet) by MC, bank name, "
                 "and placement, maps statuses through the NOTES I:J key, and "
                 "fills VISIT DATE / VISITATION RESULT / INFORMANT / VISITATION "
                 "REMARKS as static values. VISITATION THRU is left untouched."),
        "live": True,
    },
    "💰 Payments on New Endo & PTP": {
        "id": "payments",
        "icon": "💰",
        "name": "PAYMENTS ON NEW ENDO & PTP",
        "pipeline": "payments",
        "effort_key": "call outs",
        "sheet_keywords": ("payments and ptp", "payments", "ptp"),
        "alloc_label": "ALLOC PAYMENTS AND PTP file",
        "cols": "S..AA — TOTAL PAYMENT / PAYMENT TYPE / RECEIVED THRU / RFD / NEGO DATE / STATUS / AMOUNT / PTP DATE / AGENT",
        "date_label": "Nego date range",
        "out_file": "ALLOC PAYMENTS AND PTP - FINAL.xlsx",
        "report_file": "PAYMENTS AND PTP STATUS REPORT",
        "desc": ("rebuilds VOLARE STATUS from the raw DRR + NOTES key, sorts by "
                 "hierarchy, extracts CONFIRMED payments (NOTES col U) and PTP "
                 "commitments (PTP NEW NEGO / PTP OLD / PTP EPA in NOTES col A), "
                 "then fills S..AA as static values. Statuses missing from NOTES "
                 "are flagged in the status report."),
        "live": True,
    },
}


def _swap_mode_from_label(label):
    if label.startswith("Force"):
        return "force"
    if label.startswith("No swap"):
        return "off"
    return "auto"


# ======================================================================
# STREAMLIT UI
# ======================================================================
st.set_page_config(
    page_title="ALLOC AUTOMATION REMASTERD",
    page_icon="📞",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      /* ------- base ------- */
      .stApp { background: radial-gradient(1200px 500px at 20% -10%, #14213d 0%, #0b0f19 55%) fixed; }
      .block-container { max-width: 1050px; padding-top: 1.2rem; }
      h1, h2, h3, h4 { color: #f8fafc !important; }
      p, label, .stMarkdown { color: #cbd5e1; }

      /* ------- hero banner ------- */
      .hero {
        background: linear-gradient(120deg, #1e3a8a 0%, #312e81 55%, #0f172a 100%);
        border: 1px solid #334155; border-radius: 18px;
        padding: 26px 32px; margin-bottom: 6px;
        box-shadow: 0 8px 30px rgba(0,0,0,.45);
      }
      .hero h1 { margin: 0; font-size: 1.9rem; letter-spacing: 1.5px; }
      .hero .sub { color: #93c5fd; font-size: .95rem; margin-top: 6px; }
      .hero .chip {
        display: inline-block; background: rgba(147,197,253,.12);
        border: 1px solid rgba(147,197,253,.35); color: #bfdbfe;
        border-radius: 999px; padding: 3px 14px; font-size: .75rem;
        letter-spacing: 2px; margin-bottom: 10px;
      }

      /* ------- step cards ------- */
      .stepcard {
        background: #0f172acc; border: 1px solid #1e293b; border-radius: 14px;
        padding: 18px 22px; margin: 14px 0 6px 0;
      }
      .stepcard .stepno {
        display:inline-flex; align-items:center; justify-content:center;
        width: 26px; height: 26px; border-radius: 8px;
        background: #2563eb; color: white; font-weight: 700; font-size: .85rem;
        margin-right: 10px;
      }
      .stepcard .steptitle { color:#f1f5f9; font-weight:600; font-size:1.05rem; }
      .stepcard .stepdesc { color:#94a3b8; font-size:.85rem; margin-top:4px; }

      /* ------- uploaders ------- */
      div[data-testid="stFileUploader"] {
        background:#111827; border:1px dashed #334155; border-radius:12px;
        padding: 10px 14px;
      }
      div[data-testid="stFileUploader"] label { color:#e2e8f0 !important; font-weight:600; }

      /* ------- metrics ------- */
      div[data-testid="stMetric"] {
        background: linear-gradient(160deg, #111827 0%, #0b1120 100%);
        border: 1px solid #1f2937; border-radius: 14px; padding: 14px 16px;
      }
      div[data-testid="stMetric"] label { color:#94a3b8 !important; }
      div[data-testid="stMetricValue"] { color:#f8fafc !important; }

      /* ------- buttons ------- */
      .stButton button[kind="primary"] {
        background: linear-gradient(90deg,#2563eb,#4f46e5) !important;
        border: 0 !important; border-radius: 12px !important;
        font-weight: 700 !important; letter-spacing: 1px;
        padding: .7rem 1rem !important;
        box-shadow: 0 6px 18px rgba(37,99,235,.35);
      }
      .stDownloadButton button {
        background: linear-gradient(90deg,#059669,#10b981) !important;
        color: white !important; border: 0 !important;
        border-radius: 12px !important; font-weight: 700 !important;
        letter-spacing: 1px; padding: .7rem 1rem !important;
        box-shadow: 0 6px 18px rgba(16,185,129,.3);
      }

      /* ------- report ------- */
      .reportbox {
        background:#0b1120; border:1px solid #1e293b; border-radius:14px;
        padding: 8px 14px;
      }
      .pillrow { margin: 4px 0 10px 0; }
      .pill {
        display:inline-block; border-radius:999px; padding:4px 14px;
        font-size:.8rem; font-weight:600; margin-right:8px; margin-bottom:6px;
      }
      .pill.pos { background:rgba(16,185,129,.15); color:#34d399; border:1px solid rgba(16,185,129,.4);}
      .pill.neg { background:rgba(239,68,68,.14); color:#f87171; border:1px solid rgba(239,68,68,.4);}
      .pill.noeff { background:rgba(148,163,184,.14); color:#cbd5e1; border:1px solid rgba(148,163,184,.4);}

      /* sidebar */
      section[data-testid="stSidebar"] { background:#0b1120; border-right:1px solid #1e293b; }
      section[data-testid="stSidebar"] .stRadio label { color:#e2e8f0; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------- sidebar : module selector ----------------
with st.sidebar:
    st.markdown("### 🗂️ ALLOC MODULES")
    mode = st.radio(
        "Select automation",
        list(MODULES.keys()),
        label_visibility="collapsed",
    )
    MOD = MODULES[mode]
    st.markdown("---")
    st.markdown(
        "<div style='color:#64748b;font-size:.78rem'>"
        "MADRECO SUITE · ALLOC AUTOMATION REMASTERD<br>"
        "All-in-one + 4 modules · v3.0</div>",
        unsafe_allow_html=True,
    )

# ---------------- hero ----------------
st.markdown(
    """
    <div class="hero">
      <div class="chip">MADRECO SUITE</div>
      <h1>ALLOC AUTOMATION REMASTERD</h1>
      <div class="sub">{icon} {name} — {desc} Output = alloc file only, formatting preserved.</div>
    </div>
    """.format(icon=MOD.get("icon", "🚧"), name=MOD.get("name", mode),
               desc=MOD.get("desc", "")),
    unsafe_allow_html=True,
)

if not MOD.get("live"):
    st.info("🚧 This module isn't built yet — the other three are live. Payments on New Endo & PTP is coming next.")
    st.stop()

# ================= ALL-IN-ONE MODE =================
if MOD.get("pipeline") == "all":
    st.markdown(
        '<div class="stepcard"><span class="stepno">1</span>'
        '<span class="steptitle">Upload files</span>'
        '<div class="stepdesc">Upload what you have — only the modules with '
        'files provided will run. VOLARE STATUS and the FIELD reference are '
        'rebuilt internally.</div></div>',
        unsafe_allow_html=True,
    )
    r1c1, r1c2 = st.columns(2)
    with r1c1:
        a_drr = st.file_uploader("📄 RAW DRR — for Call Outs / Skip Trace / Payments",
                                 type=["xlsx", "xlsm", "xlsb", "xls", "csv"], key="all_drr")
    with r1c2:
        a_fr = st.file_uploader("📄 RAW FIELD RESULT — for Visitation",
                                type=["xlsx", "xlsm", "xlsb", "xls", "csv"], key="all_fr")
    r2c1, r2c2 = st.columns(2)
    with r2c1:
        a_co = st.file_uploader("📘 CALL OUTS file (xlsx/xlsm)", type=["xlsx", "xlsm"], key="all_co")
        a_vis = st.file_uploader("📘 ALLOC FIELD VISITATION file (xlsx/xlsm)", type=["xlsx", "xlsm"], key="all_vis")
    with r2c2:
        a_sk = st.file_uploader("📘 ALLOC SKIP TRACE file (xlsx/xlsm)", type=["xlsx", "xlsm"], key="all_sk")
        a_pay = st.file_uploader("📘 ALLOC PAYMENTS AND PTP file (xlsx/xlsm)", type=["xlsx", "xlsm"], key="all_pay")
    a_notes = st.file_uploader(
        "🔑 NOTES key file — optional if any uploaded alloc workbook has a NOTES sheet",
        type=["xlsx", "xlsm", "xlsb", "xls", "csv"], key="all_notes")

    with st.expander("⚙️ Visitation filter settings", expanded=False):
        f1, f2 = st.columns(2)
        with f1:
            all_mc = st.text_input("MC", value="MC6", key="all_flt_mc")
        with f2:
            all_place = st.text_input("PLACEMENT contains", value="RECO 1", key="all_flt_place")
        all_banks = st.text_input("NEW BANK NAME (comma-separated)",
                                  value="BPI PL RECO LUZ PL, BPI CARDS RECO LUZ CC",
                                  key="all_flt_banks")
    all_inc_ref = st.toggle("Include generated CONFIRMED & PTP sheets in the Payments output",
                            value=True, key="all_inc_ref")
    all_inc_vol = st.toggle("Include the VOLARE STATUS file (final working order) in the ZIP",
                            value=True, key="all_inc_vol")
    all_date_dir = st.radio(
        "Step 1 date sort (before SORT smallest → largest)",
        ["Newest to Oldest — latest effort wins ties",
         "Oldest to Newest — earliest effort wins ties"],
        horizontal=True, key="all_date_dir",
    )
    all_swap_mode = st.radio(
        "Date day/month swap fix",
        ["Auto-detect per sheet (recommended)",
         "Force swap (dd/mm misread export)",
         "No swap (dates already correct)"],
        horizontal=True, key="all_swap_mode",
    )

    will_run = []
    if a_drr and a_co: will_run.append("Call Outs")
    if a_drr and a_sk: will_run.append("Skip Trace")
    if a_fr and a_vis: will_run.append("Visitation")
    if a_drr and a_pay: will_run.append("Payments & PTP")
    pills = "".join(
        f"<span class='pill {'pos' if m in will_run else 'noeff'}'>"
        f"{'✓' if m in will_run else '•'} {m}</span>"
        for m in ["Call Outs", "Skip Trace", "Visitation", "Payments & PTP"])
    st.markdown(f"<div class='pillrow'>{pills}</div>", unsafe_allow_html=True)

    st.markdown(
        '<div class="stepcard"><span class="stepno">2</span>'
        '<span class="steptitle">Run everything</span>'
        '<div class="stepdesc">The DRR is cleaned and VOLARE STATUS is built '
        'ONCE, then shared across all DRR-based modules — much faster than '
        'running them one by one.</div></div>',
        unsafe_allow_html=True,
    )
    run_all = st.button("🚀  RUN ALL-IN-ONE", type="primary",
                        use_container_width=True, disabled=not will_run)
    if not will_run:
        st.caption("Upload at least one raw file + its matching alloc file to enable the run.")

    if run_all:
        import time as _time
        import zipfile as _zip
        t_all = _time.time()
        results = {}     # name -> (out_bytes, fill_stats, report, elapsed)
        errors = {}
        prog = st.progress(0, text="Reading NOTES key…")
        try:
            notes_bytes = None
            notes_fn = "notes.xlsx"
            if a_notes:
                notes_bytes = a_notes.getvalue(); notes_src = a_notes.name; notes_fn = a_notes.name
            else:
                for f in (a_co, a_sk, a_pay, a_vis):
                    if f and "notes" in [norm_header(n) for n in peek_sheetnames(f.getvalue(), f.name)]:
                        notes_bytes = f.getvalue()
                        notes_src = f"{f.name} (NOTES sheet)"
                        notes_fn = f.name
                        break
            if notes_bytes is None:
                st.error("No NOTES key found in any upload. Add a NOTES file.")
                st.stop()
            notes_af, notes_pq, notes_ij, notes_ux = _apply_notes_patch(
                *load_notes(notes_bytes, notes_fn))
            st.session_state["notes_srcfile"] = (notes_bytes, notes_fn)

            vol_sorted = index = None
            stats_drr = None
            need_drr = bool(a_drr and (a_co or a_sk or a_pay))
            if need_drr:
                prog.progress(5, text="Cleaning DRR + building VOLARE STATUS (shared, one pass)…")

                def drr_cb(n):
                    prog.progress(min(5 + int(n / 300000 * 40), 45),
                                  text=f"Cleaning DRR + building VOLARE STATUS… {n:,} rows")

                vol_rows, stats_drr = build_volare(
                    a_drr.getvalue(), notes_af, notes_pq, drr_cb, a_drr.name,
                    swap_mode=_swap_mode_from_label(
                        st.session_state.get("all_swap_mode", "Auto")))
                nf_all = st.session_state.get("all_date_dir", "N").startswith("Newest")
                prog.progress(48, text="Step 1: Date sort · Step 2: SORT smallest → largest…")
                vol_sorted = sort_volare(vol_rows, newest_first=nf_all)
                index = build_lookup_index(vol_sorted)
                del vol_rows

            run_ts = datetime.now()
            steps = [m for m in ["Call Outs", "Skip Trace", "Payments & PTP", "Visitation"] if m in will_run]
            base = 52
            per = int(45 / max(len(steps), 1))

            for si, mname in enumerate(steps):
                p0 = base + si * per
                prog.progress(p0, text=f"Running {mname}…")
                t0 = _time.time()
                try:
                    if mname == "Call Outs":
                        mod = MODULES["📞 For Call Outs"]
                        out, fst = fill_alloc(a_co.getvalue(), index, "call outs", mod["sheet_keywords"], vba=a_co.name.lower().endswith(".xlsm"))
                        rep = build_status_report(stats_drr, fst, run_ts, a_drr.name,
                                                  a_co.name, notes_src, _time.time() - t0, mod)
                    elif mname == "Skip Trace":
                        mod = MODULES["🔎 For Skiptrace"]
                        out, fst = fill_alloc(a_sk.getvalue(), index, "skiptrace", mod["sheet_keywords"], vba=a_sk.name.lower().endswith(".xlsm"))
                        rep = build_status_report(stats_drr, fst, run_ts, a_drr.name,
                                                  a_sk.name, notes_src, _time.time() - t0, mod)
                    elif mname == "Payments & PTP":
                        mod = MODULES["💰 Payments on New Endo & PTP"]
                        pay = extract_payments(vol_sorted, notes_af, notes_ux)
                        out, fst = fill_payments(a_pay.getvalue(), pay, index, notes_ux,
                                                 mod["sheet_keywords"],
                                                 include_ref_sheets=all_inc_ref,
                                                 vba=a_pay.name.lower().endswith(".xlsm"))
                        rep = build_payments_report(stats_drr, pay, fst, run_ts, a_drr.name,
                                                    a_pay.name, notes_src, _time.time() - t0,
                                                    mod, all_inc_ref)
                        fst["_flags"] = (dict(pay["conf_missing"]), dict(pay["ptp_missing"]))
                    else:  # Visitation
                        mod = MODULES["🏠 For Visitation"]
                        fr_sheets = [norm_header(n) for n in peek_sheetnames(a_fr.getvalue(), a_fr.name)]
                        vis_sheets = [norm_header(n) for n in peek_sheetnames(a_vis.getvalue(), a_vis.name)]
                        if "result" not in fr_sheets and "result" in vis_sheets:
                            raise ValueError("Visitation uploads look swapped — the RESULT "
                                             "sheet is in the alloc slot.")
                        banks = [b for b in st.session_state.get(
                            "all_flt_banks", "BPI PL RECO LUZ PL, BPI CARDS RECO LUZ CC").split(",")]
                        fidx, fstats, frows = build_field_index(
                            a_fr.getvalue(), notes_ij,
                            st.session_state.get("all_flt_mc", "MC6"), banks,
                            st.session_state.get("all_flt_place", "RECO 1"),
                            None, a_fr.name)
                        out, fst = fill_visitation(a_vis.getvalue(), fidx, mod["sheet_keywords"], frows, vba=a_vis.name.lower().endswith(".xlsm"))
                        rep = build_visit_report(fstats, fst, run_ts, a_fr.name,
                                                 a_vis.name, notes_src, _time.time() - t0, mod)
                    results[mname] = (out, fst, rep, _time.time() - t0, mod["out_file"])
                except Exception as me:
                    errors[mname] = str(me)
                prog.progress(min(p0 + per, 97), text=f"{mname} done.")

            # combined report + zip
            bar = "=" * 62
            combined = [bar,
                        "  ALLOC AUTOMATION REMASTERD  ·  ALL-IN-ONE COMBINED REPORT",
                        f"  Run: {run_ts:%m/%d/%Y %I:%M %p}  ·  "
                        f"Total time: {_time.time() - t_all:.1f}s",
                        f"  Modules run: {', '.join(results) if results else 'None'}",
                        f"  Modules failed: {', '.join(errors) if errors else 'None'}",
                        bar, ""]
            for mname, (_, _, rep, _, _) in results.items():
                combined.append(rep); combined.append("")
            for mname, msg in errors.items():
                combined.append(f"  ✖ {mname} FAILED: {msg}"); combined.append("")
            combined_txt = "\n".join(combined)

            vol_bytes_all = None
            if vol_sorted is not None and st.session_state.get("all_inc_vol", True):
                prog.progress(98, text="Writing VOLARE STATUS file…")
                vol_bytes_all = build_volare_file(vol_sorted)

            zbuf = io.BytesIO()
            with _zip.ZipFile(zbuf, "w", _zip.ZIP_DEFLATED) as zf:
                for mname, (out, _, rep, _, fname) in results.items():
                    zf.writestr(fname, out)
                if vol_bytes_all:
                    zf.writestr("VOLARE STATUS - FINAL.xlsx", vol_bytes_all)
                zf.writestr(f"ALL-IN-ONE STATUS REPORT {run_ts:%m-%d-%Y}.txt", combined_txt)
            zbuf.seek(0)

            st.session_state["all_results"] = results
            st.session_state["all_errors"] = errors
            st.session_state["all_zip"] = zbuf.getvalue()
            st.session_state["all_report"] = combined_txt
            st.session_state["all_run_ts"] = run_ts
            st.session_state["all_elapsed"] = _time.time() - t_all
            prog.progress(100, text=f"All done ✔  ({_time.time() - t_all:.1f}s)")
        except Exception as e:
            prog.empty()
            st.error(f"Processing failed: {e}")
            st.stop()

    if "all_results" in st.session_state:
        results = st.session_state["all_results"]
        errors = st.session_state["all_errors"]
        run_ts = st.session_state["all_run_ts"]
        st.markdown(
            '<div class="stepcard"><span class="stepno">3</span>'
            '<span class="steptitle">Results</span>'
            f'<div class="stepdesc">Run completed {run_ts:%m/%d/%Y %I:%M %p} in '
            f'{st.session_state["all_elapsed"]:.1f}s total.</div></div>',
            unsafe_allow_html=True,
        )
        for mname, (out, fst, rep, took, fname) in results.items():
            with st.container(border=True):
                h1, h2 = st.columns([3, 1])
                h1.markdown(f"**{'📞' if mname=='Call Outs' else '🔎' if mname=='Skip Trace' else '🏠' if mname=='Visitation' else '💰'} {mname}** · {took:.1f}s")
                h2.download_button("⬇ Download", data=out, file_name=fname,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="dl_" + mname, use_container_width=True)
                if mname == "Payments & PTP":
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Accounts", f"{fst['filled']:,}")
                    m2.metric("With payment", f"{fst['with_pay']:,}")
                    m3.metric("With PTP", f"{fst['with_ptp']:,}")
                    m4.metric("Collected", f"{fst['total_collected']:,.0f}")
                    flags = fst.get("_flags", ({}, {}))
                    if flags[0] or flags[1]:
                        st.error("⚠ Statuses missing from NOTES were excluded — "
                                 "see the combined report.")
                else:
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Accounts", f"{fst['filled']:,}")
                    m2.metric("🟢 POS", f"{fst['pos']:,}")
                    m3.metric("🔴 NEG", f"{fst['neg']:,}")
                    m4.metric("⚪ No result", f"{fst['no_effort']:,}")
        for mname, msg in errors.items():
            st.error(f"✖ {mname} failed: {msg}")

        st.markdown(
            '<div class="stepcard"><span class="stepno">4</span>'
            '<span class="steptitle">Download everything</span>'
            '<div class="stepdesc">All outputs plus the combined status report '
            'in one ZIP.</div></div>',
            unsafe_allow_html=True,
        )
        z1, z2 = st.columns([2, 1])
        with z1:
            st.download_button("📦  DOWNLOAD ALL (ZIP)",
                               data=st.session_state["all_zip"],
                               file_name=f"ALLOC ALL-IN-ONE {run_ts:%m-%d-%Y}.zip",
                               mime="application/zip", use_container_width=True)
        with z2:
            st.download_button("🧾  COMBINED REPORT (.txt)",
                               data=st.session_state["all_report"],
                               file_name=f"ALL-IN-ONE STATUS REPORT {run_ts:%m-%d-%Y}.txt",
                               mime="text/plain", use_container_width=True)
        with st.expander("🧾 View combined status report"):
            st.code(st.session_state["all_report"], language=None)
    st.stop()

# ---------------- step 1 : uploads ----------------
st.markdown(
    '<div class="stepcard"><span class="stepno">1</span>'
    '<span class="steptitle">Upload files</span>'
    '<div class="stepdesc">VOLARE STATUS is rebuilt internally — no need to upload it. '
    f"For best speed, the {MOD['alloc_label']} should NOT contain a VOLARE STATUS sheet.</div></div>",
    unsafe_allow_html=True,
)

IS_FIELD = MOD.get("pipeline") == "field"
IS_PAY = MOD.get("pipeline") == "payments"
u1, u2 = st.columns(2)
with u1:
    if IS_FIELD:
        drr_file = st.file_uploader("📄 RAW FIELD RESULT file — must contain a 'RESULT' sheet",
                                    type=["xlsx", "xlsm", "xlsb", "xls", "csv"], key="fr_" + MOD["id"])
    else:
        drr_file = st.file_uploader("📄 RAW DRR file (xlsx/xlsm/xlsb/xls)", type=["xlsx", "xlsm", "xlsb", "xls", "csv"], key="drr_" + MOD["id"])
with u2:
    co_file = st.file_uploader(f"📘 {MOD['alloc_label']} (xlsx/xlsm — edited in place)", type=["xlsx", "xlsm"], key="co_" + MOD["id"])
notes_file = st.file_uploader(
    f"🔑 NOTES key file — optional if your {MOD['alloc_label']} workbook already has a NOTES sheet",
    type=["xlsx", "xlsm", "xlsb", "xls", "csv"], key="notes_" + MOD["id"],
)

if IS_FIELD:
    with st.expander("⚙️ Filter settings (defaults match the validated ALLOC review)", expanded=False):
        f1, f2 = st.columns(2)
        with f1:
            flt_mc = st.text_input("MC", value="MC6", key="flt_mc")
        with f2:
            flt_place = st.text_input("PLACEMENT contains", value="RECO 1", key="flt_place")
        flt_banks = st.text_input(
            "NEW BANK NAME (comma-separated)",
            value="BPI PL RECO LUZ PL, BPI CARDS RECO LUZ CC",
            key="flt_banks",
        )

if IS_PAY:
    inc_ref = st.toggle(
        "Include generated CONFIRMED & PTP sheets in the output file",
        value=True, key="inc_ref_sheets",
        help="The alloc sheet is always included. Turn this off for an "
             "alloc-only output like the other modules.",
    )

if not IS_FIELD:
    inc_vol = st.toggle(
        "Also generate the VOLARE STATUS file (final working order)",
        value=True, key="inc_vol_" + MOD["id"],
        help="Written as its own workbook because huge row counts can't fit "
             "inside the formatted alloc file. If the data exceeds Excel's "
             "1,048,576-row sheet limit, it continues onto VOLARE STATUS (2), "
             "(3)… in the same working order. Adds time on big files.",
    )
    date_dir = st.radio(
        "Step 1 date sort (before SORT smallest → largest)",
        ["Newest to Oldest — latest effort wins ties",
         "Oldest to Newest — earliest effort wins ties"],
        horizontal=True, key="date_dir_" + MOD["id"],
    )
    swap_mode_label = st.radio(
        "Date day/month swap fix",
        ["Auto-detect per sheet (recommended)",
         "Force swap (dd/mm misread export)",
         "No swap (dates already correct)"],
        horizontal=True, key="swap_mode_" + MOD["id"],
    )

ready = bool(drr_file and co_file)
st.markdown(
    f"<div class='pillrow'>"
    f"<span class='pill {'pos' if drr_file else 'noeff'}'>{'✓' if drr_file else '•'} {'RAW FIELD RESULT' if IS_FIELD else 'RAW DRR'}</span>"
    f"<span class='pill {'pos' if co_file else 'noeff'}'>{'✓' if co_file else '•'} {MOD['alloc_label'].upper()}</span>"
    f"<span class='pill {'pos' if notes_file else 'noeff'}'>{'✓ NOTES (file)' if notes_file else '• NOTES (from alloc workbook)'}</span>"
    f"</div>",
    unsafe_allow_html=True,
)

# ---------------- step 2 : run ----------------
st.markdown(
    '<div class="stepcard"><span class="stepno">2</span>'
    '<span class="steptitle">Run the automation</span>'
    '<div class="stepdesc">One click — cleaning, VOLARE build, hierarchy sort, and S/T/U fill '
    'all run in a single pass with live progress.</div></div>',
    unsafe_allow_html=True,
)

run = st.button(f"▶  RUN {MOD['name']} AUTOMATION", type="primary",
                use_container_width=True, disabled=not ready)
if not ready:
    st.caption(f"Upload the {'RAW FIELD RESULT' if IS_FIELD else 'RAW DRR'} and {MOD['alloc_label']} to enable the run button.")

if run:
    import time as _time
    t0 = _time.time()
    co_bytes = co_file.getvalue()
    notes_bytes = notes_file.getvalue() if notes_file else co_bytes
    notes_src = notes_file.name if notes_file else f"{co_file.name} (NOTES sheet)"

    prog = st.progress(0, text="Reading NOTES key…")
    try:
        notes_fn = notes_file.name if notes_file else co_file.name
        notes_af, notes_pq, notes_ij, notes_ux = _apply_notes_patch(
            *load_notes(notes_bytes, notes_fn))
        st.session_state["notes_srcfile"] = (notes_bytes, notes_fn)
        st.session_state["notes_tables"] = (dict(notes_af), dict(notes_pq),
                                            dict(notes_ux))
        if not notes_af:
            st.error(
                "No NOTES key found. Upload a NOTES file, or make sure your "
                "CALL OUTS workbook contains a 'NOTES' sheet."
            )
            st.stop()
        if IS_FIELD:
            # pre-flight: detect swapped uploads before heavy processing
            fr_sheets = [norm_header(n) for n in peek_sheetnames(drr_file.getvalue(), drr_file.name)]
            co_sheets = [norm_header(n) for n in peek_sheetnames(co_bytes, co_file.name)]
            if ("result" not in fr_sheets and "csv" not in fr_sheets
                    and "result" in co_sheets):
                st.error(
                    "🔁 The uploads look SWAPPED: the file in the second slot "
                    "contains the 'RESULT' sheet, while the first one doesn't. "
                    "Slot 1 = raw FIELD RESULT file · Slot 2 = ALLOC FIELD "
                    "VISITATION file. Please switch them and run again. "
                    "(Reminder: this module doesn't use the raw DRR at all.)"
                )
                st.stop()
            if not notes_ij:
                st.error("The NOTES key has no I:J field-status entries. "
                         "Check the NOTES sheet.")
                st.stop()
            prog.progress(10, text=f"NOTES loaded ({len(notes_ij)} field statuses). "
                                   "Filtering FIELD RESULT…")

            def fr_cb(n):
                prog.progress(min(10 + int(n / 120000 * 60), 70),
                              text=f"Filtering FIELD RESULT… {n:,} rows scanned")

            banks = [b for b in st.session_state.get(
                "flt_banks", "BPI PL RECO LUZ PL, BPI CARDS RECO LUZ CC").split(",")]
            index, stats, field_rows = build_field_index(
                drr_file.getvalue(), notes_ij,
                st.session_state.get("flt_mc", "MC6"),
                banks,
                st.session_state.get("flt_place", "RECO 1"),
                fr_cb, drr_file.name,
            )
            prog.progress(75, text=f"Filling {MOD['name']} S/U/V/W…")

            def co_cb(n):
                prog.progress(min(75 + int(n / 15000 * 22), 97),
                              text=f"Filling {MOD['name']}… {n:,} accounts")

            out_bytes, fill_stats = fill_visitation(
                co_bytes, index, MOD["sheet_keywords"], field_rows, co_cb,
                vba=co_file.name.lower().endswith(".xlsm"))
            elapsed = _time.time() - t0
            prog.progress(100, text=f"Done ✔  ({elapsed:.1f}s)")

            run_ts = datetime.now()
            report = build_visit_report(
                stats, fill_stats, run_ts,
                drr_file.name, co_file.name, notes_src, elapsed, MOD,
            )
        elif IS_PAY:
            if not notes_ux:
                st.error("The NOTES key has no U:X CONFIRMED entries. "
                         "Check the NOTES sheet.")
                st.stop()
            prog.progress(8, text=f"NOTES loaded ({len(notes_af)} statuses, "
                                  f"{len(notes_ux)} CONFIRMED entries). Cleaning raw DRR…")

            def drr_cb(n):
                prog.progress(min(8 + int(n / 300000 * 50), 60),
                              text=f"Cleaning DRR + building VOLARE STATUS… {n:,} rows")

            vol_rows, stats = build_volare(
                drr_file.getvalue(), notes_af, notes_pq, drr_cb, drr_file.name,
                swap_mode=_swap_mode_from_label(
                    st.session_state.get("swap_mode_" + MOD["id"], "Auto")))
            nf = st.session_state.get("date_dir_" + MOD["id"], "N").startswith("Newest")
            prog.progress(63, text="Step 1: Date sort · Step 2: SORT smallest → largest…")
            vol_sorted = sort_volare(vol_rows, newest_first=nf)
            index = build_lookup_index(vol_sorted)

            prog.progress(70, text="Extracting CONFIRMED payments and PTP…")
            pay = extract_payments(vol_sorted, notes_af, notes_ux)

            prog.progress(78, text=f"Filling {MOD['name']} S..AA…")

            def co_cb(n):
                prog.progress(min(78 + int(n / 15000 * 19), 97),
                              text=f"Filling {MOD['name']}… {n:,} accounts")

            inc = st.session_state.get("inc_ref_sheets", True)
            out_bytes, fill_stats = fill_payments(
                co_bytes, pay, index, notes_ux, MOD["sheet_keywords"],
                include_ref_sheets=inc, progress_cb=co_cb,
                vba=co_file.name.lower().endswith(".xlsm"))
            if st.session_state.get("inc_vol_" + MOD["id"], True):
                prog.progress(97, text="Writing VOLARE STATUS file…")
                st.session_state["vol_file_" + MOD["id"]] = build_volare_file(vol_sorted)
            elapsed = _time.time() - t0
            prog.progress(100, text=f"Done ✔  ({elapsed:.1f}s)")

            run_ts = datetime.now()
            report = build_payments_report(
                stats, pay, fill_stats, run_ts,
                drr_file.name, co_file.name, notes_src, elapsed, MOD, inc,
            )
            st.session_state["pay_flags_" + MOD["id"]] = (
                dict(pay["conf_missing"]), dict(pay["ptp_missing"]))
        else:
            prog.progress(10, text=f"NOTES loaded ({len(notes_af)} statuses, "
                                   f"{len(notes_pq)} RFD entries). Cleaning raw DRR…")

            def drr_cb(n):
                prog.progress(min(10 + int(n / 300000 * 55), 65),
                              text=f"Cleaning DRR + building VOLARE STATUS… {n:,} rows")

            vol_rows, stats = build_volare(
                drr_file.getvalue(), notes_af, notes_pq, drr_cb, drr_file.name,
                swap_mode=_swap_mode_from_label(
                    st.session_state.get("swap_mode_" + MOD["id"], "Auto")))
            nf = st.session_state.get("date_dir_" + MOD["id"], "N").startswith("Newest")
            prog.progress(68, text="Step 1: Date sort · Step 2: SORT smallest → largest…")

            vol_sorted = sort_volare(vol_rows, newest_first=nf)
            index = build_lookup_index(vol_sorted)
            prog.progress(75, text=f"Filling {MOD['name']} S/T/U…")

            def co_cb(n):
                prog.progress(min(75 + int(n / 15000 * 22), 97),
                              text=f"Filling {MOD['name']}… {n:,} accounts")

            out_bytes, fill_stats = fill_alloc(
                co_bytes, index, MOD["effort_key"], MOD["sheet_keywords"], co_cb,
                vba=co_file.name.lower().endswith(".xlsm"))
            if st.session_state.get("inc_vol_" + MOD["id"], True):
                prog.progress(97, text="Writing VOLARE STATUS file…")
                st.session_state["vol_file_" + MOD["id"]] = build_volare_file(vol_sorted)
            elapsed = _time.time() - t0
            prog.progress(100, text=f"Done ✔  ({elapsed:.1f}s)")

            run_ts = datetime.now()
            report = build_status_report(
                stats, fill_stats, run_ts,
                drr_file.name, co_file.name, notes_src, elapsed, MOD,
            )

        st.session_state["alloc_out_" + MOD["id"]] = out_bytes
        st.session_state["alloc_ext_" + MOD["id"]] = (
            ".xlsm" if co_file.name.lower().endswith(".xlsm") else ".xlsx")
        st.session_state["alloc_stats_" + MOD["id"]] = (stats, fill_stats)
        st.session_state["alloc_report_" + MOD["id"]] = report
        st.session_state["alloc_run_ts_" + MOD["id"]] = run_ts

    except Exception as e:
        prog.empty()
        st.error(f"Processing failed: {e}")
        st.stop()

# ---------------- step 3 : status report ----------------
if "alloc_out_" + MOD["id"] in st.session_state:
    stats, fill_stats = st.session_state["alloc_stats_" + MOD["id"]]
    report = st.session_state["alloc_report_" + MOD["id"]]
    run_ts = st.session_state["alloc_run_ts_" + MOD["id"]]

    st.markdown(
        '<div class="stepcard"><span class="stepno">3</span>'
        '<span class="steptitle">Status report</span>'
        f'<div class="stepdesc">Run completed {run_ts:%m/%d/%Y %I:%M %p}. '
        'Review the numbers below before submitting.</div></div>',
        unsafe_allow_html=True,
    )

    tab_sum, tab_break, tab_report = st.tabs(
        ["📊 Summary", "📈 Breakdown", "🧾 Full report"]
    )

    with tab_sum:
        c1, c2, c3, c4 = st.columns(4)
        if IS_FIELD:
            c1.metric("RESULT rows scanned", f'{stats["total_rows"]:,}')
            c2.metric("Rows kept (filter)", f'{stats["kept_rows"]:,}')
        else:
            c1.metric("DRR rows cleaned", f'{stats["total_rows"]:,}')
            c2.metric("Dates fixed (swap)", f'{stats["date_swapped"]:,}')
        c3.metric("Accounts filled", f'{fill_stats["filled"]:,}')
        unmapped_total = sum(stats["unmapped_status"].values())
        c4.metric("Unmapped statuses", f"{unmapped_total:,}")

        filled = fill_stats["filled"] or 1
        if IS_PAY:
            c5, c6, c7, c8 = st.columns(4)
            c5.metric("💵 With PAYMENT", f'{fill_stats["with_pay"]:,}')
            c6.metric("🤝 With PTP", f'{fill_stats["with_ptp"]:,}')
            c7.metric("Total collected", f'{fill_stats["total_collected"]:,.2f}')
            c8.metric("Total PTP amount", f'{fill_stats["ptp_amount_sum"]:,.2f}')
            flags = st.session_state.get("pay_flags_" + MOD["id"], ({}, {}))
            if flags[0]:
                st.error("⚠ CONFIRMED-type statuses with payment amounts are "
                         "MISSING from NOTES (col U) and were excluded — see the "
                         "Full report tab: "
                         + "; ".join(f"{k} ({v:,})" for k, v in flags[0].items()))
            if flags[1]:
                st.error("⚠ PTP-type statuses with PTP amounts are MISSING from "
                         "NOTES (col A) and were excluded — see the Full report "
                         "tab: "
                         + "; ".join(f"{k} ({v:,})" for k, v in flags[1].items()))
        else:
            c5, c6, c7 = st.columns(3)
            c5.metric("🟢 POS", f'{fill_stats["pos"]:,}',
                      f'{fill_stats["pos"] / filled * 100:.1f}%')
            c6.metric("🔴 NEG", f'{fill_stats["neg"]:,}',
                      f'-{fill_stats["neg"] / filled * 100:.1f}%')
            c7.metric("⚪ " + ("NO VISITATION" if IS_FIELD else "NO EFFORT"),
                      f'{fill_stats["no_effort"]:,}',
                      f'-{fill_stats["no_effort"] / filled * 100:.1f}%')

        if fill_stats["min_date"] and fill_stats["max_date"]:
            st.caption(
                f"📅 {MOD['date_label']}: **{fill_stats['min_date']:%m/%d/%Y} – "
                f"{fill_stats['max_date']:%m/%d/%Y}**"
            )

        if stats.get("sheets_used"):
            st.caption("📑 Sheets read: " + " · ".join(
                f"**{sh['name']}** ({sh['rows']:,} rows, {sh['swap']})"
                for sh in stats["sheets_used"]))
        if stats["missing_headers"]:
            st.warning("Expected headers not found in the raw file (left blank): "
                       + ", ".join(stats["missing_headers"]))
        if stats.get("extra_headers"):
            st.info("Extra DRR columns ignored: "
                    + ", ".join(map(str, stats["extra_headers"])))
        if stats["unmapped_status"]:
            top = stats["unmapped_status"].most_common(10)
            st.warning(
                f'{unmapped_total:,} DRR rows have statuses missing from the NOTES '
                f"key (they can never match, same as Excel #N/A). Top offenders: "
                + "; ".join(f"{s} ({c:,})" for s, c in top)
            )

    with tab_break:
        left, right = st.columns([1, 1])
        with left:
            st.markdown("**" + ("Payment type mix (T column)" if IS_PAY
                        else "Outcome mix (" + ("U column" if IS_FIELD else "T column") + ")") + "**")
            if IS_PAY:
                mix = dict(fill_stats["status_breakdown"]) or {"(none)": 0}
            else:
                mix = {
                    "POS": fill_stats["pos"],
                    "NEG": fill_stats["neg"],
                    ("NO VISITATION" if IS_FIELD else "NO EFFORT"): fill_stats["no_effort"],
                }
            st.bar_chart(mix, horizontal=True, color="#3b82f6")
        with right:
            st.markdown("**Top 10 " + ("PTP statuses (X column)" if IS_PAY
                        else "results (" + ("W column" if IS_FIELD else "U column") + ")") + "**")
            top10 = fill_stats["result_breakdown"].most_common(10)
            if top10:
                st.dataframe(
                    {"RESULT": [k for k, _ in top10],
                     "ACCOUNTS": [v for _, v in top10]},
                    use_container_width=True, hide_index=True,
                )
            else:
                st.caption("No matched results to display.")

    with tab_report:
        st.markdown('<div class="reportbox">', unsafe_allow_html=True)
        st.code(report, language=None)
        st.markdown('</div>', unsafe_allow_html=True)

    # ---------------- fixer : add missing statuses to NOTES ----------------
    if not IS_FIELD:
        cand = list(stats["unmapped_status"].keys())
        if IS_PAY:
            fl = st.session_state.get("pay_flags_" + MOD["id"], ({}, {}))
            cand += [s for s in list(fl[0]) + list(fl[1]) if s not in cand]
        if cand:
            with st.expander(f"🧩 Add {len(cand)} missing status(es) to NOTES",
                             expanded=False):
                st.caption(
                    "Suggested mappings are copied from the closest existing "
                    "NOTES entry (shown in COPIED FROM) — review, edit if "
                    "needed, untick any you don't want, then save."
                )
                naf_s, npq_s, nux_s = st.session_state.get(
                    "notes_tables", ({}, {}, {}))
                fk = "fixer_rows_" + MOD["id"]
                if (fk not in st.session_state
                        or {r["STATUS"] for r in st.session_state[fk]}
                        != set(map(str, cand))):
                    st.session_state[fk] = suggest_notes_mappings(
                        cand, naf_s, npq_s, nux_s)
                edited = st.data_editor(
                    st.session_state[fk],
                    use_container_width=True, hide_index=True,
                    key="fixer_editor_" + MOD["id"],
                    disabled=["STATUS", "COPIED FROM"],
                )
                if st.button("💾 Save mappings — used on the next RUN, and "
                             "prepare the updated NOTES file",
                             key="fixer_save_" + MOD["id"],
                             use_container_width=True):
                    newp = _rows_to_patch(edited)
                    merged = st.session_state.get("notes_patch", {})
                    merged.update(newp)
                    st.session_state["notes_patch"] = merged
                    nb = st.session_state.get("notes_srcfile")
                    if nb:
                        st.session_state["notes_updated_bytes"] = \
                            build_updated_notes(nb[0], nb[1], merged)
                    st.success(
                        f"Saved {len(newp)} mapping(s). Press RUN again to "
                        "apply them to the results — or download the updated "
                        "NOTES file below and use it from now on."
                    )
                if st.session_state.get("notes_updated_bytes"):
                    st.download_button(
                        "⬇  NOTES - UPDATED.xlsx (with the new statuses)",
                        data=st.session_state["notes_updated_bytes"],
                        file_name="NOTES - UPDATED.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="fixer_dl_" + MOD["id"],
                    )

    # ---------------- step 4 : download ----------------
    st.markdown(
        '<div class="stepcard"><span class="stepno">4</span>'
        '<span class="steptitle">Download</span>'
        f"<div class='stepdesc'>{MOD['alloc_label']} with {'S..AA' if IS_PAY else ('S/U/V/W' if IS_FIELD else 'S/T/U')} as static values, plus "
        "the status report for documentation.</div></div>",
        unsafe_allow_html=True,
    )
    d1, d2 = st.columns([2, 1])
    with d1:
        st.download_button(
            f"⬇  DOWNLOAD {MOD['name']} (FINAL)",
            data=st.session_state["alloc_out_" + MOD["id"]],
            file_name=MOD["out_file"].replace(
                ".xlsx", st.session_state.get("alloc_ext_" + MOD["id"], ".xlsx")),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "🧾  STATUS REPORT (.txt)",
            data=report,
            file_name=f"{MOD['report_file']} {run_ts:%m-%d-%Y}.txt",
            mime="text/plain",
            use_container_width=True,
        )
    if st.session_state.get("vol_file_" + MOD["id"]):
        st.download_button(
            "📊  VOLARE STATUS (FINAL) — final working order",
            data=st.session_state["vol_file_" + MOD["id"]],
            file_name="VOLARE STATUS - FINAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )