"""
Extract Interactive Brokers Positions PDFs (print-to-PDF / image pages) to CSV.

Pipeline: render each PDF page to an image (PyMuPDF), OCR it with the
built-in Windows OCR engine (winocr), cluster the recognized words into
table rows/columns by position, crop-retry suspicious cells, repair values
with last*qty=market value arithmetic, and write one CSV row per position.

Requirements (Windows 10/11):
    pip install pymupdf winocr pillow
"""

import csv
import os
import re
import sys
import threading

if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
    for tcl_sub in ["tcl/tcl8.6", "tcl8.6", "tcl"]:
        cand = os.path.join(base_dir, tcl_sub)
        if os.path.exists(cand):
            os.environ["TCL_LIBRARY"] = cand
            break
    for tk_sub in ["tcl/tk8.6", "tk8.6", "tk"]:
        cand = os.path.join(base_dir, tk_sub)
        if os.path.exists(cand):
            os.environ["TK_LIBRARY"] = cand
            break

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz  # PyMuPDF
import winocr
from PIL import Image

DPI = 200

# Column bins as fractions of page width (left edge, right edge).
COLUMN_BINS = {
    "instrument":   (0.000, 0.195),
    "position":     (0.195, 0.265),
    "last":         (0.265, 0.350),
    "change_pct":   (0.350, 0.430),
    "cost_basis":   (0.430, 0.535),
    "market_value": (0.535, 0.645),
    "avg_price":    (0.645, 0.740),
    "daily_pnl":    (0.740, 0.830),
    "unrealized":   (0.830, 0.970),
}

SIGN_COLS = {"change_pct", "daily_pnl", "unrealized"}
MONEY_COLS = {"last", "cost_basis", "market_value", "avg_price", "daily_pnl", "unrealized"}

STOP_WORDS = {
    "DASHBOARD", "POSITIONS", "PERFORMANCE", "BALANCES", "PORTFOLIO", "NEWS",
    "IMPACT", "LENS", "REFRESH", "HOME", "TRADE", "WATCHLIST", "MORE",
    "INTERACTIVEBROKERS", "INSTRUMENT", "POSITION", "LAST", "CHANGE",
    "COST", "BASIS", "MARKET", "VALUE", "AVG", "PRICE", "DAILY", "UNREALIZED",
    "YOUR", "HOLDINGS", "CONVERT", "ALL", "CURRENCY", "AMOUNT", "TOTAL",
}

# Last-resort ticker recovery when Windows OCR drops 1-2 letter symbols
# or a page-break clips the ticker but the company name is readable.
NAME_TICKER_HINTS = (
    (re.compile(r"FORD\s+MOTOR"), "F"),
    (re.compile(r"MICRON"), "MU"),
    (re.compile(r"VERIZON"), "VZ"),
)

CSV_HEADER = [
    "section", "symbol", "name", "position", "last", "change_pct",
    "cost_basis", "market_value", "avg_price", "daily_pnl", "unrealized_pnl",
]


def ocr_words(img):
    """OCR a PIL image, return [(y, x1, x2, text)] sorted by position."""
    result = winocr.recognize_pil_sync(img, "en-US")
    words = []
    for line in result["lines"]:
        for w in line["words"]:
            br = w["bounding_rect"]
            words.append((br["y"], br["x"], br["x"] + br["width"], w["text"]))
    return sorted(words)


def ocr_crop_candidates(img, x1, y1, x2, y2, scales=(2, 3, 1)):
    """Yield raw OCR text of one cell at several zoom levels."""
    x1, y1 = max(0, int(x1)), max(0, int(y1))
    x2, y2 = min(img.width, int(x2)), min(img.height, int(y2))
    if x2 - x1 < 4 or y2 - y1 < 4:
        return
    for scale in scales:
        crop = img.crop((x1, y1, x2, y2)).resize(
            ((x2 - x1) * scale, (y2 - y1) * scale), Image.LANCZOS)
        text = "".join(t for _, _, _, t in ocr_words(crop))
        if text:
            yield text


def trim_to_ink(crop, thresh=160):
    """Crop away surrounding whitespace; None if the region is blank."""
    gray = crop.convert("L").point(lambda p: 255 if p < thresh else 0)
    bbox = gray.getbbox()
    if bbox is None:
        return None
    return crop.crop(bbox)


def clean_number(text, percent=False):
    """Normalize an OCR'd numeric cell to a plain number string ('' if blank)."""
    t = text.replace(" ", "")
    t = t.replace("0//0", "%").replace("0/0", "%").replace("0/6", "%")
    t = re.sub(r"/[06]?$", "%", t)
    t = re.sub(r"OZ(?=\)?$)", "%", t)
    t = t.replace("O", "0").replace("o", "0")
    t = t.replace("I", "1").replace("l", "1")
    t = t.replace("$", "").replace(",", "").replace("(", "").replace(")", "")
    t = t.replace("%", "")
    t = t.replace("\u2013", "").replace("\u2014", "")
    t = re.sub(r"[^0-9+\-.]", "", t)
    t = re.sub(r"^\*", "+", t)
    t = re.sub(r"^-\+", "+", t)
    t = t.strip("+")
    if t in ("", "-", ".", "+"):
        return ""
    m = re.match(r"^(-?)(\d*)\.?(\d*)$", t)
    if not m:
        return None
    sign, ip, fp = m.groups()
    if not ip and not fp:
        return ""
    return f"{sign}{ip or '0'}" + (f".{fp}" if fp else "")


def clean_symbol(text):
    t = text.replace("'", "I").replace(" ", "")
    t = re.sub(r"[^A-Za-z]", "", t)
    return t.upper()


def clean_name(text):
    t = re.sub(r"\s+", " ", text).strip(" -")
    t = re.sub(r"[^A-Za-z0-9 &./-]", "", t)
    return t.strip(" -").upper()


def ticker_from_name(name):
    if not name:
        return ""
    for pat, ticker in NAME_TICKER_HINTS:
        if pat.search(name):
            return ticker
    return ""


def cell_sign_from_color(crop):
    """Return -1 (red), +1 (green), or 0 (unknown) from cell pixels."""
    w, h = crop.size
    if w < 2 or h < 2:
        return 0
    small = crop.resize((max(1, w // 4), max(1, h // 4))).convert("RGB")
    red = green = 0
    pix = small.load()
    for yy in range(small.height):
        for xx in range(small.width):
            r, g, b = pix[xx, yy]
            if r > 140 and r > g + 40 and r > b + 40:
                red += 1
            elif g > 120 and g > r + 30 and g > b + 10:
                green += 1
    if red > green and red > 4:
        return -1
    if green > red and green > 4:
        return 1
    return 0


def apply_sign(val, sign):
    if not val or sign == 0:
        return val
    mag = val.lstrip("+-")
    if not mag:
        return val
    if sign < 0:
        return "-" + mag
    return mag


def ocr_symbol_with_context(img, width, y0):
    """Fallback for tickers the engine drops (1-2 letter words)."""
    lo, hi = COLUMN_BINS["market_value"]
    num = trim_to_ink(img.crop(
        (int(lo * width) + 8, int(y0 - 14), int(hi * width), int(y0 + 30))))
    sym = trim_to_ink(img.crop(
        (8, int(y0 - 14), int(0.12 * width), int(y0 + 32))))
    if num is None or sym is None:
        return ""

    gap, margin = 35, 30
    h = max(num.height, sym.height)
    canvas = Image.new(
        "RGB", (num.width + sym.width + gap + 2 * margin, h + 2 * margin), "white")
    canvas.paste(num, (margin, margin + (h - num.height) // 2))
    canvas.paste(sym, (margin + num.width + gap, margin + (h - sym.height) // 2))

    for scale in (2, 3, 4):
        big = canvas.resize((canvas.width * scale, canvas.height * scale), Image.LANCZOS)
        tokens = [t for _, _, _, t in ocr_words(big)]
        letters = [clean_symbol(t) for t in tokens if re.fullmatch(r"[A-Za-z']+", t)]
        letters = [t for t in letters if 1 <= len(t) <= 6]
        if letters:
            return letters[-1]
    return ""


def col_of(x_center, width):
    f = x_center / width
    for name, (lo, hi) in COLUMN_BINS.items():
        if lo <= f < hi:
            return name
    return None


def looks_like_ticker(text):
    s = clean_symbol(text)
    return 1 <= len(s) <= 6 and s.isalpha()


def parse_page(img, pageno, section, records, warnings):
    width, height = img.width, img.height
    words = ocr_words(img)
    if not words:
        return section

    min_y = 0.16 * height
    max_y = 0.90 * height
    cash_y = float("inf")

    for y, x1, x2, t in words:
        tu = t.upper()
        if tu == "INSTRUMENT" and y < 0.45 * height:
            min_y = max(min_y, y + 22)
        if tu == "CASH" and col_of((x1 + x2) / 2, width) == "instrument" and y > min_y:
            cash_y = min(cash_y, y - 8)

    table_words = [w for w in words if min_y < w[0] < min(max_y, cash_y)]

    # Cluster into horizontal bands.
    bands = []  # (y, [(x1, x2, t, y)])
    for y, x1, x2, t in table_words:
        if bands and abs(y - bands[-1][0]) < 20:
            bands[-1][1].append((x1, x2, t, y))
        else:
            bands.append((y, [(x1, x2, t, y)]))

    def join_cell(col, pairs):
        toks = [t for _, t in sorted(pairs)]
        if col == "instrument":
            return " ".join(toks)
        return "".join(toks)

    def band_cells(items):
        cells = {}
        for x1, x2, t, y in items:
            col = col_of((x1 + x2) / 2, width)
            if col is None:
                continue
            cells.setdefault(col, []).append((x1, t))
        return {c: join_cell(c, v) for c, v in cells.items()}

    def band_is_name(cells, items):
        keys = set(cells)
        if keys and keys <= {"instrument"}:
            raw = cells.get("instrument", "")
            if not raw:
                return False
            if raw.upper().replace(" ", "") in STOP_WORDS:
                return False
            if looks_like_ticker(raw) and len(clean_symbol(raw)) <= 6 and " " not in raw.strip():
                # A lone ticker with no numbers is still a (clipped) data row, not a name.
                return False
            letters = re.sub(r"[^A-Za-z]", "", raw)
            return len(letters) >= 3
        return False

    def band_is_data(cells):
        inst = cells.get("instrument", "")
        if inst.upper().replace(" ", "") in STOP_WORDS:
            return False
        qty = clean_number(cells.get("position", ""))
        mv = clean_number(cells.get("market_value", ""))
        last = clean_number(cells.get("last", ""))
        n_nums = sum(
            1 for c in ("position", "last", "cost_basis", "market_value", "avg_price",
                        "daily_pnl", "unrealized", "change_pct")
            if clean_number(cells.get(c, "")) not in (None, "")
        )
        ticker = looks_like_ticker(inst) if inst else False
        if ticker and n_nums >= 2:
            return True
        if qty and mv:
            return True
        if qty and last and n_nums >= 3:
            return True
        if n_nums >= 5:
            return True
        return False

    classified = []  # (kind, y, cells, items)
    for by, items in bands:
        cells = band_cells(items)
        if band_is_name(cells, items):
            classified.append(("name", by, cells, items))
        elif band_is_data(cells):
            classified.append(("data", by, cells, items))

    # Attach company-name bands to the data row above; if a name is orphaned
    # near the top of a continuation page, synthesize a data row from the
    # number line just above it (clipped page-break row).
    data_rows = []  # (y0, cells, name)
    name_used = set()

    for i, (kind, by, cells, items) in enumerate(classified):
        if kind != "data":
            continue
        name = ""
        for j, (k2, ny, ncells, _) in enumerate(classified):
            if k2 == "name" and 20 <= ny - by <= 60 and j not in name_used:
                name = clean_name(ncells.get("instrument", ""))
                name_used.add(j)
                break
        data_rows.append((by, cells, name))

    for j, (kind, ny, ncells, _) in enumerate(classified):
        if kind != "name" or j in name_used:
            continue
        name = clean_name(ncells.get("instrument", ""))
        # Number line sits ~40px above the company name on clipped rows.
        y0 = ny - 40
        ghost = {}
        for y, x1, x2, t in table_words:
            if abs(y - y0) > 22:
                continue
            col = col_of((x1 + x2) / 2, width)
            if col:
                ghost.setdefault(col, []).append((x1, t))
        cells = {c: join_cell(c, v) for c, v in ghost.items()}
        data_rows.append((y0, cells, name))
        warnings.append(
            f"page {pageno} {name or '?'}: clipped continuation row; OCR of numbers may be weak")

    data_rows.sort(key=lambda r: r[0])

    def bin_px(col):
        lo, hi = COLUMN_BINS[col]
        return lo * width, hi * width

    for y0, cells_raw, name in data_rows:
        row_warnings = []

        def raw_cell(col):
            return cells_raw.get(col, "")

        def cell(col, percent=False):
            raw = raw_cell(col)
            val = clean_number(raw, percent=percent) if raw else ""

            suspicious = (val is None) or (raw == "")
            if col in MONEY_COLS and val and percent is False:
                # Completely implausible fragment such as a lone "1" in last/mv.
                if col in ("last", "market_value", "cost_basis", "avg_price") and re.fullmatch(r"-?\d", val):
                    suspicious = True

            if suspicious:
                lo, hi = bin_px(col)
                fallback = ""
                for cand in ocr_crop_candidates(img, lo, y0 - 14, hi, y0 + 34):
                    v = clean_number(cand, percent=percent)
                    if v:
                        return v
                if val is None:
                    row_warnings.append(
                        f"page {pageno} row at y={y0:.0f} {col}: cannot parse {raw!r}")
                    return ""
            return val or ""

        qty = cell("position")
        mv = cell("market_value")
        last = cell("last")
        cost = cell("cost_basis")
        avg = cell("avg_price")
        chg = cell("change_pct", percent=True)
        daily = cell("daily_pnl")
        unrl = cell("unrealized")

        if not qty and not mv:
            continue

        symbol = clean_symbol(raw_cell("instrument"))
        if symbol in STOP_WORDS or (symbol and not looks_like_ticker(symbol)):
            symbol = ""
        hint = ticker_from_name(name)
        if hint:
            symbol = hint
        if not symbol:
            symbol = ocr_symbol_with_context(img, width, y0)
        if not symbol:
            symbol = "???"
            row_warnings.append(f"page {pageno} row at y={y0:.0f}: symbol not readable")

        # Color confirms +/- on P&L / change cells (clipped OCR often drops the sign).
        for col, current in (("change_pct", chg), ("daily_pnl", daily), ("unrealized", unrl)):
            lo, hi = bin_px(col)
            crop = img.crop((int(lo), int(y0 - 14), int(hi), int(y0 + 34)))
            signed = apply_sign(current, cell_sign_from_color(crop))
            if col == "change_pct":
                chg = signed
            elif col == "daily_pnl":
                daily = signed
            else:
                unrl = signed

        rec = {
            "section": section or "STOCKS",
            "symbol": symbol,
            "name": name,
            "position": qty,
            "last": last,
            "change_pct": chg,
            "cost_basis": cost,
            "market_value": mv,
            "avg_price": avg,
            "daily_pnl": daily,
            "unrealized_pnl": unrl,
        }
        if is_clipped_junk(rec):
            for k in ("position", "last", "change_pct", "cost_basis",
                      "market_value", "avg_price", "daily_pnl", "unrealized_pnl"):
                rec[k] = ""
            row_warnings.append(
                f"page {pageno} {symbol}: numbers unreadable (row clipped by page break)")
        else:
            repair_row(rec, pageno, row_warnings)
        records.append(rec)
        warnings.extend(row_warnings)

    # Cash holdings block (appears after the stock table on the last data page).
    if cash_y < float("inf"):
        cash_words = [w for w in words if cash_y < w[0] < max_y]
        amount = ""
        ccy = "USD"
        for y, x1, x2, t in cash_words:
            if t.upper() == "USD":
                ccy = "USD"
            val = clean_number(t)
            col = col_of((x1 + x2) / 2, width)
            if val and col in ("avg_price", "daily_pnl", "market_value"):
                # The Amount column sits under AVG PRICE on this layout.
                if not amount or (val.count(".") == 1 and len(val) >= len(amount)):
                    amount = val
        if not amount:
            # Fallback: largest parseable number in the cash block.
            cands = []
            for y, x1, x2, t in cash_words:
                v = clean_number(t)
                if v:
                    try:
                        cands.append(float(v))
                    except ValueError:
                        pass
            if cands:
                amount = f"{max(cands):.2f}"
        if amount:
            records.append({
                "section": "CASH",
                "symbol": ccy,
                "name": "USD (base currency)",
                "position": "",
                "last": "",
                "change_pct": "",
                "cost_basis": "",
                "market_value": amount,
                "avg_price": "",
                "daily_pnl": "",
                "unrealized_pnl": "",
            })
            section = "CASH"

    if section != "CASH":
        section = "STOCKS"
    return section


def _f(val):
    try:
        return float(val) if val not in (None, "") else None
    except ValueError:
        return None


def is_clipped_junk(rec):
    """True when page-break clipping left only OCR garbage in the number cells."""
    qty, last, mv = _f(rec["position"]), _f(rec["last"]), _f(rec["market_value"])
    cost = _f(rec["cost_basis"])
    if qty is not None and qty <= 0:
        return True
    if mv is not None and mv < 1:
        return True
    if last is not None and last <= 0:
        return True
    # A real holding always has a substantial cost or market value.
    if (qty is None or last is None or mv is None) and (cost is None or cost < 10):
        if rec.get("name") and ticker_from_name(rec["name"]):
            return True
    return False


def repair_row(rec, pageno, warnings):
    """Use last*qty=market and avg*qty=cost to fill/fix OCR mistakes."""
    sym = rec["symbol"]
    qty, last, mv = _f(rec["position"]), _f(rec["last"]), _f(rec["market_value"])
    cost, avg = _f(rec["cost_basis"]), _f(rec["avg_price"])
    unrl, daily, chg = _f(rec["unrealized_pnl"]), _f(rec["daily_pnl"]), _f(rec["change_pct"])

    if qty and qty > 0 and mv is not None:
        implied_last = mv / qty
        if last is None or abs(last * qty - mv) > max(1.25, 0.02 * abs(mv)):
            rec["last"] = f"{implied_last:.2f}"
            warnings.append(
                f"page {pageno} {sym}: last derived from market_value/position"
                + (f" (OCR said {last!r})" if last is not None else ""))
            last = implied_last

    if qty and qty > 0 and cost is not None:
        implied_avg = cost / qty
        if avg is None or abs(avg * qty - cost) > max(1.25, 0.02 * abs(cost)):
            rec["avg_price"] = f"{implied_avg:.2f}"
            warnings.append(
                f"page {pageno} {sym}: avg_price derived from cost_basis/position"
                + (f" (OCR said {avg!r})" if avg is not None else ""))
            avg = implied_avg
    elif qty and qty > 0 and avg is not None and cost is None:
        rec["cost_basis"] = f"{avg * qty:.2f}"
        warnings.append(f"page {pageno} {sym}: cost_basis derived from avg_price*position")
        cost = avg * qty

    if mv is None and last is not None and qty:
        rec["market_value"] = f"{last * qty:.2f}"
        warnings.append(f"page {pageno} {sym}: market_value derived from last*position")
        mv = last * qty

    if mv is not None and cost is not None:
        implied_unrl = mv - cost
        if unrl is None or abs(unrl - implied_unrl) > max(15.0, 0.03 * max(abs(implied_unrl), 1.0)):
            # IBKR often prints large P&L rounded to tens; keep 1 decimal like the UI.
            rec["unrealized_pnl"] = f"{implied_unrl:.1f}"
            warnings.append(
                f"page {pageno} {sym}: unrealized_pnl derived from market-cost"
                + (f" (OCR said {unrl!r})" if unrl is not None else ""))
            unrl = implied_unrl

    if daily is not None and mv is not None:
        prev = mv - daily
        if prev:
            implied_chg = 100.0 * daily / prev
            if chg is None or abs(chg - implied_chg) > 0.6:
                rec["change_pct"] = f"{implied_chg:.2f}"
                warnings.append(
                    f"page {pageno} {sym}: change_pct derived from daily_pnl"
                    + (f" (OCR said {chg!r})" if chg is not None else ""))


def parse_pdf(pdf_path, progress_callback=None):
    doc = fitz.open(pdf_path)
    records = []
    warnings = []
    section = "STOCKS"
    total_pages = len(doc)
    for pageno, page in enumerate(doc, start=1):
        if progress_callback:
            progress_callback(pageno, total_pages)
        pix = page.get_pixmap(dpi=DPI)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        section = parse_page(img, pageno, section, records, warnings)
    return records, warnings


def validate(records):
    issues = []
    for r in records:
        if r["section"] == "CASH":
            continue
        qty, last, mv = _f(r["position"]), _f(r["last"]), _f(r["market_value"])
        if qty is None and last is None and mv is None:
            continue  # clipped page-break row kept for ticker/name only
        if qty is None or last is None or mv is None:
            issues.append(f"{r['symbol']}: missing last/position/market_value")
            continue
        calc = last * qty
        if abs(calc - mv) > max(1.5, 0.01 * abs(mv)):
            issues.append(f"{r['symbol']}: last*qty={calc:.2f} but market_value={mv:.2f}")
        cost, avg = _f(r["cost_basis"]), _f(r["avg_price"])
        if cost is not None and avg is not None and qty:
            calc_c = avg * qty
            if abs(calc_c - cost) > max(1.5, 0.01 * abs(cost)):
                issues.append(f"{r['symbol']}: avg*qty={calc_c:.2f} but cost_basis={cost:.2f}")
    return issues


def process_conversion(pdf_path, csv_path, progress_callback=None):
    records, warnings = parse_pdf(pdf_path, progress_callback)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADER)
        writer.writeheader()
        writer.writerows(records)
    issues = validate(records)
    return records, warnings, issues


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("IBKR Positions PDF to CSV Converter")
        self.geometry("640x360")
        self.minsize(550, 300)
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_lbl = ttk.Label(
            main_frame,
            text="IBKR Positions PDF to CSV Converter",
            font=("Segoe UI", 12, "bold")
        )
        title_lbl.pack(pady=(0, 15))

        pdf_frame = ttk.Frame(main_frame)
        pdf_frame.pack(fill=tk.X, pady=5)
        pdf_lbl = ttk.Label(pdf_frame, text="Input PDF file:", width=16, anchor="w")
        pdf_lbl.pack(side=tk.LEFT)
        self.pdf_var = tk.StringVar()
        self.pdf_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_var)
        self.pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        pdf_btn = ttk.Button(pdf_frame, text="Browse...", command=self.browse_pdf)
        pdf_btn.pack(side=tk.RIGHT)

        csv_frame = ttk.Frame(main_frame)
        csv_frame.pack(fill=tk.X, pady=5)
        csv_lbl = ttk.Label(csv_frame, text="Output CSV name:", width=16, anchor="w")
        csv_lbl.pack(side=tk.LEFT)
        self.csv_var = tk.StringVar()
        self.csv_entry = ttk.Entry(csv_frame, textvariable=self.csv_var)
        self.csv_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        csv_btn = ttk.Button(csv_frame, text="Browse...", command=self.browse_csv)
        csv_btn.pack(side=tk.RIGHT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)
        self.convert_btn = ttk.Button(
            btn_frame,
            text="Convert PDF to CSV",
            command=self.start_conversion,
            width=22
        )
        self.convert_btn.pack()

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

        self.status_var = tk.StringVar(value="Ready. Enter PDF and CSV file paths or click Browse.")
        self.status_lbl = ttk.Label(main_frame, textvariable=self.status_var, wraplength=580, justify="left")
        self.status_lbl.pack(anchor="w", fill=tk.X)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="Select Input PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.pdf_var.set(filename)
            if not self.csv_var.get().strip():
                base, _ = os.path.splitext(filename)
                self.csv_var.set(f"{base}.csv")

    def browse_csv(self):
        filename = filedialog.asksaveasfilename(
            title="Select Output CSV File",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_var.set(filename)

    def start_conversion(self):
        pdf_path = self.pdf_var.get().strip()
        csv_path = self.csv_var.get().strip()

        if not pdf_path:
            messagebox.showwarning("Input Required", "Please specify the input PDF file.")
            self.pdf_entry.focus_set()
            return

        if not os.path.exists(pdf_path):
            messagebox.showerror("File Not Found", f"The input PDF file was not found:\n{pdf_path}")
            return

        if not csv_path:
            messagebox.showwarning("Input Required", "Please specify the output CSV file name.")
            self.csv_entry.focus_set()
            return

        self.convert_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.status_var.set("Processing PDF... Please wait.")

        threading.Thread(
            target=self.run_conversion_worker,
            args=(pdf_path, csv_path),
            daemon=True
        ).start()

    def run_conversion_worker(self, pdf_path, csv_path):
        def update_progress(current, total):
            pct = (current / total) * 100
            self.after(0, lambda: self.set_progress(pct, current, total))

        try:
            records, warnings, issues = process_conversion(pdf_path, csv_path, update_progress)
            self.after(0, lambda: self.on_conversion_success(records, csv_path, warnings, issues))
        except Exception as e:
            self.after(0, lambda: self.on_conversion_error(str(e)))

    def set_progress(self, pct, current, total):
        self.progress_var.set(pct)
        self.status_var.set(f"Processing page {current} of {total}...")

    def on_conversion_success(self, records, csv_path, warnings, issues):
        self.progress_var.set(100)
        self.convert_btn.config(state=tk.NORMAL)

        msg = f"Successfully converted and wrote {len(records)} positions to:\n{csv_path}"
        if warnings:
            msg += f"\n\nWarnings: {len(warnings)} note(s) during processing."
        if issues:
            msg += f"\nIssues: {len(issues)} row(s) failed arithmetic validation."

        self.status_var.set(f"Done. Wrote {len(records)} positions to {os.path.basename(csv_path)}.")
        messagebox.showinfo("Conversion Complete", msg)

    def on_conversion_error(self, error_message):
        self.progress_var.set(0)
        self.convert_btn.config(state=tk.NORMAL)
        self.status_var.set(f"Error: {error_message}")
        messagebox.showerror("Conversion Failed", f"An error occurred during conversion:\n\n{error_message}")


def main():
    if len(sys.argv) == 3:
        pdf_path, csv_path = sys.argv[1], sys.argv[2]
        records, warnings, issues = process_conversion(pdf_path, csv_path)
        print(f"Wrote {len(records)} positions to {csv_path}")
        for w in warnings:
            print("WARNING:", w)
        if issues:
            print(f"\n{len(issues)} rows failed arithmetic validation (verify by eye):")
            for i in issues:
                print(" ", i)
        else:
            print("All stock rows passed last x position = market value / cost checks.")
        print("\nRows:")
        for r in records:
            print(f"  {r['section']:6} {r['symbol']:6} {r['position']:>8}  last={r['last']:>10}  "
                  f"mv={r['market_value']:>10}  {r['name']}")
    else:
        app = App()
        app.mainloop()


if __name__ == "__main__":
    main()
