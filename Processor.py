"""
Personal Finance Dashboard
===========================
Parses Zopa and Nationwide PDF bank statements (and CSV files) into a
categorised transaction table with charts.

Dependencies: pandas, pdfplumber, matplotlib, tkinter (stdlib)
Install:  pip install pandas pdfplumber matplotlib --break-system-packages
"""

import csv
import json
import os
import re
from datetime import datetime

import matplotlib.pyplot as plt
import pandas as pd
import pdfplumber
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog, messagebox, simpledialog, ttk

# ── Constants ────────────────────────────────────────────────────────────────

AMOUNT_RE   = re.compile(r"^-?[\d,]+\.\d{2}$")
MONTH_NAMES = "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
DATE_RE     = re.compile(rf"^\d{{1,2}}\s+({MONTH_NAMES})$")
DATE_YEAR_RE = re.compile(rf"^\d{{1,2}}\s+({MONTH_NAMES})\s+\d{{4}}$")

OVERRIDES_FILE = "category_overrides.json"

DEFAULT_CATEGORIES = {
    "Income":          ["SALARY","WAGES","PAYROLL","BANK CREDIT","KEVA","INTEREST",
                        "CASHBACK","SWITCHING","REFUND","CREDIT"],
    "Transport":       ["TRANSPENNINE","NORTHERN RAIL","TPEXPRESS","TRAINLINE",
                        "NATIONAL RAIL","TFL","TFGM","METROLINK","BEENETWORK",
                        "STAGECOACH","ARRIVA","MEGABUS","UBER","BOLT","TAXI",
                        "EASYJET","RYANAIR","EUROSTAR","EAST MIDS RAILWAY",
                        "DVLA","PARKING","NCP","PETROL","FUEL","SHELL","BP ","ESSO"],
    "Groceries":       ["TESCO","SAINSBURY","ASDA","LIDL","ALDI","WAITROSE",
                        "MARKS&SPENCER","M&S FOOD","CO-OP","COOP","ICELAND",
                        "MORRISONS","WHOLEFOOD","OCADO","BUDGENS"],
    "Eating Out":      ["MCDONALD","BURGER KING","KFC","PIZZA","DOMINO","SUBWAY",
                        "NANDO","WAGAMAMA","PRET","GREGGS","COSTA","STARBUCKS",
                        "CAFFE NERO","CAFE NERO","ITSU","DELIVEROO","JUST EAT",
                        "UBER EATS","RESTAURANT","TAKEAWAY","TGTG","BOTANIST",
                        "FOOD WORKS","EZRA","BAKERY","SQ *"],
    "Shopping":        ["AMAZON","EBAY","ASOS","NEXT ","JOHN LEWIS","PRIMARK",
                        "H&M","ZARA","UNIQLO","TK MAXX","ARGOS","CURRYS",
                        "IKEA","HALFORDS","DECATHLON","ALIEXPRESS","B&Q","BQ "],
    "Bills":           ["OCTOPUS ENERGY","BRITISH GAS","EON","UNITED UTILITIES",
                        "THAMES WATER","EE LTD","VODAFONE","O2 ","THREE UK",
                        "LEBARA","COUNCIL TAX","MANCHESTER C C","MCC INTERNET",
                        "BROADBAND","SKY ","BT ","VIRGIN MEDIA","YW INTERNET",
                        "BEENETWORK.COM"],
    "Entertainment":   ["NETFLIX","SPOTIFY","STEAM","GOOGLE PLAY","YOUTUBE",
                        "DISNEY","APPLE MUSIC","AMAZON PRIME","NOW TV",
                        "CINEMA","ODEON","VUE ","CINEWORLD","TICKETMASTER",
                        "PLAYSTATION","XBOX","NINTENDO"],
    "Health":          ["BOOTS","SUPERDRUG","LLOYDS PHARMACY","PHARMACY",
                        "DENTIST","207 DENTAL","NHSBSA","SPECSAVERS","NHS ",
                        "PUREGYM","DAVID LLOYD","THE GYM","ANYTIME FITNESS",
                        "VIRGIN ACTIVE"],
    "Subscriptions":   ["SPOTIFY","ADOBE","MICROSOFT","GOOGLE ONE","ICLOUD",
                        "DROPBOX","1PASSWORD","NOTION","CANVA","SP FUSSY",
                        "SP MAKES SENSE","BEER 52","EVERYDAY APP"],
    "Charity":         ["CHARITY","GUIDE DOGS","CATS PROTECTION","BLUE CROSS",
                        "OXFAM","RED CROSS","CANCER RESEARCH"],
    "Insurance":       ["INSURANCE","AVIVA","AXA","DIRECT LINE","ADMIRAL",
                        "HASTINGS","LV="],
    "ATM / Cash":      ["ATM","CASH WITHDRAWAL","CASHPOINT","LINK ATM"],
    "Savings":         ["MONEYBOX","VANGUARD","FIDELITY","TRADING 212",
                        "FREETRADE","ISA ","SAVER","SAVINGS"],
    "Internal Transfer": ["FINLAY DAVIS","DAVIS F J","F J DAVIS",
                          "ZOPA BANK","SENT WITH ZOPA","SAVER 877","SAVER 5DF",
                          "071520 82165285","071120 52755329",
                          "ASHE HODGSON","HODGSON A","A HODGSON",
                          "ASHE SHEENA","SARAH SAATZER","HOLLIE LLOYD",
                          "THOMAS HODGSON","G W HODGSON","MS F J ARMSTRONG",
                          "MICHAEL SUDDABY","RIVER AULTD","MONEYBOX CASH"],
}

CATEGORY_ORDER = [
    "Income","Internal Transfer","Groceries","Eating Out","Transport",
    "Shopping","Bills","Entertainment","Health","Subscriptions",
    "Charity","Insurance","ATM / Cash","Savings","Other",
]

# ── PDF Parsers ───────────────────────────────────────────────────────────────

def _words_by_line(page) -> list[tuple[float, list[tuple[float, str]]]]:
    """Return page words grouped by y-coordinate → [(y, [(x, text), ...]), ...]"""
    groups: dict[float, list] = {}
    for w in page.extract_words(x_tolerance=3, y_tolerance=3):
        y = round(w["top"], 0)
        groups.setdefault(y, []).append((w["x0"], w["text"]))
    return [(y, sorted(tokens)) for y, tokens in sorted(groups.items())]


def _clean_amount(text: str) -> float:
    """'£1,234.56' → 1234.56  |  '-1,234.56' → -1234.56"""
    t = re.sub(r"[£,\s]", "", text)
    try:
        return float(t)
    except ValueError:
        return 0.0


def parse_zopa_pdf(path: str) -> tuple[list[dict], dict]:
    """
    Parse a Zopa bank statement PDF.

    Strategy: pdfplumber's table extractor finds each transaction as a
    separate 1-row table (Zopa uses thin rules between rows).  We collect
    every table row that starts with a valid date token and reconstruct the
    full transaction.  Multi-line description cells (e.g. Reference lines)
    are already joined by pdfplumber with '\\n'.

    Returns (transactions, meta) where meta = {account, opening, closing, ...}
    """
    transactions = []
    meta = {}

    with pdfplumber.open(path) as pdf:
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        # ── Meta from text ──────────────────────────────────────────────────
        for label, key in [
            (r"Sort code\s+([\d-]+)", "sort_code"),
            (r"Account number\s+(\d+)", "account_number"),
            (r"Opening (?:total )?balance\s+£([\d,]+\.\d{2})", "opening_balance"),
            (r"Closing (?:total )?balance\s+£([\d,]+\.\d{2})", "closing_balance"),
            (r"Total in\s+£([\d,]+\.\d{2})", "total_in"),
            (r"Total out\s+£([\d,]+\.\d{2})", "total_out"),
        ]:
            m = re.search(label, full_text, re.IGNORECASE)
            if m:
                meta[key] = m.group(1)

        # Grab account holder name
        m = re.search(r"Zopa Bank\n.+?\n(.+?)\n", full_text)
        if m:
            meta["holder"] = m.group(1).strip()

        # Determine statement year from "for DD Mon YYYY" line
        m = re.search(r"for \d+ \w+ (\d{4})", full_text, re.IGNORECASE)
        stmt_year = int(m.group(1)) if m else datetime.now().year

        # ── Transaction rows ────────────────────────────────────────────────
        # Columns in Zopa tables: [date, type, description, in, out, balance]
        date_pat = re.compile(rf"^\d{{2}}\s+({MONTH_NAMES})$")

        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    if not row or len(row) < 4:
                        continue
                    date_str = (row[0] or "").strip()
                    if not date_pat.match(date_str):
                        continue

                    txn_type = (row[1] or "").strip()
                    desc_raw = (row[2] or "").strip().replace("\n", " — ")
                    # Strip "Reference: …" suffix into a note
                    desc = re.sub(r"\s*—\s*Reference:.*", "", desc_raw).strip()
                    note = ""
                    m = re.search(r"Reference:\s*(.+)", desc_raw)
                    if m:
                        note = m.group(1).strip()

                    amt_in  = _clean_amount(row[3] or "")
                    amt_out = _clean_amount(row[4] or "") if len(row) > 4 else 0.0
                    balance = _clean_amount(row[5] or "") if len(row) > 5 else 0.0

                    try:
                        date = datetime.strptime(f"{date_str} {stmt_year}", "%d %b %Y").date()
                    except ValueError:
                        continue

                    amount = amt_in - amt_out  # positive = money in

                    transactions.append({
                        "date":        date,
                        "type":        txn_type,
                        "description": desc,
                        "note":        note,
                        "amount_in":   amt_in,
                        "amount_out":  amt_out,
                        "balance":     balance,
                        "bank":        "Zopa",
                        "amount":      amount,
                    })

    return transactions, meta


def parse_nationwide_pdf(path: str) -> tuple[list[dict], dict]:
    """
    Parse a Nationwide FlexAccount/FlexGraduate PDF statement.

    Strategy: pdfplumber's table extractor misses rows where the date column
    is blank (continuation rows on the same date) and rows where two adjacent
    transactions are too close together for the table detector.

    Instead we extract words grouped by y-coordinate, which gives us every
    printed line regardless of table structure.  We then reconstruct
    transactions by:
      1. Keeping lines that fall within the transaction table's x-range.
      2. Detecting date lines to start new transaction groups.
      3. Parsing amount columns by x-position (Out ≈ x<400, In ≈ x>400,
         Balance = last column).
    """
    transactions = []
    meta = {}

    # Column x boundaries discovered empirically from these statements:
    #   Date col:        x  <  90
    #   Description col: 90 <= x < 370
    #   £Out col:       370 <= x < 430
    #   £In col:        430 <= x < 490
    #   £Balance col:   490 <= x
    # These are approximate; we use amount-pattern matching as the real filter.
    TABLE_X_MIN = 55   # ignore page header/footer text to the left of this
    TABLE_X_MAX = 560  # and to the right

    # Y range: skip page header/footer.  We detect the header row "Date Description"
    # and the footer marker dynamically per page.

    with pdfplumber.open(path) as pdf:
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        # ── Meta ────────────────────────────────────────────────────────────
        for label, key in [
            (r"Sort code\s+([\d-]+)", "sort_code"),
            (r"Account no\s+(\d+)", "account_number"),
            (r"Start balance\s+£?([-\d,]+\.\d{2})", "opening_balance"),
            (r"End balance\s+£?([-\d,]+\.\d{2})", "closing_balance"),
        ]:
            m = re.search(label, full_text, re.IGNORECASE)
            if m:
                meta[key] = m.group(1)

        m = re.search(r"(Mx?[rs]?\s+[A-Z]\s*[A-Z]?\s+[A-Za-z]+)", full_text)
        if m:
            meta["holder"] = m.group(1).strip()

        # Determine year from "Statement date: DD Month YYYY"
        m = re.search(r"Statement\s*date[:\s]+\d+\s+\w+\s+(\d{4})", full_text, re.IGNORECASE)
        stmt_year = int(m.group(1)) if m else datetime.now().year
        # Nationwide statements can span two calendar years (e.g. Mar–Apr)
        # We'll infer year per transaction from the statement year and month ordering.

        # ── Per-page word extraction ─────────────────────────────────────────
        # We accumulate (date_str, description_parts, out, in_, balance) tuples.
        # A new date on a line starts a new "date group".
        # Lines without a date belong to the most recent date.

        # Raw line buffer across all pages
        raw_lines: list[tuple[float, float, str]] = []  # (page_idx, y, text)

        for page_idx, page in enumerate(pdf.pages):
            line_groups = _words_by_line(page)

            # Find y-range of the transaction table on this page.
            # The table header contains "Date" and "Description" on the same line.
            table_y_start = 0.0
            table_y_end   = 9999.0

            header_ys = []
            for y, tokens in line_groups:
                joined = " ".join(t for _, t in tokens)
                if "Date" in joined and "Description" in joined and "£Out" in joined:
                    header_ys.append(y)
                # Footer: Nationwide puts DC83/DC86 codes near the bottom
                if re.search(r"DC\d+\s*\(", joined):
                    table_y_end = min(table_y_end, y)

            if header_ys:
                table_y_start = header_ys[-1]  # last header on this page

            for y, tokens in line_groups:
                if y <= table_y_start or y >= table_y_end:
                    continue
                # Filter to table x-range
                filtered = [(x, t) for x, t in tokens if TABLE_X_MIN <= x <= TABLE_X_MAX]
                if not filtered:
                    continue
                line_text = " ".join(t for _, t in filtered)
                raw_lines.append((page_idx, y, line_text))

        # ── Parse raw lines into transactions ────────────────────────────────
        # Nationwide lines look like:
        #   "11 Mar Transfer from 071520 82165285 50.00"
        #   "Payment to SARAH SAATZER 50.00 -2,991.06"   ← continuation, no date
        #   "2026 Balance from statement 131 …"           ← skip
        #   "2026 -2,998.17"                              ← skip (page carry-over)

        # We group lines into "date groups" where the first line has a date.
        # Within a group every line gets the same date.
        # Each line may have 0, 1, 2 or 3 amount tokens at the end.

        SKIP_RE = re.compile(
            r"^2026\b|^Balance from statement|Statement\s*date|"
            r"Statement\s*no|Sort\s*code|Account\s*no|"
            r"Effective Date|Receiving an|International Payment|"
            r"NAIA|MIDLGB|IBAN|BIC|Swift|Intermediary",
            re.IGNORECASE
        )

        # Infer year for a given month number relative to statement year
        def infer_year(month_num: int) -> int:
            # Statement ends in stmt_year; if month > current statement month,
            # it must be the previous year.
            end_month = int(re.search(r"(\d{4})", meta.get("closing_balance", str(stmt_year))).group(1)) \
                if False else stmt_year  # simplified: just use stmt_year
            # For a March–April statement ending in April 2026, all months are 2026.
            return stmt_year

        current_date_str = None
        pending: list[dict] = []  # lines awaiting date assignment

        def flush_group(lines_in_group: list[str], date_str: str) -> list[dict]:
            """Convert a group of text lines sharing a date into transaction dicts."""
            result = []
            for line in lines_in_group:
                if SKIP_RE.match(line):
                    continue
                tokens = line.split()
                if not tokens:
                    continue

                # Strip leading date tokens if present
                if len(tokens) >= 2 and DATE_RE.match(" ".join(tokens[:2])):
                    tokens = tokens[2:]
                if not tokens:
                    continue

                # Extract trailing amount tokens
                amounts = []
                while tokens and AMOUNT_RE.match(tokens[-1].lstrip("-")):
                    amounts.insert(0, tokens.pop())
                
                if not tokens and not amounts:
                    continue

                desc = " ".join(tokens).strip()
                # Skip pure metadata lines that sneaked through
                if not desc and not amounts:
                    continue
                if desc in ("", "IBAN", "BIC", "Swift"):
                    continue

                # Map amounts to out/in/balance
                # Nationwide: columns are £Out | £In | £Balance
                # Amounts are always positive in the PDF; balance can be negative (shown as -N)
                amt_out = amt_in = balance = 0.0
                if len(amounts) == 3:
                    amt_out  = _clean_amount(amounts[0])
                    amt_in   = _clean_amount(amounts[1])
                    balance  = _clean_amount(amounts[2])
                elif len(amounts) == 2:
                    # Could be (out, balance) or (in, balance)
                    # We disambiguate by description keywords
                    desc_up = desc.upper()
                    is_credit = any(k in desc_up for k in (
                        "TRANSFER FROM","BANK CREDIT","SALARY","CREDIT","CASHBACK",
                        "INTEREST","REFUND","SWITCHING"
                    ))
                    if is_credit:
                        amt_in  = _clean_amount(amounts[0])
                        balance = _clean_amount(amounts[1])
                    else:
                        amt_out = _clean_amount(amounts[0])
                        balance = _clean_amount(amounts[1])
                elif len(amounts) == 1:
                    desc_up = desc.upper()
                    is_credit = any(k in desc_up for k in (
                        "TRANSFER FROM","BANK CREDIT","SALARY","CREDIT","CASHBACK",
                        "INTEREST","REFUND","SWITCHING"
                    ))
                    if is_credit:
                        amt_in = _clean_amount(amounts[0])
                    else:
                        amt_out = _clean_amount(amounts[0])
                else:
                    # No amounts on this line – it's a sub-description (merchant name, etc.)
                    # Merge into previous transaction if possible
                    if result and not desc.startswith("Effective Date"):
                        result[-1]["description"] += f" / {desc}"
                    continue

                if amt_out == 0 and amt_in == 0:
                    continue

                # Parse date
                try:
                    m = re.match(rf"(\d{{1,2}})\s+({MONTH_NAMES})", date_str)
                    if not m:
                        continue
                    day, mon = m.group(1), m.group(2)
                    date = datetime.strptime(f"{day} {mon} {stmt_year}", "%d %b %Y").date()
                except ValueError:
                    continue

                # Determine transaction type from description
                desc_up = desc.upper()
                if "CONTACTLESS" in desc_up:
                    txn_type = "Contactless"
                elif "TRANSFER FROM" in desc_up:
                    txn_type = "Transfer In"
                elif "TRANSFER TO" in desc_up or "PAYMENT TO" in desc_up:
                    txn_type = "Transfer Out"
                elif "BANK CREDIT" in desc_up:
                    txn_type = "Bank Credit"
                elif "DIRECT DEBIT" in desc_up or "Direct debit" in desc:
                    txn_type = "Direct Debit"
                elif "ATM" in desc_up:
                    txn_type = "ATM"
                elif "STANDING ORDER" in desc_up:
                    txn_type = "Standing Order"
                else:
                    txn_type = "Debit" if amt_out > 0 else "Credit"

                result.append({
                    "date":        date,
                    "type":        txn_type,
                    "description": desc,
                    "note":        "",
                    "amount_in":   amt_in,
                    "amount_out":  amt_out,
                    "balance":     balance,
                    "bank":        "Nationwide",
                    "amount":      amt_in - amt_out,
                })
            return result

        # Stream through raw_lines grouping by date
        current_date = None
        current_group: list[str] = []

        for _, _y, line in raw_lines:
            if SKIP_RE.match(line):
                continue
            tokens = line.split()
            # Check if this line starts with a date
            if len(tokens) >= 2 and DATE_RE.match(" ".join(tokens[:2])):
                if current_date and current_group:
                    transactions.extend(flush_group(current_group, current_date))
                current_date = " ".join(tokens[:2])
                current_group = [line]
            else:
                current_group.append(line)

        if current_date and current_group:
            transactions.extend(flush_group(current_group, current_date))

    return transactions, meta


def parse_csv_file(path: str) -> tuple[list[dict], dict]:
    """
    Generic CSV parser for common UK bank exports.
    Detects columns by header name matching.
    """
    DATE_FMTS = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d %b %Y", "%d/%m/%y"]

    def parse_date(s: str):
        for fmt in DATE_FMTS:
            try:
                return datetime.strptime(s.strip(), fmt).date()
            except ValueError:
                pass
        return None

    def find_col(headers: list[str], *candidates: str) -> str | None:
        hl = [h.lower().strip() for h in headers]
        for c in candidates:
            for i, h in enumerate(hl):
                if c.lower() in h:
                    return headers[i]
        return None

    transactions = []
    meta = {}

    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(path, encoding=enc, newline="") as f:
                lines = f.readlines()
            break
        except UnicodeDecodeError:
            continue

    # Skip preamble rows until we find a header row with Date
    header_idx = 0
    for i, line in enumerate(lines[:15]):
        low = line.lower()
        if "date" in low and ("description" in low or "amount" in low or "debit" in low):
            header_idx = i
            break

    reader = csv.DictReader(lines[header_idx:])
    headers = reader.fieldnames or []

    date_col   = find_col(headers, "date", "transaction date", "posted")
    desc_col   = find_col(headers, "description", "narrative", "memo", "payee",
                          "details", "transaction description", "counter party", "name")
    amt_col    = find_col(headers, "amount", "value", "net amount")
    debit_col  = find_col(headers, "debit", "paid out", "money out", "debit amount", "out")
    credit_col = find_col(headers, "credit", "paid in", "money in", "credit amount", "in")

    for row in reader:
        n = {k: (v or "").strip() for k, v in row.items() if k}
        if not any(n.values()):
            continue
        date = parse_date(n.get(date_col, "")) if date_col else None
        if date is None:
            continue
        desc = n.get(desc_col, "") if desc_col else ""
        if debit_col and credit_col:
            d = _clean_amount(n.get(debit_col, "") or "")
            c = _clean_amount(n.get(credit_col, "") or "")
            amount = c - d
        elif amt_col:
            amount = _clean_amount(n.get(amt_col, "0"))
        else:
            continue
        transactions.append({
            "date":        date,
            "type":        "Credit" if amount >= 0 else "Debit",
            "description": desc,
            "note":        "",
            "amount_in":   max(amount, 0),
            "amount_out":  max(-amount, 0),
            "balance":     0.0,
            "bank":        "CSV",
            "amount":      amount,
        })

    return transactions, meta


# ── Categorisation ────────────────────────────────────────────────────────────

def build_category_patterns(cats: dict[str, list[str]]) -> list[tuple[str, re.Pattern]]:
    return [
        (cat, re.compile("|".join(re.escape(k) for k in kws), re.IGNORECASE))
        for cat, kws in cats.items() if kws
    ]

def categorise(desc: str, patterns: list[tuple[str, re.Pattern]],
               overrides: dict[str, str]) -> str:
    if desc in overrides:
        return overrides[desc]
    for cat, pat in patterns:
        if pat.search(desc):
            return cat
    return "Other"


# ── Main Application ──────────────────────────────────────────────────────────

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Personal Finance Dashboard")
        self.root.geometry("1280x760")
        self.root.minsize(900, 600)

        self.df: pd.DataFrame | None = None
        self.meta: dict = {}
        self.current_file = ""
        self.overrides: dict[str, str] = self._load_overrides()
        self.patterns = build_category_patterns(DEFAULT_CATEGORIES)

        self._build_menu()
        self._show_welcome()

    # ── Persistence ──────────────────────────────────────────────────────────

    def _load_overrides(self) -> dict:
        try:
            if os.path.exists(OVERRIDES_FILE):
                with open(OVERRIDES_FILE) as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_overrides(self):
        try:
            with open(OVERRIDES_FILE, "w") as f:
                json.dump(self.overrides, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save overrides: {e}")

    # ── Menu ─────────────────────────────────────────────────────────────────

    def _build_menu(self):
        menu = tk.Menu(self.root)
        self.root.config(menu=menu)

        fm = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="File", menu=fm)
        fm.add_command(label="Open PDF…",  command=self._open_pdf)
        fm.add_command(label="Open CSV…",  command=self._open_csv)
        fm.add_separator()
        fm.add_command(label="Export transactions (CSV)…", command=self._export_csv)
        fm.add_command(label="Export category rules…",     command=self._export_rules)
        fm.add_command(label="Import category rules…",     command=self._import_rules)
        fm.add_separator()
        fm.add_command(label="Exit", command=self.root.quit)

        hm = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="Help", menu=hm)
        hm.add_command(label="About", command=self._show_about)

    # ── Welcome ───────────────────────────────────────────────────────────────

    def _show_welcome(self):
        self._clear()
        f = tk.Frame(self.root)
        f.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(f, text="Personal Finance Dashboard",
                 font=("Arial", 22, "bold")).pack(pady=12)
        tk.Label(f, text="Load a bank statement to begin",
                 font=("Arial", 12), fg="gray").pack()
        tk.Label(f, text="Supports: Zopa PDF · Nationwide PDF · Generic CSV",
                 font=("Arial", 10), fg="gray").pack(pady=4)

        bf = tk.Frame(f)
        bf.pack(pady=20)
        tk.Button(bf, text="Open PDF",  command=self._open_pdf,
                  font=("Arial", 12), bg="#2196F3", fg="white",
                  padx=20, pady=10).pack(side="left", padx=10)
        tk.Button(bf, text="Open CSV",  command=self._open_csv,
                  font=("Arial", 12), bg="#4CAF50", fg="white",
                  padx=20, pady=10).pack(side="left", padx=10)

    def _clear(self):
        for w in self.root.winfo_children():
            w.destroy()
        self._build_menu()

    # ── File loading ──────────────────────────────────────────────────────────

    def _open_pdf(self):
        path = filedialog.askopenfilename(
            title="Select PDF statement",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self._load_file(path)

    def _open_csv(self):
        path = filedialog.askopenfilename(
            title="Select CSV statement",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self._load_file(path)

    def _load_file(self, path: str):
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".pdf":
                # Detect bank from content
                with pdfplumber.open(path) as pdf:
                    snippet = (pdf.pages[0].extract_text() or "").lower()
                if "zopa" in snippet:
                    txns, meta = parse_zopa_pdf(path)
                    bank = "Zopa"
                elif "nationwide" in snippet or "flexgraduate" in snippet or "flexdirect" in snippet:
                    txns, meta = parse_nationwide_pdf(path)
                    bank = "Nationwide"
                else:
                    messagebox.showerror("Unsupported PDF",
                        "Could not identify bank.\n"
                        "Currently supported: Zopa, Nationwide.\n"
                        "Try exporting as CSV instead.")
                    return
            else:
                txns, meta = parse_csv_file(path)
                bank = "CSV"

            if not txns:
                messagebox.showerror("No data", "No transactions found in file.")
                return

            self.current_file = path
            self.meta = meta
            self._build_dataframe(txns, bank)
            self._show_dashboard()

        except Exception as e:
            import traceback
            messagebox.showerror("Error loading file", str(e))
            traceback.print_exc()

    def _build_dataframe(self, txns: list[dict], bank: str):
        df = pd.DataFrame(txns)
        df["date"] = pd.to_datetime(df["date"])
        df["category"] = df["description"].apply(
            lambda d: categorise(d, self.patterns, self.overrides)
        )
        df["bank"] = bank
        df = df.sort_values("date").reset_index(drop=True)
        self.df = df

    # ── Dashboard ─────────────────────────────────────────────────────────────

    def _show_dashboard(self):
        self._clear()

        # Status bar
        sb = tk.Frame(self.root, bg="#f0f0f0", relief="sunken", bd=1)
        sb.pack(fill="x", side="bottom")
        tk.Label(sb, text=f"  {os.path.basename(self.current_file)}",
                 bg="#f0f0f0", font=("Arial", 9)).pack(side="left")
        tk.Label(sb, text=f"{len(self.df)} transactions  ",
                 bg="#f0f0f0", font=("Arial", 9)).pack(side="right")

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=5, pady=5)

        tabs = [
            ("Summary",       self._tab_summary),
            ("Transactions",  self._tab_transactions),
            ("By Category",   self._tab_category),
            ("Charts",        self._tab_charts),
        ]
        for label, builder in tabs:
            frame = ttk.Frame(nb)
            nb.add(frame, text=label)
            builder(frame)

    # ── Tab: Summary ──────────────────────────────────────────────────────────

    def _tab_summary(self, parent):
        df = self.df
        is_internal = df["category"] == "Internal Transfer"
        fin = df[~is_internal]

        total_in  = fin["amount_in"].sum()
        total_out = fin["amount_out"].sum()
        net       = total_in - total_out
        n_int     = is_internal.sum()

        # Account info card
        info = tk.Frame(parent, bg="#E3F2FD", relief="ridge", bd=2)
        info.pack(fill="x", padx=12, pady=10)
        holder  = self.meta.get("holder", df["bank"].iloc[0] + " account")
        balance = self.meta.get("closing_balance", "—")
        tk.Label(info, text=holder,  font=("Arial", 15, "bold"), bg="#E3F2FD").pack(pady=6)
        tk.Label(info, text=f"Closing balance: £{balance}",
                 font=("Arial", 11), bg="#E3F2FD").pack(pady=2)

        # Stat cards
        cards_data = [
            ("Total in",     f"£{total_in:,.2f}",  "#C8E6C9"),
            ("Total out",    f"£{total_out:,.2f}", "#FFCDD2"),
            ("Net",          f"£{net:,.2f}",        "#FFF9C4"),
            ("Transactions", str(len(df)),           "#E1BEE7"),
        ]
        cf = tk.Frame(parent)
        cf.pack(fill="x", padx=12, pady=6)
        for i, (label, val, colour) in enumerate(cards_data):
            c = tk.Frame(cf, bg=colour, relief="raised", bd=2)
            c.grid(row=0, column=i, padx=8, pady=6, sticky="nsew")
            tk.Label(c, text=label, font=("Arial", 11),    bg=colour).pack(pady=8)
            tk.Label(c, text=val,   font=("Arial", 15, "bold"), bg=colour).pack(pady=8)
        for i in range(4):
            cf.columnconfigure(i, weight=1)

        if n_int:
            ib = tk.Frame(parent, bg="#FFFDE7", relief="flat", bd=1)
            ib.pack(fill="x", padx=12)
            tk.Label(ib, text=f"ℹ  {n_int} internal transfers excluded from totals",
                     font=("Arial", 9), fg="#777", bg="#FFFDE7").pack(pady=4)

        # Category breakdown table
        tk.Label(parent, text="Spending by category",
                 font=("Arial", 12, "bold")).pack(anchor="w", padx=14, pady=(10, 2))

        tf = tk.Frame(parent)
        tf.pack(fill="both", expand=True, padx=12, pady=6)

        cols = ("Category", "Transactions", "Total out")
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=12)
        for c in cols:
            tree.heading(c, text=c)
        tree.column("Category",     width=180)
        tree.column("Transactions", width=110, anchor="center")
        tree.column("Total out",    width=120, anchor="e")

        sb2 = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb2.set)
        tree.pack(side="left", fill="both", expand=True)
        sb2.pack(side="right", fill="y")

        expenses = fin[fin["amount_out"] > 0]
        grp = expenses.groupby("category")["amount_out"].agg(["count", "sum"]) \
                      .sort_values("sum", ascending=False)
        for cat, row in grp.iterrows():
            tree.insert("", "end", values=(cat, int(row["count"]), f"£{row['sum']:,.2f}"))

    # ── Tab: Transactions ─────────────────────────────────────────────────────

    def _tab_transactions(self, parent):
        # Search bar
        sf = tk.Frame(parent)
        sf.pack(fill="x", padx=8, pady=5)

        tk.Label(sf, text="Search:").pack(side="left", padx=4)
        sv = tk.StringVar()
        tk.Entry(sf, textvariable=sv, width=38).pack(side="left", padx=4)

        tk.Label(sf, text="in").pack(side="left")
        col_var = tk.StringVar(value="Description")
        ttk.Combobox(sf, textvariable=col_var,
                     values=["Description", "Category", "Type", "All"],
                     width=14).pack(side="left", padx=4)

        count_lbl = tk.Label(sf, text="", fg="gray", font=("Arial", 9))
        count_lbl.pack(side="right", padx=8)

        # Treeview
        tf = tk.Frame(parent)
        tf.pack(fill="both", expand=True, padx=8, pady=4)

        cols = ("Date", "Type", "Description", "Out", "In", "Balance", "Category")
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=22)
        widths = dict(Date=90, Type=110, Description=310, Out=90, In=90, Balance=90, Category=120)
        for c in cols:
            tree.heading(c, text=c,
                         command=lambda _c=c: self._sort_tree(tree, _c, False))
            tree.column(c, width=widths[c],
                        anchor="e" if c in ("Out", "In", "Balance") else "w")

        vsb = ttk.Scrollbar(tf, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)

        tree.tag_configure("internal", foreground="#aaa")
        tree.tag_configure("credit",   foreground="#2e7d32")

        def populate(rows: pd.DataFrame):
            tree.delete(*tree.get_children())
            for _, r in rows.iterrows():
                out = f"£{r.amount_out:,.2f}" if r.amount_out > 0 else ""
                inn = f"£{r.amount_in:,.2f}"  if r.amount_in  > 0 else ""
                bal = f"£{r.balance:,.2f}"     if r.balance   != 0 else ""
                tag = "internal" if r.category == "Internal Transfer" else \
                      "credit"   if r.amount_in > 0                  else ""
                tree.insert("", "end",
                            iid=str(r.name),
                            values=(r.date.strftime("%d %b %Y"),
                                    r.type,
                                    r.description[:70],
                                    out, inn, bal,
                                    r.category),
                            tags=(tag,))
            count_lbl.config(text=f"{len(rows)} transactions")

        def search(_=None):
            term = sv.get().lower()
            col  = col_var.get()
            if not term:
                populate(self.df)
                return
            df2 = self.df
            if col == "Description":
                mask = df2["description"].str.lower().str.contains(term, na=False)
            elif col == "Category":
                mask = df2["category"].str.lower().str.contains(term, na=False)
            elif col == "Type":
                mask = df2["type"].str.lower().str.contains(term, na=False)
            else:
                mask = (
                    df2["description"].str.lower().str.contains(term, na=False) |
                    df2["category"].str.lower().str.contains(term, na=False)    |
                    df2["type"].str.lower().str.contains(term, na=False)
                )
            populate(df2[mask])

        sv.trace_add("write", search)
        col_var.trace_add("write", search)
        populate(self.df)

        # Right-click → change category
        def on_right_click(event):
            item = tree.identify_row(event.y)
            if not item:
                return
            tree.selection_set(item)
            row = self.df.loc[int(item)]
            self._change_category(row["description"], row["category"],
                                  on_change=lambda: (populate(self.df), search()))
            context_menu.post(event.x_root, event.y_root)

        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(
            label="Change category…",
            command=lambda: [
                self._change_category(
                    self.df.loc[int(tree.selection()[0]), "description"],
                    self.df.loc[int(tree.selection()[0]), "category"],
                    on_change=lambda: (populate(self.df), search())
                ) if tree.selection() else None
            ]
        )
        tree.bind("<Button-3>", on_right_click)

    def _sort_tree(self, tree: ttk.Treeview, col: str, reverse: bool):
        items = [(tree.set(k, col), k) for k in tree.get_children("")]
        try:
            items.sort(key=lambda t: float(t[0].replace("£", "").replace(",", "")),
                       reverse=reverse)
        except ValueError:
            items.sort(reverse=reverse)
        for i, (_, k) in enumerate(items):
            tree.move(k, "", i)
        tree.heading(col, command=lambda: self._sort_tree(tree, col, not reverse))

    def _change_category(self, description: str, current: str, on_change=None):
        cats = sorted(set(CATEGORY_ORDER) | set(self.df["category"].unique()))
        dlg  = tk.Toplevel(self.root)
        dlg.title("Change category")
        dlg.geometry("380x240")
        dlg.resizable(False, False)
        dlg.grab_set()

        tk.Label(dlg, text=description[:60], font=("Arial", 10, "bold"),
                 wraplength=340).pack(pady=10)
        tk.Label(dlg, text=f"Current: {current}", fg="blue").pack()

        cv = tk.StringVar(value=current)
        combo = ttk.Combobox(dlg, textvariable=cv, values=cats, width=32)
        combo.pack(pady=8)

        tk.Label(dlg, text="Or type a new category:").pack()
        custom = tk.Entry(dlg, width=34)
        custom.pack(pady=4)

        def apply():
            new_cat = custom.get().strip() or cv.get()
            if not new_cat:
                return
            self.overrides[description] = new_cat
            self._save_overrides()
            self.df.loc[self.df["description"] == description, "category"] = new_cat
            if on_change:
                on_change()
            dlg.destroy()

        tk.Button(dlg, text="Apply", command=apply,
                  bg="#4CAF50", fg="white", padx=16, pady=6).pack(pady=8)

    # ── Tab: Category filter ──────────────────────────────────────────────────

    def _tab_category(self, parent):
        cf = tk.Frame(parent)
        cf.pack(fill="x", padx=8, pady=6)

        tk.Label(cf, text="Category:").pack(side="left", padx=4)
        cats = ["All"] + sorted(self.df["category"].unique())
        cv = tk.StringVar(value="All")
        ttk.Combobox(cf, textvariable=cv, values=cats, width=22).pack(side="left", padx=4)

        tk.Label(cf, text="Direction:").pack(side="left", padx=4)
        dv = tk.StringVar(value="All")
        ttk.Combobox(cf, textvariable=dv,
                     values=["All", "Money in", "Money out"],
                     width=12).pack(side="left", padx=4)

        stat_lbl = tk.Label(cf, text="", font=("Arial", 10))
        stat_lbl.pack(side="right", padx=8)

        tf = tk.Frame(parent)
        tf.pack(fill="both", expand=True, padx=8, pady=4)

        cols = ("Date", "Description", "Category", "Out", "In")
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=22)
        for c in cols:
            tree.heading(c, text=c)
        tree.column("Date",        width=90)
        tree.column("Description", width=350)
        tree.column("Category",    width=150)
        tree.column("Out",         width=100, anchor="e")
        tree.column("In",          width=100, anchor="e")

        vsb = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def refresh(_=None):
            data = self.df.copy()
            if cv.get() != "All":
                data = data[data["category"] == cv.get()]
            if dv.get() == "Money in":
                data = data[data["amount_in"] > 0]
            elif dv.get() == "Money out":
                data = data[data["amount_out"] > 0]

            tree.delete(*tree.get_children())
            for _, r in data.iterrows():
                tree.insert("", "end", values=(
                    r.date.strftime("%d %b %Y"),
                    r.description[:65],
                    r.category,
                    f"£{r.amount_out:,.2f}" if r.amount_out > 0 else "",
                    f"£{r.amount_in:,.2f}"  if r.amount_in  > 0 else "",
                ))

            fin_data = data[data["category"] != "Internal Transfer"]
            t_in  = fin_data["amount_in"].sum()
            t_out = fin_data["amount_out"].sum()
            stat_lbl.config(text=f"In: £{t_in:,.2f}  |  Out: £{t_out:,.2f}  |  {len(data)} rows")

        cv.trace_add("write", refresh)
        dv.trace_add("write", refresh)
        tk.Button(cf, text="Refresh", command=refresh).pack(side="left", padx=6)
        refresh()

    # ── Tab: Charts ───────────────────────────────────────────────────────────

    def _tab_charts(self, parent):
        nb2 = ttk.Notebook(parent)
        nb2.pack(fill="both", expand=True)

        pie_frame = ttk.Frame(nb2)
        nb2.add(pie_frame, text="Spending by category")
        bar_frame = ttk.Frame(nb2)
        nb2.add(bar_frame, text="Monthly in vs out")

        fin = self.df[self.df["category"] != "Internal Transfer"]

        # ── Pie chart ───────────────────────────────────────────────────────
        fig1, ax1 = plt.subplots(figsize=(6, 5), tight_layout=True)
        expenses = fin[fin["amount_out"] > 0].groupby("category")["amount_out"].sum()
        expenses = expenses[expenses > 0].sort_values(ascending=False)
        if not expenses.empty:
            wedge_props = {"linewidth": 0.5, "edgecolor": "white"}
            ax1.pie(expenses.values,
                    labels=[f"{c}\n£{v:,.0f}" for c, v in expenses.items()],
                    autopct="%1.0f%%", startangle=140,
                    wedgeprops=wedge_props, textprops={"fontsize": 8})
            ax1.set_title("Expenses by category", fontsize=12)
        FigureCanvasTkAgg(fig1, pie_frame).get_tk_widget().pack(fill="both", expand=True)

        # ── Bar chart ────────────────────────────────────────────────────────
        fig2, ax2 = plt.subplots(figsize=(8, 5), tight_layout=True)
        fin2 = fin.copy()
        fin2["month"] = fin2["date"].dt.to_period("M").astype(str)
        m_in  = fin2.groupby("month")["amount_in"].sum()
        m_out = fin2.groupby("month")["amount_out"].sum()
        months = sorted(set(m_in.index) | set(m_out.index))
        x = range(len(months))
        w = 0.35
        b1 = ax2.bar([i - w/2 for i in x],
                     [m_in.get(m, 0)  for m in months], w,
                     label="In",  color="#43A047", alpha=0.85)
        b2 = ax2.bar([i + w/2 for i in x],
                     [m_out.get(m, 0) for m in months], w,
                     label="Out", color="#E53935", alpha=0.85)
        ax2.set_xticks(list(x))
        ax2.set_xticklabels(months, rotation=30, ha="right")
        ax2.set_ylabel("£")
        ax2.set_title("Monthly income vs spending")
        ax2.legend()
        ax2.bar_label(b1, fmt="£%.0f", fontsize=7, padding=2)
        ax2.bar_label(b2, fmt="£%.0f", fontsize=7, padding=2)
        FigureCanvasTkAgg(fig2, bar_frame).get_tk_widget().pack(fill="both", expand=True)

    # ── Export / Import ───────────────────────────────────────────────────────

    def _export_csv(self):
        if self.df is None:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if path:
            out = self.df.copy()
            out["date"] = out["date"].dt.strftime("%d/%m/%Y")
            out.to_csv(path, index=False)
            messagebox.showinfo("Exported", f"Saved to {path}")

    def _export_rules(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")]
        )
        if path:
            with open(path, "w") as f:
                json.dump(self.overrides, f, indent=2)
            messagebox.showinfo("Exported", f"Rules saved to {path}")

    def _import_rules(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            try:
                with open(path) as f:
                    imported = json.load(f)
                self.overrides.update(imported)
                self._save_overrides()
                if self.df is not None:
                    self.df["category"] = self.df["description"].apply(
                        lambda d: categorise(d, self.patterns, self.overrides)
                    )
                    self._show_dashboard()
                messagebox.showinfo("Imported", f"Loaded {len(imported)} rules.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _show_about(self):
        messagebox.showinfo("About",
            "Personal Finance Dashboard\n\n"
            "Supported formats:\n"
            "  • Zopa Bank PDF statements\n"
            "  • Nationwide PDF statements\n"
            "  • Generic CSV exports (most UK banks)\n\n"
            "Right-click a transaction to change its category.\n"
            "Changes are saved automatically."
        )


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()