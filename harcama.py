# -*- coding: utf-8 -*-
"""
Harcama takibi — kredi kartı ekstrelerinden kalıcı, kategorize edilmiş harcama defteri.

Kullanım:
    python harcama.py           # ekstreler/ içindeki YENİ PDF'leri içe aktar + Harcamalarim.xlsx'i yeniden üret
    python harcama.py --build   # sadece Excel'i yeniden üret (kural / kur / etiket değişikliğinden sonra)
    python app.py               # yerel uygulama: tarayıcıdan dosya yükle, kategorile, bütçe gir (Harcamalarim.bat)

Belgeler/ klasöründeki tüm dosyalar (kart ekstresi PDF, hesap dökümü PDF/XLSX; alt klasörler dahil) okunur; daha önce
işlenenler içerik özetinden (SHA-256) tanınır ve atlanır. Veri klasörü: home.txt (veya HARCAMA_HOME).

Asıl kayıt data/ klasöründeki CSV'lerdir; Harcamalarim.html (panel) ve Harcamalarim.xlsx (ham defter) her seferinde
bu CSV'lerden üretilir.
  data/transactions.csv   tüm işlemler (merchant_override / category_override / note sütunları elle düzenlenebilir)
  data/statements.csv     içe aktarılan ekstreler (aynı ekstre iki kez alınmaz)
  data/fx_rates.csv       EUR/USD → TRY kurları (TCMB döviz satış, son ödeme tarihi)
  data/category_rules.csv işyeri → kategori kuralları (ilk eşleşen kazanır)
  data/tags.csv           seyahat / olay etiketleri (tarih aralığı)
"""
from __future__ import annotations

import calendar
import csv
import datetime as dt
import difflib
import hashlib
import os
import re
import statistics
import subprocess
import sys
import urllib.request
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

import pdfplumber
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule

from dashboard import build_payload, write_dashboard

sys.stdout.reconfigure(encoding="utf-8")

# HARCAMA_HOME = where your private files live (data/, ekstreler/, hesap_dokumu/, inbox/, outputs); the code can sit
# elsewhere. Unset → the code folder itself. `python demo.py` builds a fake home in demo/ for developing features.
CODE_DIR = Path(__file__).resolve().parent


def _home() -> Path:
    """HARCAMA_HOME, else the path written in home.txt next to the code, else the code folder."""
    if os.environ.get("HARCAMA_HOME"):
        return Path(os.environ["HARCAMA_HOME"])
    cfg = CODE_DIR / "home.txt"
    if cfg.exists():
        line = next((ln.strip() for ln in cfg.read_text(encoding="utf-8-sig").splitlines()
                     if ln.strip() and not ln.lstrip().startswith("#")), "")
        if line:
            return Path(os.path.expandvars(line))
    return CODE_DIR


ROOT = _home().resolve()
DOCS_DIR = ROOT / "Belgeler"            # <- put every statement / export here (subfolders are fine)
INBOX_DIR = ROOT / "inbox"              # older layout, still read
STATEMENT_DIR = ROOT / "ekstreler"
ACCOUNT_DIR = ROOT / "hesap_dokumu"
DATA = ROOT / "data"
TX_CSV = DATA / "transactions.csv"
ST_CSV = DATA / "statements.csv"
FX_CSV = DATA / "fx_rates.csv"
RULES_CSV = DATA / "category_rules.csv"
TAGS_CSV = DATA / "tags.csv"
PEOPLE_CSV = DATA / "people.csv"
INCOME_PLAN_CSV = DATA / "income_plan.csv"
BUDGETS_CSV = DATA / "budgets.csv"
NO_ACTIVITY_CSV = DATA / "confirmed_no_activity.csv"
EXCLUDED_SOURCES_CSV = DATA / "excluded_coverage_sources.csv"
FILES_CSV = DATA / "files.csv"          # every document seen: content hash -> what happened to it
OUT_XLSX = ROOT / "Harcamalarim.xlsx"
OUT_HTML = ROOT / "Harcamalarim.html"

TX_FIELDS = ["id", "statement_id", "bank", "card", "date", "description", "bank_category", "section", "type",
             "currency", "amount", "orig_amount", "orig_currency", "installment", "bonus", "payment_date",
             "source_file", "merchant_override", "category_override", "subcategory_override", "note",
             "ext_id", "counterparty", "memo"]
ST_FIELDS = ["statement_id", "bank", "card", "currency", "ekstre_no", "period_start", "cutoff", "payment_date",
             "prev_balance", "spending", "fees_interest", "payments_credits", "balance_due", "min_payment",
             "source_file", "kind"]
FX_FIELDS = ["key", "currency", "payment_date", "rate", "rate_date", "source", "status"]
FILES_FIELDS = ["sha256", "file", "kind", "result", "date"]

# Purchases/fees on cards, plus from the bank account: money sent to people, friends paying you back (negative),
# and ATM cash. Card payments, own-account transfers, fund moves, salary and cash-advance principal are not spending.
SPENDING_TYPES = {"Purchase", "Refund", "Discount", "Fee", "Transfer out", "Reimbursement", "Cash withdrawal"}
# Cards/accounts already in use before their first imported statement: earlier days count as "data missing",
# not "card not in use yet". Anything else counts from its first statement period (e.g. the new QNB card).
# Local-only configuration: one exact source name per line. Never commit account identifiers.
_early_sources = CODE_DIR / "coverage_sources.txt"
IN_USE_BEFORE_FIRST_STATEMENT = {
    line.strip() for line in _early_sources.read_text(encoding="utf-8-sig").splitlines()
    if line.strip() and not line.lstrip().startswith("#")
} if _early_sources.exists() else set()
# Categories that aren't shops: kept off the Merchants sheet (they get their own Insights sections).
STORE_EXCLUDE = {"People (IBAN)", "Cash"}
NON_SPENDING_CATEGORIES = {"Cash Advance", "Transfers & Payments"}
TR_MONTHS = ["OCAK", "SUBAT", "MART", "NISAN", "MAYIS", "HAZIRAN", "TEMMUZ", "AGUSTOS", "EYLUL", "EKIM",
             "KASIM", "ARALIK"]
WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
# Card rotation habit (maximises the interest-free period): QNB from the day after its cutoff (1st) to the 14th,
# Garanti from the 16th through the 1st; the 15th (around Garanti's cutoff) counts for either. Checked from this date
# on (Aug 2026 was the exception: brand-new QNB card).
ROTATION_START = dt.date(2026, 9, 1)
ROTATION = [(2, 14, "QNB"), (16, 31, "Garanti"), (1, 1, "Garanti")]
ROTATION_TEXT = ("Your rule: QNB from the 2nd to the 14th, Garanti from the 16th to the 1st (the 15th counts for "
                 "either) — that gives every purchase the longest interest-free period.")
_TR = str.maketrans("İIıŞşĞğÜüÖöÇç", "IIiSsGgUuOoCc")
TRN = r"\d{1,3}(?:\.\d{3})*,\d{2}"          # Turkish number: 1.250,00
ENN = r"-?\d{1,3}(?:,\d{3})*\.\d{2}"        # English number: 1,250.00


# ───────────────────────────── helpers ─────────────────────────────
def fold(s: str) -> str:
    """Turkish-safe ASCII uppercase, used for all matching."""
    return s.translate(_TR).upper()


def month_no(word: str):
    w = fold(re.sub(r"\(cid:\d+\)", "", word))
    if w in TR_MONTHS:
        return TR_MONTHS.index(w) + 1
    m = difflib.get_close_matches(w, TR_MONTHS, n=1, cutoff=0.6)
    return TR_MONTHS.index(m[0]) + 1 if m else None


def tr_date(d, mon, y):
    m = month_no(mon)
    return dt.date(int(y), m, int(d)) if m else None


def tr_num(s: str) -> float:
    return float(s.replace(".", "").replace(",", "."))


def en_num(s: str) -> float:
    return float(s.replace(",", ""))


def add_months(d: dt.date, n: int) -> dt.date:
    y = d.year + (d.month - 1 + n) // 12
    m = (d.month - 1 + n) % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def iso(d):
    return d.isoformat() if isinstance(d, dt.date) else (d or "")


def to_date(s):
    if isinstance(s, dt.date):
        return s
    return dt.date.fromisoformat(s) if s else None


def read_csv(path):
    if not path.exists():
        return []
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_csv(path, fields, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: iso(r.get(k, "")) if isinstance(r.get(k), dt.date) else r.get(k, "") for k in fields})


def money(x):
    return "" if x in ("", None) else f"{float(x):.2f}"


def tr_title(s: str) -> str:
    """Turkish-aware title case: 'HAMİT EFE ÇINAR' -> 'Hamit Efe Çınar'."""
    low = s.replace("I", "ı").replace("İ", "i").lower()
    first = {"i": "İ", "ı": "I"}
    return " ".join(first.get(w[0], w[0].upper()) + w[1:] for w in low.split())


def acc_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return dt.datetime.strptime(str(v).strip()[:10], "%d/%m/%Y").date()


def acc_num(v) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("TL", "").replace(" ", "").strip()
    if "," in s and "." in s:
        return tr_num(s) if s.rfind(",") > s.rfind(".") else en_num(s)
    return float(s.replace(",", "."))


# ───────────────────────────── parsers ─────────────────────────────
def parse_garanti(text: str, fname: str):
    F = fold(text)
    cur = re.search(r"TUTAR \((TL|EUR|USD)\)", F).group(1)
    cur = "TRY" if cur == "TL" else cur
    last4 = re.search(r"KART NUMARASI \d{4} \d{2}\*\* \*{4} (\d{4})", F).group(1)
    cutoff = tr_date(*re.search(r"HESAP KESIM TARIHI (\d{1,2}) (\S+) (\d{4})", F).groups())
    due = tr_date(*re.search(r"SON ODEME TARIHI:? (\d{1,2}) (\S+) (\d{4})", F).groups())
    m = re.search(r"ONCEKI HESAP KESIM TARIHINIZ (\d{1,2}) (\S+)\s+(\d{4})", F)
    prev_cut = tr_date(*m.groups()) if m else None
    ekstre_no = re.search(r"EKSTRE NO : (\S+)", F).group(1)
    m = re.search(rf"({TRN}) \w+ \+ ({TRN}) \w+ \+ ({TRN}) \w+ - ({TRN}) \w+ = ({TRN}) \w+ ({TRN})", F)
    prev, spend, fees, pays, bal, minp = map(tr_num, m.groups())
    card = f"Garanti Bonus •{last4}"
    st = dict(statement_id=f"GARANTI-{last4}-{cur}-{cutoff}", bank="Garanti BBVA", card=card, currency=cur,
              ekstre_no=ekstre_no,
              period_start=prev_cut + dt.timedelta(days=1) if prev_cut and (cutoff - prev_cut).days <= 40 else "",
              cutoff=cutoff, payment_date=due, prev_balance=prev, spending=spend, fees_interest=fees,
              payments_credits=pays, balance_due=bal, min_payment=minp, source_file=fname)

    rows, in_tx, section, bank_cat = [], False, "Domestic", ""
    for raw in text.splitlines():
        line = re.sub(r"\s*(bosluk|bboosslluukk)\s*$", "", raw.strip(), flags=re.I).strip()
        if not line:
            continue
        f = fold(line)
        if f.startswith("ISLEM TARIHI"):
            in_tx = True
            continue
        if not in_tx:
            continue
        if f.startswith("TOPLAM "):
            in_tx = False
            continue
        m = re.match(r"^(\d{1,2}) (\S+) (\d{4}) (.+)$", line)
        d = tr_date(*m.groups()[:3]) if m else None
        if not d:
            fm = re.match(rf"^(.*\S)\s+({TRN})$", line)
            if fm and re.search(r"FAIZ|KKDF|BSMV|UCRET|MASRAF", f) and "DEVIR" not in f:
                # interest / fee lines at the top of the statement carry no date: book them on the cutoff
                rows.append(dict(statement_id=st["statement_id"], bank=st["bank"], card=card, date=cutoff,
                                 description=fm.group(1), bank_category="", section=section, type="Fee",
                                 currency=cur, amount=tr_num(fm.group(2)), orig_amount="", orig_currency="",
                                 installment="", bonus="", payment_date=due, source_file=fname))
                continue
            if "YURT DISI HARCAMALARINIZ" in f:
                section, bank_cat = "Foreign", ""
            elif "NAKIT AVANS" in f and len(f) < 40:
                section, bank_cat = "Cash advance", ""
            elif "BONUS PROGRAM ORTAKLARI" in f:
                section, bank_cat = "Domestic", ""
            elif "DIGER HARCAMALARINIZ" in f:
                section, bank_cat = "Domestic", "Diğer"
            elif len(line) < 35 and not re.search(r"\d", line):
                bank_cat = line
            continue

        rest = m.group(4).strip()
        redeemed = re.search(rf"\s({TRN})-$", rest)
        if redeemed:  # "BONUS BEDAVA ALIŞVERİŞ 560,00-": paid with Bonus points (bonus column), no TL charged
            rows.append(dict(statement_id=st["statement_id"], bank=st["bank"], card=card, date=d,
                             description=rest[:redeemed.start()].strip(), bank_category=bank_cat, section=section,
                             type="Bonus used", currency=cur, amount=0.0, orig_amount="", orig_currency="",
                             installment="", bonus=-tr_num(redeemed.group(1)), payment_date=due, source_file=fname))
            continue
        credit = rest.endswith("+")
        rest = rest.rstrip("+").strip()
        am = re.search(rf"({TRN})$", rest)
        if not am:
            continue
        amount = tr_num(am.group(1))
        rest = rest[:am.start()].strip()
        inst = ""
        im = re.search(rf"({TRN})x(\d+)=({TRN})\s+(\d+)\.\s*Taksit", rest, re.I)
        if im:
            inst = f"{im.group(4)}/{im.group(2)} of {tr_num(im.group(3)):.2f}"
            rest = rest[:im.start()].strip()
        orig_amt = orig_cur = ""
        om = re.search(rf"({TRN}) (TL|USD|EUR|GBP|CHF)$", rest)
        if om:
            orig_amt, orig_cur = tr_num(om.group(1)), ("TRY" if om.group(2) == "TL" else om.group(2))
            rest = rest[:om.start()].strip()
        bonus = ""
        bm = re.search(rf"\s({TRN})$", rest)
        if bm:
            bonus = tr_num(bm.group(1))
            rest = rest[:bm.start()].strip()

        fd = fold(rest)
        if "DEVIR BAKIYESI" in fd or "TLYE DONEN" in fd:
            # unpaid EUR/USD balance converted into TL debt: "TLye dönen EUR Borç Ödeme" (FX statement) and
            # "EUR ekstre devir bakiyesi" (TL statement). Not new spending.
            typ = "FX balance transfer"
        elif "ODEMENIZ ICIN TESEKKUR" in fd:
            typ = "Payment"
        elif section == "Cash advance" or "NAKIT AVANS" in fd:
            typ = "Cash advance"
        elif re.search(r"UCRET|FAIZ|KKDF|BSMV|GECIKME", fd):
            typ = "Fee"
        elif credit:
            typ = "Refund"
        else:
            typ = "Purchase"
        rows.append(dict(statement_id=st["statement_id"], bank=st["bank"], card=card, date=d, description=rest,
                         bank_category=bank_cat, section=section, type=typ, currency=cur,
                         amount=-amount if credit else amount, orig_amount=orig_amt, orig_currency=orig_cur,
                         installment=inst, bonus=bonus, payment_date=due, source_file=fname))
    return st, rows


def parse_qnb(text: str, fname: str):
    t = re.sub(r"\(cid:\d+\)", "", text)
    last4 = re.search(r"\d{4} \d{2}\*\* \*{4} (\d{4})", t).group(1)
    cutoff = dt.datetime.strptime(re.search(r"Hesap Kesim Tarihi\s*:\s*(\d{2}/\d{2}/\d{4})", t).group(1),
                                  "%d/%m/%Y").date()
    m = re.search(r"deme Tarihi:\s*(\d{1,2})\s+(\S+)\s+(\d{4})", t)
    due = (tr_date(*m.groups()) if m else None) or cutoff + dt.timedelta(days=10)
    m = re.search(rf"Asgari\s+\S*deme\s+Tutar\S*\s*:\s*({ENN})", t)
    minp = en_num(m.group(1)) if m else ""
    m = re.search(r"Seri-S\S*\s*No:\s*(\S+)", t)
    ekstre_no = m.group(1) if m else ""
    m = re.search(rf"^({ENN})\s+({ENN})\s+({ENN})\s+({ENN})\s+({ENN})\s+({ENN})\s+({ENN})\s+TL\s*$", t, re.M)
    prev, spend, fees, interest, tax, pays, bal = map(en_num, m.groups())
    card = f"QNB Fix •{last4}"
    st = dict(statement_id=f"QNB-{last4}-TRY-{cutoff}", bank="QNB", card=card, currency="TRY",
              ekstre_no=ekstre_no, period_start="", cutoff=cutoff, payment_date=due, prev_balance=prev,
              spending=spend, fees_interest=fees + interest + tax, payments_credits=-pays, balance_due=bal,
              min_payment=minp, source_file=fname)
    rows = []
    for line in t.splitlines():
        m = re.match(rf"^(\d{{2}}/\d{{2}}/\d{{4}})\s+(.+?)\s+({ENN})(?:\s+(\d+/\d+))?\s*$", line.strip())
        if not m:
            continue
        d = dt.datetime.strptime(m.group(1), "%d/%m/%Y").date()
        desc, amount, inst = m.group(2).strip(), en_num(m.group(3)), m.group(4) or ""
        disc = re.match(rf"^(.*?)\s+({ENN})\s+al\S*\s+ind\b", desc)
        fd = fold(desc)
        if disc:
            typ = "Discount"
        elif amount < 0 and re.search(r"DEME", fd):
            typ = "Payment"
        elif amount < 0:
            typ = "Refund"
        elif re.search(r"UCRET|FAIZ|KKDF|BSMV", fd):
            typ = "Fee"
        elif "NAKIT" in fd:
            typ = "Cash advance"
        else:
            typ = "Purchase"
        rows.append(dict(statement_id=st["statement_id"], bank="QNB", card=card, date=d, description=desc,
                         bank_category="", section="Domestic", type=typ, currency="TRY", amount=amount,
                         orig_amount="", orig_currency="", installment=inst, bonus="", payment_date=due,
                         source_file=fname))
    return st, rows


def parse_pdf(path: Path):
    with pdfplumber.open(path) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    head = fold(text[:3000])
    if "QNB" in head:
        return parse_qnb(text, path.name)
    if "BONUS" in head or "GARANTI" in fold(text):
        return parse_garanti(text, path.name)
    print(f"  ! {path.name}: bank format not recognised — ask Claude to add a parser for it")
    return None


# ───────────────────────────── bank account (hesap hareketleri) ─────────────────────────────
def counterparty(desc: str):
    """Garanti transfer text -> (name, memo). Formats seen:
    'NEŞENAZ YALÇIN-stockholm lets goo-FAST-8340154'   'FAST4676776739-UMUT KÜÇÜKKAYHAN-otel borc'
    'EF9035666 SABANCI ÜNİVERSİTESİ-TEOMAN ARABUL'      'SALİH KAYA--HVL-CEP ŞUBE'
    'CEP ŞUBE-HVL-      -UĞUR CAN KAVASO'               'MOBIL-FAST-2480000267 memo' (no name at all)"""
    parts = [p.strip() for p in desc.split("-")]
    f0 = fold(parts[0])
    m = re.match(r"^EFT?\d+\s+(.+)$", parts[0])
    if m:
        return m.group(1).strip(), ""
    if re.fullmatch(r"FAST\d+", f0):
        return (parts[1] if len(parts) > 1 else ""), "-".join(parts[2:]).strip("- ")
    if f0 == "CEP SUBE":
        return parts[-1] if not re.match(r"^\d", parts[-1]) else "", ""
    if f0 == "MOBIL":
        return "", re.sub(r"^\d+\s*", "", parts[-1])
    idx = next((i for i, p in enumerate(parts) if fold(p) in ("FAST", "EFT", "HAVALE", "HVL")), len(parts))
    return parts[0], "-".join(p for p in parts[1:idx] if p)


def classify_account_row(desc: str, label: str, amount: float, holder: str, name: str | None = None) -> str:
    f, e = fold(desc), fold(label or "")
    if ("KART ODEME" in e or "K.KARTI ODEME" in f or "KART ODEMESI" in f or "EKSTRE BORC TAHSILATI" in f
            or "KK TAHSILAT" in f or "KK OTOMATIK ODEME" in f):
        return "Payment"
    if "BEDELLI ASKER" in f:
        return "One-off"
    if "OGRENIM KREDISI" in f or "YURTLAR KURUMU" in f:
        return "Student loan"
    if "STOPAJ" in f or "MKK" in f:
        return "Investment cost"
    if (e == "YATIRIM" or "FON ALIS" in f or "FON SATIS" in f or "BINANCE" in f or "KRIPTO" in f
            or ("ALTIN" in f and "MEVDUAT" in f)):
        return "Investment"
    if "MAAS" in e or "MAAS" in f or "YOL YARDIMI" in f:
        return "Income"
    if "NAKIT AVANS" in e or "NAKIT AVANS" in f:
        return "Cash advance received"
    if "PARA CEKME" in e or "P.CEKME" in f or "ATM P.C" in f:
        return "Cash withdrawal"
    if "KESINTI" in f or "UCRET" in f or "KOMISYON" in f or "BSMV" in f:
        return "Fee"
    if "DOVIZ" in e:
        return "Currency exchange"
    if ("PARA TRANSFERI" in e or "TRANSFER ISLEMLERI" in f or "GOND:" in f or "FAST ISLEMI" in f
            or re.search(r"-(FAST|EFT|HAVALE|HVL)\b", f)):
        who, me = fold(name if name is not None else counterparty(desc)[0]), fold(holder)
        if who and who == me:
            return "Own transfer"
        if me and who.split()[-1:] == me.split()[-1:]:  # same surname → family
            return "Family support" if amount > 0 else "Transfer out"
        if re.search(r"UNIVERSITE|A\.S\.|LTD|SIRKETI|KURUMU|BANKASI", who):  # an institution, not a person
            return "Other income" if amount > 0 else "Purchase"
        return "Transfer out" if amount < 0 else "Reimbursement"
    return "Purchase" if amount < 0 else "Other income"


def parse_garanti_account(path: Path, known_ext: set):
    """Garanti BBVA 'hesap hareketleri' Excel export. Returns only rows not imported before."""
    ws = load_workbook(path, data_only=True).worksheets[0]
    meta, header, data = {}, None, []
    for row in ws.iter_rows(values_only=True):
        cells = ["" if v is None else v for v in row]
        if header is None:
            if str(cells[0]).strip() == "Tarih":
                header = [str(c).strip() for c in cells]
            elif cells[0] != "" and len(cells) > 1:
                meta[str(cells[0]).strip()] = cells[1]
            continue
        if cells[0] != "":
            data.append(dict(zip(header, cells)))
    if header is None or "IBAN" not in meta:
        return None
    holder = str(meta.get("Ad Soyad", "")).strip()
    last4 = re.sub(r"\D", "", str(meta.get("Hesap", "")).split("-")[-1])[-4:]
    card = f"Garanti account •{last4}"
    start, end = acc_date(meta["Başlangıç"]), acc_date(meta["Bitiş"])
    data.reverse()  # the export is newest-first
    breaks = sum(1 for a, b in zip(data, data[1:])
                 if abs(acc_num(a["Bakiye"]) + acc_num(b["Tutar"]) - acc_num(b["Bakiye"])) > 0.005)
    if data and abs(acc_num(data[-1]["Bakiye"]) - acc_num(meta.get("Bakiye", data[-1]["Bakiye"]))) > 0.005:
        breaks += 1
    sid = f"GARANTI-ACCT-{last4}-{start}-{end}"
    rows = []
    for x in data:
        amt = acc_num(x["Tutar"])
        desc = str(x["Açıklama"]).strip()
        ext = f"{x.get('Dekont No', '')}|{amt:.2f}|{fold(desc)}"
        if ext in known_ext:
            continue
        typ = classify_account_row(desc, str(x.get("Etiket", "")), amt, holder)
        name, memo = (counterparty(desc) if typ in ("Transfer out", "Reimbursement", "Own transfer", "Family support",
                                                    "Other income") else ("", ""))
        d = acc_date(x["Tarih"])
        rows.append(dict(statement_id=sid, bank="Garanti BBVA", card=card, date=d, description=desc,
                         bank_category=str(x.get("Etiket", "")), section="Account", type=typ, currency="TRY",
                         amount=-amt, orig_amount="", orig_currency="", installment="", bonus="",
                         payment_date=d, source_file=path.name, ext_id=ext, counterparty=name, memo=memo,
                         _bal=acc_num(x["Bakiye"]), _raw=amt))
    st = account_statement(sid, "Garanti BBVA", card, f"Hesap {meta.get('Hesap', '')}", start, end, rows, path.name,
                           acc_num(data[-1]["Bakiye"]) if data else 0.0)
    return st, rows, breaks


def account_statement(sid, bank, card, account_no, start, end, rows, fname, closing=0.0):
    """Pseudo-statement for an account export: opening/closing balance of the newly imported rows. Recorded even when
    every row was already imported (overlapping exports), so its date range still counts as covered."""
    return dict(statement_id=sid, bank=bank, card=card, currency="TRY", ekstre_no=account_no, period_start=start,
                cutoff=end, payment_date=end, prev_balance=rows[0]["_bal"] - rows[0]["_raw"] if rows else closing,
                spending=sum(-r["_raw"] for r in rows if r["_raw"] < 0),
                fees_interest=sum(r["amount"] for r in rows if r["type"] == "Fee"),
                payments_credits=sum(r["_raw"] for r in rows if r["_raw"] > 0), balance_due=rows[-1]["_bal"] if rows else closing,
                min_payment="", source_file=fname, kind="Account")


def qnb_counterparty(desc: str) -> str:
    """'Transfer İşlemleri - Alıcı : Teoman Arabul Türkiye Garanti Bankası A.Ş. Sorgu No: …' -> 'Teoman Arabul'."""
    m = re.search(r"(?:Alıcı|Gönderen)\s*:\s*(.+)", desc)
    if not m:
        return ""
    return re.split(r"\s+(?:Sorgu No|Türkiye|T\.C\.|\S*[Bb]ank\S*|\S*BANK\S*)", m.group(1))[0].strip()


def parse_qnb_account(path: Path, known_ext: set):
    """QNB 'Hesap Hareketleri' PDF. Returns only rows not imported before."""
    with pdfplumber.open(path) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        table = [row for p in pdf.pages for t in p.extract_tables() for row in t]
    m = re.search(r"Iban\s*:\s*(TR[\d ]+)", text)
    iban = re.sub(r"\s", "", m.group(1)) if m else ""
    rng = re.search(r"Tarih Aral\S*\s*:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})", text)
    if iban[4:9] != "00111" or not rng:  # 00111 = QNB's bank code in the IBAN
        return None
    m = re.search(r"Ad Soyad\s*:\s*(.+?)\s*$", text, re.M)
    holder = m.group(1) if m else ""
    start, end = acc_date(rng.group(1)), acc_date(rng.group(2))
    last4 = iban[-4:]
    card = f"QNB account •{last4}"
    sid = f"QNB-ACCT-{last4}-{start}-{end}"
    opening, data = None, []
    for row in table:
        cells = [(c or "").replace("\n", " ").strip() for c in row]
        if len(cells) < 5:
            continue
        if "DEVREDEN" in fold(cells[2]):
            opening = en_num(cells[4])
        elif re.fullmatch(r"\d{2}/\d{2}/\d{4}", cells[0]):
            data.append(cells)
    breaks, bal = 0, opening
    for c in data:
        if bal is not None and abs(bal + en_num(c[3]) - en_num(c[4])) > 0.005:
            breaks += 1
        bal = en_num(c[4])
    rows = []
    for c in data:
        d, desc, amt = acc_date(c[0]), c[2], en_num(c[3])
        ext = f"{d}|{amt:.2f}|{en_num(c[4]):.2f}|{fold(desc)}"
        if ext in known_ext:
            continue
        name = qnb_counterparty(desc)
        typ = classify_account_row(desc, c[1], amt, holder, name)
        rows.append(dict(statement_id=sid, bank="QNB", card=card, date=d, description=desc, bank_category=c[1],
                         section="Account", type=typ, currency="TRY", amount=-amt, orig_amount="",
                         orig_currency="", installment="", bonus="", payment_date=d, source_file=path.name,
                         ext_id=ext, memo="",
                         counterparty=name if typ in ("Transfer out", "Reimbursement", "Own transfer") else "",
                         _bal=en_num(c[4]), _raw=amt))
    return account_statement(sid, "QNB", card, f"IBAN …{iban[-8:]}", start, end, rows, path.name,
                             bal or 0.0), rows, breaks


def ziraat_counterparty(desc: str, holder: str):
    """'Gönd: YASİN MÜMTAZ KAŞKAR lahmacun 0046-Akbank T.A.Ş. FAST işlemi' -> ('YASİN MÜMTAZ KAŞKAR', 'lahmacun');
    'Akbank T.A.Ş./TR95…-Teoman Arabul/FAST işlemi' -> ('Teoman Arabul', '')."""
    m = re.search(r"Gönd:\s*(.+?)\s+\d{4}-", desc)
    if m:
        words = m.group(1).split()
        caps = []
        for w in words:  # the name is the leading run of ALL-CAPS words; what follows is the sender's note
            if w.upper() != w or not re.search(r"[A-ZÇĞİÖŞÜ]", w):
                break
            caps.append(w)
        half = len(caps) // 2
        if half and caps[:half] == caps[half:2 * half]:  # 'ARDA BARIŞ TONBİL ARDA BARIŞ TONBİL tarafından…'
            caps = caps[:half]
        surname = holder.split()[-1].upper() if holder else ""
        if surname in caps:  # 'EMEL ARABUL ASKERLIK PARASI' → stop at the family name
            caps = caps[:caps.index(surname) + 1]
        name = " ".join(caps)
        return name, m.group(1)[len(name):].strip()
    m = re.search(r"/TR\d+-(.+?)/(?:FAST|EFT|HAVALE)", desc)
    return (m.group(1).strip() if m else ""), ""


def parse_ziraat_account(path: Path, known_ext: set):
    """Ziraat Bankası 'Hesap hareketleri' Excel export (its styles break openpyxl, so it's read with calamine)."""
    try:
        import pandas as pd
        df = pd.read_excel(path, header=None, engine="calamine")
    except Exception:
        return None
    grid = df.astype(object).where(df.notna(), "").values.tolist()
    meta = {str(r[0]).strip(): next((str(c).strip() for c in r[1:] if str(c).strip()), "") for r in grid if r[0] != ""}
    iban = re.sub(r"\s", "", meta.get("IBAN", ""))
    text = "\n".join(" ".join(str(c) for c in r if c != "") for r in grid)
    rng = re.search(r"(\d{2}\.\d{2}\.\d{4}) - (\d{2}\.\d{2}\.\d{4}) tarihleri", text)
    if iban[4:9] != "00010" or not rng:  # 00010 = Ziraat's bank code in the IBAN
        return None
    m = re.search(r"Sayın\s+(.+)", text)
    holder = m.group(1).strip() if m else ""
    start, end = (dt.datetime.strptime(x, "%d.%m.%Y").date() for x in rng.groups())
    last4 = iban[-4:]
    card = f"Ziraat account •{last4}"
    sid = f"ZIRAAT-ACCT-{last4}-{start}-{end}"
    data = [r for r in grid if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", str(r[0]).strip())]
    data.reverse()  # newest-first in the export
    breaks = sum(1 for a, b in zip(data, data[1:]) if abs(acc_num(a[4]) + acc_num(b[3]) - acc_num(b[4])) > 0.005)
    rows = []
    for x in data:
        d = dt.datetime.strptime(str(x[0]).strip(), "%d.%m.%Y").date()
        desc, amt = " ".join(str(x[2]).split()), acc_num(x[3])
        ext = f"{d}|{x[1]}|{amt:.2f}|{fold(desc)}"
        if ext in known_ext:
            continue
        name, memo = ziraat_counterparty(desc, holder)
        typ = classify_account_row(desc, "", amt, holder, name)
        keep = typ in ("Transfer out", "Reimbursement", "Own transfer", "Family support", "Other income")
        rows.append(dict(statement_id=sid, bank="Ziraat Bankası", card=card, date=d, description=desc,
                         bank_category="", section="Account", type=typ, currency="TRY", amount=-amt, orig_amount="",
                         orig_currency="", installment="", bonus="", payment_date=d, source_file=path.name,
                         ext_id=ext, counterparty=name if keep else "", memo=memo if keep else "",
                         _bal=acc_num(x[4]), _raw=amt))
    return account_statement(sid, "Ziraat Bankası", card, meta.get("Hesap Numarası", ""), start, end, rows,
                             path.name, acc_num(data[-1][4]) if data else 0.0), rows, breaks


def parse_account_xlsx(path: Path, known_ext: set):
    """Excel account exports: Garanti first, then Ziraat."""
    try:
        got = parse_garanti_account(path, known_ext)
        if got:
            return got
    except Exception:
        pass  # e.g. Ziraat's export, whose styles openpyxl can't read
    return parse_ziraat_account(path, known_ext)


# ───────────────────────────── FX (TCMB) ─────────────────────────────
def tcmb_rate(cur: str, day: dt.date):
    """TCMB forex selling rate for `cur`, stepping back over weekends/holidays."""
    for back in range(0, 10):
        d = day - dt.timedelta(days=back)
        url = f"https://www.tcmb.gov.tr/kurlar/{d:%Y%m}/{d:%d%m%Y}.xml"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as r:
                root = ET.fromstring(r.read())
        except Exception:
            continue
        for c in root.findall("Currency"):
            if c.get("CurrencyCode") == cur and c.findtext("ForexSelling"):
                return float(c.findtext("ForexSelling")), d, url
    return None


def ensure_fx(statements, fx_rows):
    today = dt.date.today()
    by_key = {r["key"]: r for r in fx_rows}
    for s in statements:
        if s["currency"] == "TRY":
            continue
        pay = to_date(s["payment_date"])
        key = f"{s['currency']}|{pay}"
        row = by_key.get(key)
        if row and row["status"] in ("final", "manual", "derived"):
            continue
        got = tcmb_rate(s["currency"], min(pay, today))
        status = "final" if pay <= today else "provisional"
        if got:
            rate, rdate, url = got
            row = dict(key=key, currency=s["currency"], payment_date=iso(pay), rate=f"{rate:.4f}",
                       rate_date=iso(rdate), source=f"TCMB ForexSelling {url}", status=status)
            print(f"  FX {key}: {rate:.4f} ({status}, TCMB {rdate})")
        elif not row:
            row = dict(key=key, currency=s["currency"], payment_date=iso(pay), rate="", rate_date="",
                       source="", status="MISSING – enter rate")
            print(f"  ! FX {key}: could not fetch rate — fill it in data/fx_rates.csv")
        by_key[key] = row
    return sorted(by_key.values(), key=lambda r: r["key"])


def derive_fx_from_transfers(statements, txs, fx_rows):
    """When an unpaid EUR/USD statement is moved into TL debt, the next statements show both sides of the bank's
    own conversion: 'TLye dönen EUR Borç Ödeme 230,21' (EUR statement) and 'EUR ekstre devir bakiyesi 13.114,49'
    (TL statement, same cutoff). TL ÷ EUR is the rate that EUR statement was actually paid at."""
    by_key = {r["key"]: r for r in fx_rows}
    st_by_id = {s["statement_id"]: s for s in statements}
    fx_side, tl_side = defaultdict(float), defaultdict(float)
    for t in txs:
        if t["type"] != "FX balance transfer":
            continue
        s = st_by_id[t["statement_id"]]
        k = (s["card"], iso(to_date(s["cutoff"])))
        if s["currency"] != "TRY":
            fx_side[k + (s["currency"],)] -= float(t["amount"])
        else:
            m = re.search(r"\b(EUR|USD|GBP)\b", fold(t["description"]))
            if m:
                tl_side[k + (m.group(1),)] += float(t["amount"])
    for (card, cutoff, cur), fx_amt in fx_side.items():
        tl_amt = tl_side.get((card, cutoff, cur))
        # the statement being paid off: same card & currency, the latest cutoff before this one
        prev = max((s for s in statements if s["card"] == card and s["currency"] == cur
                    and iso(to_date(s["cutoff"])) < cutoff), key=lambda s: iso(to_date(s["cutoff"])), default=None)
        if not tl_amt or fx_amt <= 0 or not prev:
            continue
        key = f"{cur}|{iso(to_date(prev['payment_date']))}"
        if by_key.get(key, {}).get("status") == "manual":
            continue
        rate = tl_amt / fx_amt
        by_key[key] = dict(key=key, currency=cur, payment_date=iso(to_date(prev["payment_date"])), rate=f"{rate:.4f}",
                           rate_date=cutoff, status="derived",
                           source=(f"Bank's own conversion: {fx_amt:,.2f} {cur} moved into the TL statement of "
                                   f"{cutoff} as {tl_amt:,.2f} TRY"))
        print(f"  FX {key}: {rate:.4f} (the bank's actual conversion)")
    return sorted(by_key.values(), key=lambda r: r["key"])


# ───────────────────────────── categorisation ─────────────────────────────
def load_rules():
    rules = []
    for r in read_csv(RULES_CSV):
        r["rx"] = re.compile(r["pattern"], re.I)
        rules.append(r)
    return rules


def rule_categories(rules):
    cats = defaultdict(set)
    for r in rules:
        cats[r["category"]].add(r["subcategory"])
    return cats


def load_people():
    people = []
    for r in read_csv(PEOPLE_CSV):
        if r.get("pattern"):
            r["rx"] = re.compile(r["pattern"], re.I)
            people.append(r)
    return people


_BANK_CATEGORY_SUGGESTIONS = [
    (r"MARKET|SUPERMARKET|GIDA", "Groceries", "Supermarket"),
    (r"RESTORAN|YEMEK|CAFE|KAFE", "Food & Dining", "Restaurant"),
    (r"AKARYAKIT|BENZIN|PETROL", "Transport", "Fuel"),
    (r"ULASIM|TAKSI|OTOPARK", "Transport", "Transport"),
    (r"GIYIM|AYAKKABI", "Shopping", "Clothes"),
    (r"ELEKTRONIK", "Shopping", "Electronics"),
    (r"KOZMETIK|KISISEL BAKIM", "Shopping", "Personal care"),
    (r"SAGLIK|ECZANE|HASTANE", "Health", "Health"),
    (r"SEYAHAT|OTEL|HAVAYOL", "Travel", "Travel"),
    (r"FATURA|TELEKOM|ELEKTRIK|DOGALGAZ", "Bills & Utilities", "Utilities"),
    (r"ABONELIK|DIJITAL", "Subscriptions & Digital", "Digital"),
    (r"EGLENCE|SINEMA|TIYATRO", "Entertainment", "Entertainment"),
    (r"EGITIM|KITAP|KIRTASIYE", "Education", "Education"),
]


def bank_category_suggestion(raw: str):
    """Translate a bank's broad Turkish category into our categories, as a reviewable suggestion."""
    value = fold(raw or "")
    return next(((cat, sub) for pattern, cat, sub in _BANK_CATEGORY_SUGGESTIONS
                 if re.search(pattern, value)), None)


def categorize(tx, rules, people=()):
    typ, fd = tx["type"], fold(tx["description"])
    if typ == "Payment":
        res = ["Card payment", "Transfers & Payments", "Card payment", "high", ""]
    elif typ == "Cash advance":
        res = [f"Cash advance ({tx['bank']})", "Cash Advance", "Installment cash advance", "high",
               "Borrowed money, not counted as spending"]
    elif typ == "Cash advance received":
        res = [f"Cash advance ({tx['bank']})", "Cash Advance", "Received in account", "high",
               "Principal of the cash advance"]
    elif typ == "FX balance transfer":
        res = ["EUR/USD balance moved to TL", "Transfers & Payments", "FX balance transfer", "high",
               "Unpaid foreign-currency balance converted into TL debt; the purchases are counted in the EUR/USD statement"]
    elif typ == "Bonus used":
        res = ["Bonus points spent", "Transfers & Payments", "Bonus redemption", "high", "Paid with Bonus points, not TL"]
    elif typ == "Student loan":
        res = ["KYK student loan", "Income", "Student loan (KYK)", "high", "A loan to be repaid – not income"]
    elif typ == "Family support":
        res = [tr_title(tx.get("counterparty") or "Family"), "Income", "Family support", "high", ""]
    elif typ == "One-off":
        res = ["Bedelli askerlik (military service fee)" if "BEDELLI" in fd else tx["description"][:40],
               "Government & Taxes", "One-off", "high", "Kept out of regular spending"]
    elif typ == "Income":
        yol = "YOL YARDIMI" in fd
        res = ["Travel allowance (yol yardımı)" if yol else "Salary", "Income", "Allowance" if yol else "Salary",
               "high", ""]
    elif typ == "Other income":
        res = [tr_title(tx.get("counterparty") or "") or tx["description"], "Income", "Other", "guess", ""]
    elif typ == "Investment":
        buy = float(tx["amount"]) > 0
        if "BINANCE" in fd or "KRIPTO" in fd:
            res = ["Binance (crypto)", "Investments", "Crypto buy" if buy else "Crypto sell", "high", "Not spending"]
        elif "ALTIN" in fd:
            res = ["Gold account", "Investments", "Gold buy" if buy else "Gold sell", "high", "Not spending"]
        else:
            res = ["GPY 1. Para Piyasa fund", "Investments", "Fund buy" if buy else "Fund sell", "high",
                   "Not spending"]
    elif typ == "Investment cost":
        res = ["MKK fee" if "MKK" in fd else "Fund withholding tax (stopaj)", "Investments", "Tax & fees", "high", ""]
    elif typ == "Currency exchange":
        res = ["USD / EUR / gold (Garanti)", "Investments",
               "FX & gold buy" if float(tx["amount"]) > 0 else "FX & gold sell", "high", "Not spending"]
    elif typ == "Own transfer":
        res = ["Own account transfer", "Transfers & Payments", "Own accounts", "high", "Between your own accounts"]
    elif typ == "Cash withdrawal":
        res = ["ATM cash", "Cash", "ATM withdrawal", "high", "Cash spent somewhere – counted as spending"]
    elif typ in ("Transfer out", "Reimbursement", "Returned"):
        person = tx.get("counterparty") or "Unnamed transfer"
        p = next((r for r in people if r["rx"].search(fold(person))), None)
        if p:
            res = [p.get("label") or tr_title(person), p["category"], p["subcategory"], "manual", p.get("note", "")]
        else:
            sub = {"Transfer out": "Paid to friends", "Reimbursement": "Friends paid back",
                   "Returned": "Returned / cancelled"}[typ]
            res = [tr_title(person), "People (IBAN)", sub, "high", ""]
    elif typ == "Fee" and tx.get("section") == "Account":
        bank = tx["bank"].split()[0]
        res = [f"{bank} – FX fee" if "DOVIZ" in fd else f"{bank} – transfer fees", "Bank Fees & Interest",
               "Transfer fees", "high", ""]
    else:
        fd = fold(tx["description"])
        rule = next((r for r in rules if r["rx"].search(fd)), None)
        if rule:
            res = [rule["merchant"], rule["category"], rule["subcategory"], rule["confidence"], rule["note"]]
        else:
            suggestion = bank_category_suggestion(tx.get("bank_category", ""))
            if suggestion:
                cat, sub = suggestion
                res = [tx["description"].title(), cat, sub, "guess",
                       f"Suggested by the bank: {tx['bank_category']}"]
            else:
                res = [tx["description"].title(), "Other", "Uncategorized", "none", "No rule matched"]
    if tx.get("merchant_override"):
        res[0] = tx["merchant_override"]
    if tx.get("category_override"):
        res[1], res[3] = tx["category_override"], "manual"
    if tx.get("subcategory_override"):
        res[2] = tx["subcategory_override"]
    return res


# ───────────────────────────── import ─────────────────────────────
def _kind_of(path: Path):
    """'card' (kredi kartı ekstresi), 'account' (hesap dökümü) or None."""
    ext = path.suffix.lower()
    if ext == ".xlsx":
        return "account"
    if ext != ".pdf":
        return None
    try:
        with pdfplumber.open(path) as pdf:
            head = fold("\n".join((p.extract_text() or "") for p in pdf.pages[:2]))
    except Exception:
        return None
    if "HESAP KESIM" in head:
        return "card"
    if "IBAN" in head and ("TARIH ARAL" in head or "HESAP HAREKET" in head):
        return "account"
    return None


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def scan_documents():
    """Every .pdf / .xlsx in Belgeler/ (subfolders too), plus the old ekstreler/ hesap_dokumu/ inbox/ folders."""
    seen, out = set(), []
    for d in (DOCS_DIR, STATEMENT_DIR, ACCOUNT_DIR, INBOX_DIR):
        if not d.exists():
            continue
        for f in sorted(d.rglob("*")):
            if (f.is_file() and f.suffix.lower() in (".pdf", ".xlsx") and not f.name.startswith("~$")
                    and f.resolve() not in seen):
                seen.add(f.resolve())
                out.append(f)
    return out


def import_new(txs, sts, fresh=False):
    """Parse every document in the folder that hasn't been handled before. Files are recognised by their content
    (SHA-256), so renaming, moving into subfolders or re-using a bank's default file name is all fine; a file whose
    statement / rows are already in the data (a re-download, an overlapping export) is recorded and skipped."""
    register = [] if fresh else read_csv(FILES_CSV)
    docs = scan_documents()
    if not register and not fresh:  # first run with the register: files imported earlier count as handled
        done = {s["source_file"] for s in sts}
        register = [dict(sha256=_sha256(f), file=f.name, kind="", result="imported earlier", date=iso(dt.date.today()))
                    for f in docs if f.name in done]
    known_hash = {r["sha256"] for r in register}

    def rel(f):
        try:
            return f.relative_to(ROOT).as_posix()
        except ValueError:
            return str(f)

    def note(f, h, kind, result):
        register.append(dict(sha256=h, file=rel(f), kind=kind or "", result=result, date=iso(dt.date.today())))

    todo = []
    for f in docs:
        h = _sha256(f)
        if h in known_hash:
            continue
        known_hash.add(h)  # the same file twice in the folder: handle it once
        todo.append((f, h, _kind_of(f)))
    if not todo:
        print("  nothing new in the folder")
    for f, h, kind in todo:
        if kind is None:
            print(f"  ! {rel(f)}: not a card statement or account export I can read — ignored")
            note(f, h, kind, "not recognised")

    known_ids = {s["statement_id"] for s in sts}
    added = 0
    for f, h, kind in [x for x in todo if x[2] == "card"]:
        try:
            parsed = parse_pdf(f)
        except Exception as e:
            print(f"  ! {rel(f)}: couldn't read it ({type(e).__name__}) — the bank may have changed its layout")
            note(f, h, kind, f"error: {type(e).__name__}")
            continue
        if not parsed:
            note(f, h, kind, "not recognised")
            continue
        st, rows = parsed
        if st["statement_id"] in known_ids:
            print(f"  = {rel(f)}: statement {st['statement_id']} is already imported — skipped")
            note(f, h, kind, f"duplicate of {st['statement_id']}")
            continue
        total = round(st["prev_balance"] + sum(r["amount"] for r in rows), 2)
        ok = abs(total - st["balance_due"]) < 0.01
        status = "OK" if ok else f"MISMATCH ({total:.2f} vs {st['balance_due']:.2f})"
        print(f"  + {rel(f)}: {len(rows)} rows, {st['statement_id']} — reconciliation {status}")
        for i, r in enumerate(rows, 1):
            r["id"] = f"{st['statement_id']}-{i:03d}"
            for k in ("amount", "orig_amount", "bonus"):
                r[k] = money(r[k])
            txs.append(r)
        for k in ("prev_balance", "spending", "fees_interest", "payments_credits", "balance_due", "min_payment"):
            st[k] = money(st[k])
        sts.append(st)
        known_ids.add(st["statement_id"])
        note(f, h, kind, f"imported {st['statement_id']} ({len(rows)} rows, reconciliation {status})")
        added += 1
    known_ext = {t.get("ext_id") for t in txs if t.get("ext_id")}
    for f, h, kind in [x for x in todo if x[2] == "account"]:
        parser = {".xlsx": parse_account_xlsx, ".pdf": parse_qnb_account}[f.suffix.lower()]
        try:
            parsed = parser(f, known_ext)
        except Exception as e:
            print(f"  ! {rel(f)}: couldn't read it ({type(e).__name__})")
            note(f, h, kind, f"error: {type(e).__name__}")
            continue
        if not parsed:
            print(f"  ! {rel(f)}: not a Garanti / QNB / Ziraat account export I can read")
            note(f, h, kind, "not recognised")
            continue
        st, rows, breaks = parsed
        if st["statement_id"] in known_ids:
            if not rows:  # the very same export again (e.g. downloaded twice)
                print(f"  = {rel(f)}: same date range as an imported export and nothing new — skipped")
                note(f, h, kind, "nothing new")
                continue
            base, n = st["statement_id"], 2  # same range downloaded again later, with new rows: keep ids unique
            while f"{base}-{n}" in known_ids:
                n += 1
            st["statement_id"] = f"{base}-{n}"
            for r in rows:
                r["statement_id"] = st["statement_id"]
        if rows:
            print(f"  + {rel(f)}: {len(rows)} new rows, {st['statement_id']} — running balance "
                  f"{'OK' if not breaks else f'{breaks} break(s)'}")
        else:
            print(f"  = {rel(f)}: every row was already imported (overlap) — coverage recorded up to "
                  f"{st['cutoff']:%d.%m.%Y}")
        for i, r in enumerate(rows, 1):
            r["id"] = f"{st['statement_id']}-{i:03d}"
            r["amount"] = money(r["amount"])
            txs.append(r)
            known_ext.add(r["ext_id"])
        for k in ("prev_balance", "spending", "fees_interest", "payments_credits", "balance_due"):
            st[k] = money(st[k])
        sts.append(st)
        known_ids.add(st["statement_id"])
        note(f, h, kind, f"imported {len(rows)} new rows ({st['statement_id']})" if rows
             else "nothing new (all rows already imported)")
        added += 1
    write_csv(FILES_CSV, FILES_FIELDS, register)
    return added


# ───────────────────────────── workbook ─────────────────────────────
FONT = "Arial"
ACCENT = "2E6FD8"
HEAD_FILL = PatternFill("solid", fgColor="1F3A5F")
SUB_FILL = PatternFill("solid", fgColor="E8EEF7")
REVIEW_FILL = PatternFill("solid", fgColor="FFF4CC")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
MONEY = '#,##0.00;[Red]-#,##0.00;"-"'
MONEY0 = '#,##0;[Red]-#,##0;"-"'
PCT = '0.0%;-0.0%;"-"'
DATE = "dd.mm.yyyy"

TXCOLS = ["ID", "Date", "Month", "Weekday", "Card", "Merchant", "Category", "Subcategory", "Type", "Currency",
          "Amount", "FX rate", "Amount TRY", "Counts as spending", "Spending TRY", "Channel", "Tag",
          "Statement ID", "FX key", "Payment date", "Float days", "Bank category", "Description (raw)",
          "Installment", "Bonus", "Review", "Rotation", "Note", "Source file", "Complete day", "Normal day",
          "Spend month", "Adv principal"]
NORMAL_DAYS = "COUNTIFS(Daily!$D:$D,1,Daily!$I:$I,0)"   # complete days that aren't trip days
TC = {name: get_column_letter(i) for i, name in enumerate(TXCOLS, 1)}


def tx(name):
    return f"Transactions!${TC[name]}:${TC[name]}"


def style_header(ws, row, ncols, start=1):
    for c in range(start, start + ncols):
        cell = ws.cell(row, c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def put(ws, row, col, value, fmt=None, bold=False, color=None, fill=None, italic=False, size=None):
    c = ws.cell(row, col, value)
    c.font = Font(name=FONT, bold=bold, color=color, italic=italic, size=size or 10)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def load_budgets(cats):
    """data/budgets.csv (category, monthly_budget). New categories are appended with an empty budget."""
    have = {r["category"]: r.get("monthly_budget") or "" for r in read_csv(BUDGETS_CSV)}
    if any(c not in have for c in cats):
        order = cats + [c for c in have if c not in cats]
        write_csv(BUDGETS_CSV, ["category", "monthly_budget"],
                  [{"category": c, "monthly_budget": have.get(c, "")} for c in order])
    return {c: float(v) for c, v in have.items() if v}


def build_history_sheet(wb, rows, sts):
    """Month-by-month spending estimated from bank-account data alone, so months and years without card statements
    can be compared too. Card payments are moved back to the month their purchases were made."""
    pays = [r for r in rows if r.get("spend_month")]
    acct = [r for r in rows if r["t"]["section"] == "Account"]
    if not acct:
        return
    first = min([r["spend_month"] for r in pays] + [r["month"] for r in acct])
    last = max(r["month"] for r in rows)
    months, d = [], dt.date(int(first[:4]), int(first[5:7]), 1)
    while d.strftime("%Y-%m") <= last:
        months.append(d.strftime("%Y-%m"))
        d = add_months(d, 1)
    # a month is final once every card that's still being paid has had its payment for that month
    paid_months, last_pay = defaultdict(set), {}
    for r in pays:
        paid_months[r["card_no"]].add(r["spend_month"])
        last_pay[r["card_no"]] = max(last_pay.get(r["card_no"], r["date"]), r["date"])
    active = [c for c in paid_months if (dt.date.today() - last_pay[c]).days <= 45]
    settled = min((max(paid_months[c]) for c in active), default=last)
    acct_first = {}
    for r in acct:
        acct_first[r["t"]["card"]] = min(acct_first.get(r["t"]["card"], r["date"]), r["date"])
    known_cards = {s["card"][-4:] for s in sts if (s.get("kind") or "Card") != "Account"}

    ws = wb.create_sheet("History", 2)
    ws.sheet_properties.tabColor = ACCENT
    put(ws, 1, 1, "History — month by month from your bank accounts", bold=True, size=18, color="1F3A5F")
    put(ws, 2, 1, "Estimate: every card payment is moved back to the month its purchases were made (the statement "
                  "cycle before the payment), minus the cash-advance principal inside it, plus what left the accounts "
                  "directly (money to friends minus paybacks, ATM cash, transfer fees).", italic=True, color="555555")
    starts = ", ".join(f"{c.split(' •')[0].replace(' account', '')} from {d0:%d.%m.%Y}"
                       for c, d0 in sorted(acct_first.items(), key=lambda kv: kv[1]))
    put(ws, 3, 1, f"Kept apart: investments, own transfers, KYK loan, family money, one-offs. Account data: {starts} — "
                  f"months before an account starts can't show what it paid for, so they understate spending (and "
                  f"accounts you haven't sent never show). Grey months are still waiting for a card payment.",
        italic=True, color="555555")
    hdr = ["Month", "Card payments TRY", "Cash-advance principal in them", "Card spending (est.)",
           "Via accounts (friends, cash, fees)", "Estimated spending TRY", "3-month average", "Salary & allowances",
           "KYK student loan", "Family support", "Left over (salary − spending)", "Into investments (net)",
           "One-offs", "Status", "Cards paid", "Accounts visible", "Year"]
    for j, h in enumerate(hdr, 1):
        ws.cell(5, j, h)
    style_header(ws, 5, len(hdr))
    ws.row_dimensions[5].height = 44
    h0, h1 = 6, 6 + len(months) - 1
    acct_crit = f'{tx("Card")},"*account*"'
    for k, mo in enumerate(months):
        r = h0 + k
        month_end = add_months(dt.date(int(mo[:4]), int(mo[5:]), 1), 1) - dt.timedelta(1)
        cards_paid = sorted({p["card_no"] for p in pays if p["spend_month"] == mo})
        visible = [c.split(" •")[0].replace(" account", "")
                   for c, f0 in sorted(acct_first.items(), key=lambda kv: kv[1]) if f0 <= month_end]
        put(ws, r, 1, mo, bold=True)
        put(ws, r, 2, f'=SUMIFS({tx("Amount TRY")},{tx("Type")},"Payment",{acct_crit},{tx("Spend month")},A{r})',
            MONEY0)
        put(ws, r, 3, f'=SUMIFS({tx("Adv principal")},{tx("Spend month")},A{r})', MONEY0)
        put(ws, r, 4, f"=B{r}-C{r}", MONEY0)
        put(ws, r, 5, f'=SUMIFS({tx("Spending TRY")},{acct_crit},{tx("Month")},A{r})', MONEY0)
        put(ws, r, 6, f"=D{r}+E{r}", MONEY0, bold=True)
        put(ws, r, 7, f"=AVERAGE(F{max(h0, r - 2)}:F{r})", MONEY0)
        for col, typ in ((8, "Income"), (9, "Student loan"), (10, "Family support")):
            put(ws, r, col, f'=-SUMIFS({tx("Amount TRY")},{tx("Type")},"{typ}",{tx("Month")},A{r})', MONEY0)
        put(ws, r, 11, f"=H{r}-F{r}", MONEY0)
        put(ws, r, 12, f'=SUMIFS({tx("Amount TRY")},{tx("Type")},"Investment",{tx("Month")},A{r})'
                       f'+SUMIFS({tx("Amount TRY")},{tx("Type")},"Currency exchange",{tx("Month")},A{r})', MONEY0)
        put(ws, r, 13, f'=SUMIFS({tx("Amount TRY")},{tx("Type")},"One-off",{tx("Month")},A{r})', MONEY0)
        put(ws, r, 14, "final" if mo <= settled else "payments pending")
        put(ws, r, 15, " ".join(f"•{c}" for c in cards_paid))
        put(ws, r, 16, ", ".join(visible))
        put(ws, r, 17, mo[:4])
    ws.conditional_formatting.add(f"A{h0}:P{h1}", FormulaRule(formula=[f'$N{h0}="payments pending"'],
                                                              font=Font(italic=True, color="999999")))

    # per year
    y0 = h1 + 3
    put(ws, y0 - 1, 1, "Per year (final months only)", bold=True, size=12, color="1F3A5F")
    yh = ["Year", "Final months", "Avg spending / month", "Avg salary / month", "Avg left over / month",
          "KYK received", "Into investments (net)", "One-offs"]
    for j, h in enumerate(yh, 1):
        ws.cell(y0, j, h)
    style_header(ws, y0, len(yh))

    def col(c):
        return f"${c}${h0}:${c}${h1}"

    years = sorted({m[:4] for m in months})
    for k, y in enumerate(years):
        r = y0 + 1 + k
        put(ws, r, 1, y, bold=True)
        put(ws, r, 2, f'=COUNTIFS({col("Q")},A{r},{col("N")},"final")')
        put(ws, r, 3, f'=IFERROR(AVERAGEIFS({col("F")},{col("Q")},A{r},{col("N")},"final"),0)', MONEY0, bold=True)
        put(ws, r, 4, f'=IFERROR(AVERAGEIFS({col("H")},{col("Q")},A{r},{col("H")},">0"),0)', MONEY0)
        put(ws, r, 5, f'=IFERROR(AVERAGEIFS({col("K")},{col("Q")},A{r},{col("N")},"final",{col("H")},">0"),0)',
            MONEY0)
        put(ws, r, 6, f'=SUMIFS({col("I")},{col("Q")},A{r})', MONEY0)
        put(ws, r, 7, f'=SUMIFS({col("L")},{col("Q")},A{r})', MONEY0)
        put(ws, r, 8, f'=SUMIFS({col("M")},{col("Q")},A{r})', MONEY0)

    # cards paid from these accounts
    c0 = y0 + len(years) + 3
    put(ws, c0 - 1, 1, "Cards paid from these accounts", bold=True, size=12, color="1F3A5F")
    ch_ = ["Card ending", "Statements imported?", "Payments", "Total paid TRY", "Last payment"]
    for j, h in enumerate(ch_, 1):
        ws.cell(c0, j, h)
    style_header(ws, c0, len(ch_))
    for k, no in enumerate(sorted(paid_months, key=lambda c: last_pay[c], reverse=True)):
        r = c0 + 1 + k
        crit = f'{tx("Type")},"Payment",{acct_crit},{tx("Description (raw)")},"*{no}*"'
        put(ws, r, 1, f"•{no}", bold=True)
        put(ws, r, 2, "yes" if no in known_cards else "no – please send them",
            fill=None if no in known_cards else REVIEW_FILL)
        put(ws, r, 3, f"=COUNTIFS({crit})")
        put(ws, r, 4, f"=SUMIFS({tx('Amount TRY')},{crit})", MONEY0)
        put(ws, r, 5, f"=_xlfn.MAXIFS({tx('Date')},{crit})", DATE)

    widths(ws, {"A": 10, "B": 12, "C": 13, "D": 12, "E": 14, "F": 13, "G": 12, "H": 12, "I": 11, "J": 11, "K": 13,
                "L": 13, "M": 11, "N": 16, "O": 15, "P": 26, "Q": 6})
    ws.freeze_panes = "B6"


def confirmed_coverage(source_days, statements, confirmations):
    """Add explicit no-activity confirmations without changing imported statement coverage."""
    accounts = {s["card"] for s in statements if s.get("kind") == "Account"}
    coverage = {name: set(days) for name, days in source_days.items()}
    for r in confirmations:
        if r.get("source") not in accounts:
            raise ValueError("No-activity confirmation must name an imported bank account")
        start, end = to_date(r.get("start")), to_date(r.get("end"))
        if not start or not end or start > end:
            raise ValueError("Invalid no-activity confirmation date range")
        coverage[r["source"]].update(start + dt.timedelta(days=i) for i in range((end - start).days + 1))
    return coverage


def coverage_complete(day, known_days, first_days, excluded=()):
    active = [c for c in known_days if c not in excluded and first_days[c] <= day]
    return bool(active) and all(day in known_days[c] for c in active)


def build_workbook(txs, sts, fx_rows, rules, people=()):
    fx = {r["key"]: float(r["rate"]) for r in fx_rows if r["rate"]}
    tags = read_csv(TAGS_CSV)
    for s in sts:
        for k in ("cutoff", "payment_date", "period_start"):
            s[k] = to_date(s[k]) if isinstance(s[k], str) else s[k]

    # coverage windows per statement → per card per day
    starts = defaultdict(list)
    for s in sts:
        if s["period_start"]:
            starts[(s["card"], s["cutoff"])].append(s["period_start"])
    card_days = defaultdict(set)
    for s in sts:
        key = (s["card"], s["cutoff"])
        start = s["period_start"] or (min(starts[key]) if key in starts else add_months(s["cutoff"], -1) + dt.timedelta(1))
        s["window_start"] = start
        d = start
        while d <= s["cutoff"]:
            card_days[s["card"]].add(d)
            d += dt.timedelta(1)
    cards = sorted(card_days)
    confirmations = read_csv(NO_ACTIVITY_CSV)
    known_days = confirmed_coverage(card_days, sts, confirmations)
    exclusions = read_csv(EXCLUDED_SOURCES_CSV)
    excluded = {r["source"] for r in exclusions}
    accounts = {s["card"] for s in sts if s.get("kind") == "Account"}
    if not excluded <= accounts:
        raise ValueError("Coverage exclusions must name imported bank accounts")
    all_start = min([to_date(t["date"]) for t in txs] + [min(v) for v in card_days.values()])
    card_first = {c: all_start if c in IN_USE_BEFORE_FIRST_STATEMENT else min(card_days[c]) for c in cards}

    # enrich transactions
    # transfers that came straight back (same person, same amount, within 3 days, incl. FAST İPTAL) cancel out
    p2p = [t for t in txs if t["type"] in ("Transfer out", "Reimbursement")]
    returned = set()
    for a in p2p:
        if a["type"] != "Transfer out" or a["id"] in returned:
            continue
        for b in p2p:
            if (b["type"] == "Reimbursement" and b["id"] not in returned
                    and fold(b.get("counterparty") or "") == fold(a.get("counterparty") or "")
                    and abs(float(a["amount"]) + float(b["amount"])) < 0.01
                    and 0 <= (to_date(b["date"]) - to_date(a["date"])).days <= 3):
                returned |= {a["id"], b["id"]}
                break

    # unnamed outgoing transfers ('MOBIL-FAST-…') that arrive in another of your accounts the same day are yours
    own_in = [t for t in txs if t["type"] == "Own transfer" and float(t["amount"]) < 0]
    own_matched, used = set(), set()
    for a in txs:
        if a["type"] != "Transfer out" or (a.get("counterparty") or "").strip():
            continue
        b = next((b for b in own_in if b["id"] not in used and b["card"] != a["card"]
                  and abs(float(a["amount"]) + float(b["amount"])) < 0.01
                  and abs((to_date(b["date"]) - to_date(a["date"])).days) <= 1), None)
        if b:
            own_matched.add(a["id"])
            used.add(b["id"])

    rows = []
    for t in txs:
        if t["id"] in returned:
            t = dict(t, type="Returned")
        elif t["id"] in own_matched:
            t = dict(t, type="Own transfer", counterparty="(your other account)")
        d = to_date(t["date"]) if isinstance(t["date"], str) else t["date"]
        pay = to_date(t["payment_date"]) if isinstance(t["payment_date"], str) else t["payment_date"]
        merchant, cat, sub, conf, rnote = categorize(t, rules, people)
        cur = t["currency"]
        key = "" if cur == "TRY" else f"{cur}|{pay}"
        rate = 1.0 if cur == "TRY" else fx.get(key, 0.0)
        amt = float(t["amount"])
        channel = ("Abroad" if cur != "TRY" else "Online (foreign)" if t["section"] == "Foreign"
                   else "Bank account" if t["section"] == "Account" else "Domestic")
        tag = next((g["tag"] for g in tags
                    if (to_date(g["start"]) <= d <= to_date(g["end"])
                        and (not g["currency"] or g["currency"] == cur)
                        and (not g["card"] or g["card"] in t["card"]))
                    or (t["type"] == "Reimbursement" and g.get("reimbursements_until")
                        and to_date(g["start"]) <= d <= to_date(g["reimbursements_until"]))), "")
        counts = 1 if t["type"] in SPENDING_TYPES else 0
        month = d.strftime("%Y-%m")
        if t["type"] in ("Income", "Student loan"):  # "2026-Ağustos Maaş Ödemesi" paid on 4 Sep counts for August
            fd_ = fold(t["description"])
            mm = next((k + 1 for k, n in enumerate(TR_MONTHS) if n in fd_), None)
            yy = re.search(r"\b20\d\d\b", fd_)  # \b: KYK's reference '20232894591' isn't a year
            if mm and yy:
                month = f"{yy.group()}-{mm:02d}"
        rotation = ""
        if t["type"] == "Purchase" and t["section"] != "Account" and d >= ROTATION_START:
            expected = next((c for a, b, c in ROTATION if a <= d.day <= b), None)
            if expected and expected.upper() not in t["card"].upper():
                rotation = f"off-rotation (expected {expected})"
        rows.append(dict(t=t, date=d, pay=pay, merchant=merchant, cat=cat, sub=sub, conf=conf, rnote=rnote,
                         key=key, rate=rate, amt=amt, try_amt=round(amt * rate, 2), channel=channel, tag=tag,
                         counts=counts, rotation=rotation, month=month))
    rows.sort(key=lambda r: (r["date"], r["t"]["id"]))

    # ── card payments seen in accounts → the month their purchases happened (History sheet) ──
    # A payment settles the statement cut just before it; its purchases sit in the ~month before that cutoff.
    def card_no(desc):
        m = re.search(r"\*+\s*(\d{4})", desc) or re.match(r"^(\d{4})-\d{6,}", desc)
        return m.group(1) if m else ""

    card_pays = defaultdict(list)
    for r in rows:
        if r["t"]["section"] == "Account" and r["t"]["type"] == "Payment":
            r["card_no"] = card_no(r["t"]["description"])
            card_pays[r["card_no"]].append(r)
    known_cut = defaultdict(set)
    for s in sts:
        if (s.get("kind") or "Card") != "Account":
            known_cut[s["card"][-4:]].add(s["cutoff"])
    for no, lst in card_pays.items():
        # cutoff day: from imported statements, else ~10 days before the usual payment day
        cd = (Counter(c.day for c in known_cut[no]).most_common(1)[0][0] if known_cut[no]
              else max(1, min(28, int(statistics.median(r["date"].day for r in lst)) - 10)))
        cuts, have = set(known_cut[no]), {(c.year, c.month) for c in known_cut[no]}
        m = add_months(min(r["date"] for r in lst).replace(day=1), -2)
        while m <= max(r["date"] for r in lst):
            if (m.year, m.month) not in have:
                cuts.add(m.replace(day=min(cd, calendar.monthrange(m.year, m.month)[1])))  # a 31st cutoff
            m = add_months(m, 1)
        for r in lst:
            r["cycle"] = max((c for c in cuts if c < r["date"]), default=None)
            r["spend_month"] = (r["cycle"] - dt.timedelta(days=15)).strftime("%Y-%m") if r["cycle"] else r["month"]
    # cash-advance principal travels inside the next card payments (1/n per statement); it isn't spending
    adv_principal = defaultdict(float)
    for a in rows:
        if a["t"]["type"] != "Cash advance received":
            continue
        no, principal = card_no(a["t"]["description"]), -a["amt"]
        mm = re.search(r"MOBIL-(\d+)", a["t"]["description"])
        n = int(mm.group(1)) if mm else 1
        cycles = sorted({p["cycle"] for p in card_pays.get(no, []) if p["cycle"] and p["cycle"] >= a["date"]})
        for c in cycles[:n]:
            carrier = max((p for p in card_pays[no] if p["cycle"] == c), key=lambda p: p["amt"])
            adv_principal[carrier["t"]["id"]] += principal / n
    spend_rows = [r for r in rows if r["counts"]]

    first = min([r["date"] for r in rows] + [s["window_start"] for s in sts])
    last = max([r["date"] for r in rows] + [s["cutoff"] for s in sts])
    days = [first + dt.timedelta(i) for i in range((last - first).days + 1)]
    months = sorted({d.strftime("%Y-%m") for d in days})

    def complete(d):
        """Every active source is covered by imports or an explicit no-activity confirmation."""
        return coverage_complete(d, known_days, card_first, excluded)

    trip_ranges = [(to_date(g["start"]), to_date(g["end"])) for g in tags]

    wb = Workbook()
    if exclusions:
        audit = wb.create_sheet("Coverage exclusions")
        audit.append(["Account", "Reason"])
        for r in exclusions:
            audit.append([r["source"], r.get("note", "")])
        style_header(audit, 1, 2)
        widths(audit, {"A": 30, "B": 75})
    if confirmations:
        audit = wb.create_sheet("Confirmed inactivity")
        audit.append(["Account", "Start (inclusive)", "End (inclusive)", "Note"])
        for r in confirmations:
            audit.append([r["source"], to_date(r["start"]), to_date(r["end"]), r.get("note", "")])
        style_header(audit, 1, 4)
        widths(audit, {"A": 30, "B": 20, "C": 20, "D": 65})
        for row in audit.iter_rows(min_row=2, min_col=2, max_col=3):
            for cell in row:
                cell.number_format = DATE

    # ── Transactions ──
    ws = wb.active
    ws.title = "Transactions"
    for i, h in enumerate(TXCOLS, 1):
        ws.cell(1, i, h)
    style_header(ws, 1, len(TXCOLS))
    for i, r in enumerate(rows, 2):
        t = r["t"]
        review = {"guess": "check", "none": "check", "unknown": "unknown"}.get(r["conf"], "")
        vals = [t["id"], r["date"], r["month"], WEEKDAYS[r["date"].weekday()], t["card"],
                r["merchant"], r["cat"], r["sub"], t["type"], t["currency"], r["amt"],
                f'=IF({TC["Currency"]}{i}="TRY",1,INDEX(FX!$D:$D,MATCH({TC["FX key"]}{i},FX!$A:$A,0)))',
                f'=ROUND({TC["Amount"]}{i}*{TC["FX rate"]}{i},2)', r["counts"],
                f'={TC["Amount TRY"]}{i}*{TC["Counts as spending"]}{i}', r["channel"], r["tag"],
                t["statement_id"], r["key"], r["pay"], f'={TC["Payment date"]}{i}-{TC["Date"]}{i}',
                t["bank_category"], t["description"], t["installment"],
                float(t["bonus"]) if t["bonus"] else None, review, r["rotation"],
                "; ".join(x for x in (t.get("note") or "", t.get("memo") or "", r["rnote"]) if x),
                t["source_file"],
                f'=IFERROR(INDEX(Daily!$D:$D,MATCH({TC["Date"]}{i},Daily!$A:$A,0)),0)',
                f'=IFERROR({TC["Complete day"]}{i}*(1-INDEX(Daily!$I:$I,MATCH({TC["Date"]}{i},Daily!$A:$A,0))),0)',
                r.get("spend_month", ""), round(adv_principal[t["id"]], 2) if t["id"] in adv_principal else None]
        for j, v in enumerate(vals, 1):
            c = put(ws, i, j, v)
            h = TXCOLS[j - 1]
            if h in ("Date", "Payment date"):
                c.number_format = DATE
            elif h in ("Amount", "Amount TRY", "Spending TRY", "Bonus", "Adv principal"):
                c.number_format = MONEY
            elif h == "FX rate":
                c.number_format = "0.0000"
            elif h == "Float days":
                c.number_format = "0"
            if review:
                ws.cell(i, TXCOLS.index("Review") + 1).fill = REVIEW_FILL
    n_tx = len(rows) + 1
    tbl = Table(displayName="tblTransactions", ref=f"A1:{get_column_letter(len(TXCOLS))}{n_tx}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(tbl)
    ws.freeze_panes = "C2"
    tx_widths = {"ID": 30, "Date": 11, "Month": 9, "Weekday": 8, "Card": 18, "Merchant": 32, "Category": 22,
                 "Subcategory": 22, "Type": 12, "Currency": 8, "Amount": 11, "FX rate": 9, "Amount TRY": 12,
                 "Counts as spending": 9, "Spending TRY": 12, "Channel": 15, "Tag": 22, "Statement ID": 30,
                 "FX key": 15, "Payment date": 11, "Float days": 8, "Bank category": 16, "Description (raw)": 34,
                 "Installment": 16, "Bonus": 7, "Review": 9, "Rotation": 18, "Note": 40, "Source file": 32,
                 "Complete day": 9, "Normal day": 9, "Spend month": 10, "Adv principal": 11}
    widths(ws, {TC[k]: v for k, v in tx_widths.items()})

    # ── Daily ──
    wd = wb.create_sheet("Daily")
    dh = ["Date", "Weekday", "Month", "Complete", "Spending TRY", "Purchases", "Cumulative TRY", "7-day avg TRY",
          "Trip day"]
    dh += [f"Covered: {c}" + (" (excluded)" if c in excluded else "") for c in cards]
    for i, h in enumerate(dh, 1):
        wd.cell(1, i, h)
    style_header(wd, 1, len(dh))
    for i, d in enumerate(days, 2):
        covered = [1 if d in known_days[c] else 0 for c in cards]
        put(wd, i, 1, d, DATE)
        put(wd, i, 2, WEEKDAYS[d.weekday()])
        put(wd, i, 3, d.strftime("%Y-%m"))
        put(wd, i, 4, 1 if complete(d) else 0)
        put(wd, i, 5, f'=SUMIFS({tx("Spending TRY")},{tx("Date")},A{i})', MONEY)
        put(wd, i, 6, f'=COUNTIFS({tx("Date")},A{i},{tx("Type")},"Purchase")')
        put(wd, i, 7, f"=E{i}" if i == 2 else f"=G{i-1}+E{i}", MONEY)
        put(wd, i, 8, f"=AVERAGE(E{max(2, i-6)}:E{i})", MONEY)
        put(wd, i, 9, 1 if any(a <= d <= b for a, b in trip_ranges) else 0)
        for j, v in enumerate(covered):
            put(wd, i, 10 + j, v)
    n_days = len(days) + 1
    wd.freeze_panes = "B2"
    widths(wd, {"A": 11, "B": 9, "C": 9, "D": 9, "E": 13, "F": 10, "G": 14, "H": 13})
    for j in range(len(cards)):
        wd.column_dimensions[get_column_letter(10 + j)].width = 24
    card_cov_col = {c: get_column_letter(10 + j) for j, c in enumerate(cards)}

    # ── Monthly ──
    wm = wb.create_sheet("Monthly")
    mh = ["Month", "Spending TRY (all known)", "Card purchases", "Complete days", "Spending on complete days TRY",
          "Daily avg TRY (complete days)", "Days pending", "Est. full month TRY", "Card spending TRY",
          "Via bank account TRY (net)", "Friends paid back TRY", "Abroad TRY", "Refunds & discounts TRY",
          "Bank fees TRY", "Cash advance TRY (not spending)", "Income TRY", "MoM change", "Statement coverage"]
    for i, h in enumerate(mh, 1):
        wm.cell(1, i, h)
    style_header(wm, 1, len(mh))
    for i, mo in enumerate(months, 2):
        y, m = map(int, mo.split("-"))
        ndays = calendar.monthrange(y, m)[1]
        cov = ", ".join(f"{c.split(' •')[0]} {sum(1 for d in known_days[c] if d.year == y and d.month == m)}/{ndays}d"
                        for c in cards if c not in excluded)
        put(wm, i, 1, mo, bold=True)
        put(wm, i, 2, f'=SUMIFS({tx("Spending TRY")},{tx("Month")},A{i})', MONEY)
        put(wm, i, 3, f'=COUNTIFS({tx("Month")},A{i},{tx("Type")},"Purchase")')
        put(wm, i, 4, f'=COUNTIFS(Daily!$C:$C,A{i},Daily!$D:$D,1)')
        put(wm, i, 5, f'=SUMIFS(Daily!$E:$E,Daily!$C:$C,A{i},Daily!$D:$D,1)', MONEY)
        put(wm, i, 6, f"=IF(D{i}=0,0,E{i}/D{i})", MONEY)
        put(wm, i, 7, f'=COUNTIFS(Daily!$C:$C,A{i},Daily!$D:$D,0)')
        put(wm, i, 8, f"=F{i}*{ndays}", MONEY)
        acct = f'{tx("Card")},"*account*"'
        put(wm, i, 9, f'=B{i}-SUMIFS({tx("Spending TRY")},{tx("Month")},A{i},{acct})', MONEY)
        put(wm, i, 10, f'=SUMIFS({tx("Spending TRY")},{tx("Month")},A{i},{acct})', MONEY)
        put(wm, i, 11, f'=-SUMIFS({tx("Amount TRY")},{tx("Month")},A{i},{tx("Type")},"Reimbursement")', MONEY)
        put(wm, i, 12, f'=SUMIFS({tx("Spending TRY")},{tx("Month")},A{i},{tx("Channel")},"Abroad")', MONEY)
        put(wm, i, 13, f'=-(SUMIFS({tx("Amount TRY")},{tx("Month")},A{i},{tx("Type")},"Refund")'
                       f'+SUMIFS({tx("Amount TRY")},{tx("Month")},A{i},{tx("Type")},"Discount"))', MONEY)
        put(wm, i, 14, f'=SUMIFS({tx("Amount TRY")},{tx("Month")},A{i},{tx("Type")},"Fee")', MONEY)
        put(wm, i, 15, f'=SUMIFS({tx("Amount TRY")},{tx("Month")},A{i},{tx("Type")},"Cash advance")', MONEY)
        put(wm, i, 16, f'=-SUMIFS({tx("Amount TRY")},{tx("Month")},A{i},{tx("Type")},"Income")', MONEY)
        put(wm, i, 17, "" if i == 2 else f"=IF(B{i-1}=0,\"\",B{i}/B{i-1}-1)", PCT)
        put(wm, i, 18, cov, italic=True, color="555555")
    tr = len(months) + 2
    put(wm, tr, 1, "Total", bold=True, fill=SUB_FILL)
    for col in range(2, 17):
        if col == 8:
            continue
        L = get_column_letter(col)
        f = f"=IF(D{tr}=0,0,E{tr}/D{tr})" if col == 6 else f"=SUM({L}2:{L}{tr-1})"
        put(wm, tr, col, f, MONEY if col not in (3, 4, 7) else None, bold=True, fill=SUB_FILL)
    put(wm, tr + 2, 1, "Complete day = every active card/account has imported coverage or an explicit account "
                       "no-activity confirmation in data/confirmed_no_activity.csv. Accounts listed in "
                       "data/excluded_coverage_sources.csv do not gate completeness. Daily averages use complete days only; pending "
                       "days fill in when the missing statement arrives. Est. full month = complete-day average "
                       "× days in the month.", italic=True, color="555555")
    wm.freeze_panes = "B2"
    widths(wm, {"A": 10, "B": 14, "C": 10, "D": 10, "E": 14, "F": 14, "G": 9, "H": 14, "I": 13, "J": 14,
                "K": 13, "L": 12, "M": 14, "N": 12, "O": 16, "P": 13, "Q": 10, "R": 70})

    # ── Categories ──
    wc = wb.create_sheet("Categories")
    cat_tot = defaultdict(float)
    sub_tot = defaultdict(float)
    for r in spend_rows:
        cat_tot[r["cat"]] += r["try_amt"]
        sub_tot[(r["cat"], r["sub"])] += r["try_amt"]
    cats_sorted = sorted(cat_tot, key=lambda c: -cat_tot[c])
    ch_ = ["Category"] + months + ["Total TRY", "Share", "Purchases", "Avg purchase TRY"]
    for i, h in enumerate(ch_, 1):
        wc.cell(1, i, h)
    style_header(wc, 1, len(ch_))
    tcol = get_column_letter(len(months) + 2)
    for i, c in enumerate(cats_sorted, 2):
        put(wc, i, 1, c, bold=True)
        for j, mo in enumerate(months):
            L = get_column_letter(2 + j)
            put(wc, i, 2 + j, f'=SUMIFS({tx("Spending TRY")},{tx("Category")},$A{i},{tx("Month")},{L}$1)', MONEY)
        k = len(months) + 2
        put(wc, i, k, f"=SUM(B{i}:{get_column_letter(k-1)}{i})", MONEY, bold=True)
        put(wc, i, k + 1, f"={tcol}{i}/{tcol}${len(cats_sorted)+2}", PCT)
        put(wc, i, k + 2, f'=COUNTIFS({tx("Category")},$A{i},{tx("Type")},"Purchase")')
        put(wc, i, k + 3, f'=IF({get_column_letter(k+2)}{i}=0,0,SUMIFS({tx("Amount TRY")},{tx("Category")},$A{i},'
                          f'{tx("Type")},"Purchase")/{get_column_letter(k+2)}{i})', MONEY)
    ctr = len(cats_sorted) + 2
    put(wc, ctr, 1, "Total", bold=True, fill=SUB_FILL)
    for col in range(2, len(months) + 3):
        L = get_column_letter(col)
        put(wc, ctr, col, f"=SUM({L}2:{L}{ctr-1})", MONEY, bold=True, fill=SUB_FILL)
    put(wc, ctr, len(months) + 3, f"=SUM({get_column_letter(len(months)+3)}2:{get_column_letter(len(months)+3)}{ctr-1})",
        PCT, bold=True, fill=SUB_FILL)
    sr = ctr + 3
    put(wc, sr - 1, 1, "Subcategory breakdown", bold=True, size=12)
    for i, h in enumerate(["Category", "Subcategory", "Total TRY", "Purchases", "Share of category",
                           "Share of all spending"], 1):
        wc.cell(sr, i, h)
    style_header(wc, sr, 6)
    subs_sorted = sorted(sub_tot, key=lambda k: (cats_sorted.index(k[0]), -sub_tot[k]))
    for i, (c, s) in enumerate(subs_sorted, sr + 1):
        put(wc, i, 1, c)
        put(wc, i, 2, s)
        put(wc, i, 3, f'=SUMIFS({tx("Spending TRY")},{tx("Category")},A{i},{tx("Subcategory")},B{i})', MONEY)
        put(wc, i, 4, f'=COUNTIFS({tx("Category")},A{i},{tx("Subcategory")},B{i},{tx("Type")},"Purchase")')
        put(wc, i, 5, f"=IFERROR(C{i}/INDEX(${tcol}$2:${tcol}${ctr-1},MATCH(A{i},$A$2:$A${ctr-1},0)),0)", PCT)
        put(wc, i, 6, f"=C{i}/${tcol}${ctr}", PCT)
    wc.freeze_panes = "B2"
    wc.column_dimensions["A"].width = 26
    wc.column_dimensions["B"].width = 24
    for col in range(3, len(months) + 6):
        wc.column_dimensions[get_column_letter(col)].width = 14

    # ── Merchants ──
    wmr = wb.create_sheet("Merchants")
    m_tot, m_cat = defaultdict(float), {}
    for r in spend_rows:
        if r["cat"] in STORE_EXCLUDE:
            continue
        m_tot[r["merchant"]] += r["try_amt"]
        m_cat.setdefault(r["merchant"], r["cat"])
    merch_sorted = sorted(m_tot, key=lambda m: -m_tot[m])
    mh = ["Merchant", "Category", "Visits", "Total TRY", "Avg ticket TRY", "First visit", "Last visit", "Share",
          "Cumulative share"]
    for i, h in enumerate(mh, 1):
        wmr.cell(1, i, h)
    style_header(wmr, 1, len(mh))
    mlast = len(merch_sorted) + 1
    for i, m in enumerate(merch_sorted, 2):
        put(wmr, i, 1, m, bold=True)
        put(wmr, i, 2, m_cat[m])
        put(wmr, i, 3, f'=COUNTIFS({tx("Merchant")},A{i},{tx("Type")},"Purchase")')
        put(wmr, i, 4, f'=SUMIFS({tx("Spending TRY")},{tx("Merchant")},A{i})', MONEY)
        put(wmr, i, 5, f"=IF(C{i}=0,0,SUMIFS({tx('Amount TRY')},{tx('Merchant')},A{i},{tx('Type')},\"Purchase\")/C{i})", MONEY)
        put(wmr, i, 6, f'=_xlfn.MINIFS({tx("Date")},{tx("Merchant")},A{i},{tx("Counts as spending")},1)', DATE)
        put(wmr, i, 7, f'=_xlfn.MAXIFS({tx("Date")},{tx("Merchant")},A{i},{tx("Counts as spending")},1)', DATE)
        put(wmr, i, 8, f"=D{i}/SUM($D$2:$D${mlast})", PCT)
        put(wmr, i, 9, f"=SUM($D$2:D{i})/SUM($D$2:$D${mlast})", PCT)
    wmr.freeze_panes = "B2"
    widths(wmr, {"A": 34, "B": 24, "C": 8, "D": 13, "E": 14, "F": 11, "G": 11, "H": 9, "I": 12})

    # ── FX ──
    wf = wb.create_sheet("FX")
    fh = ["Key", "Currency", "Payment date", "Rate (TRY)", "Rate date", "Source", "Status"]
    for i, h in enumerate(fh, 1):
        wf.cell(1, i, h)
    style_header(wf, 1, len(fh))
    for i, r in enumerate(fx_rows, 2):
        put(wf, i, 1, r["key"])
        put(wf, i, 2, r["currency"])
        put(wf, i, 3, to_date(r["payment_date"]), DATE)
        put(wf, i, 4, float(r["rate"]) if r["rate"] else None, "0.0000", color="0000FF",
            fill=INPUT_FILL if not r["rate"] else None)
        put(wf, i, 5, to_date(r["rate_date"]), DATE)
        put(wf, i, 6, r["source"])
        put(wf, i, 7, r["status"])
    put(wf, len(fx_rows) + 3, 1, "Rule: a foreign-currency statement is converted at the rate it was actually paid "
                                 "at. 'derived' = the bank's own conversion, read from the next statements "
                                 "('TLye dönen … Borç Ödeme' ÷ '… ekstre devir bakiyesi'). Until then: TCMB forex "
                                 "selling rate on the payment (son ödeme) date ('final' / 'provisional'). To force "
                                 "a rate, edit data/fx_rates.csv and set status 'manual'.",
        italic=True, color="555555")
    widths(wf, {"A": 16, "B": 9, "C": 12, "D": 11, "E": 11, "F": 70, "G": 12})

    # ── Statements ──
    wsx = wb.create_sheet("Statements")
    sh = ["Statement ID", "Bank", "Card", "Currency", "Period start", "Cutoff", "Payment date", "Prev balance",
          "Spending", "Fees & interest", "Payments & credits", "Balance due", "Min payment", "FX key", "FX rate",
          "Balance due TRY", "Ledger sum", "Check", "Ekstre no", "Source file", "Kind", "Same-bank account covers it"]
    # your upload rule: every card statement comes with the same bank's account export covering at least its period
    acct_days = defaultdict(set)
    for s in sts:
        if (s.get("kind") or "Card") == "Account":
            acct_days[s["bank"]] |= known_days[s["card"]]

    def account_cover(s):
        if (s.get("kind") or "Card") == "Account":
            return ""
        need = {s["window_start"] + dt.timedelta(k) for k in range((s["cutoff"] - s["window_start"]).days + 1)}
        miss = need - acct_days.get(s["bank"], set())
        return "yes" if not miss else f"missing {len(miss)} of {len(need)} days"
    for i, h in enumerate(sh, 1):
        wsx.cell(1, i, h)
    style_header(wsx, 1, len(sh))
    for i, s in enumerate(sorted(sts, key=lambda s: (s["cutoff"], s["statement_id"])), 2):
        key = "" if s["currency"] == "TRY" else f"{s['currency']}|{s['payment_date']}"
        vals = [s["statement_id"], s["bank"], s["card"], s["currency"], s["window_start"], s["cutoff"],
                s["payment_date"], float(s["prev_balance"]), float(s["spending"]), float(s["fees_interest"]),
                float(s["payments_credits"]), float(s["balance_due"]),
                float(s["min_payment"]) if s["min_payment"] else None, key,
                f'=IF(D{i}="TRY",1,INDEX(FX!$D:$D,MATCH(N{i},FX!$A:$A,0)))', f"=L{i}*O{i}",
                f'=SUMIFS({tx("Amount")},{tx("Statement ID")},A{i})',
                f'=IF(ABS(IF(U{i}="Account",H{i}-Q{i},H{i}+Q{i})-L{i})<0.01,"OK","MISMATCH")',
                s["ekstre_no"], s["source_file"], s.get("kind") or "Card", account_cover(s)]
        for j, v in enumerate(vals, 1):
            fmt = DATE if j in (5, 6, 7) else ("0.0000" if j == 15 else (MONEY if 8 <= j <= 17 else None))
            put(wsx, i, j, v, fmt)
    put(wsx, len(sts) + 3, 1, "Period start is inferred when the statement doesn't print the previous cutoff. "
                              "Check = previous balance + all ledger rows of that statement equals the balance due.",
        italic=True, color="555555")
    widths(wsx, {"A": 30, "B": 12, "C": 18, "D": 8, "E": 11, "F": 11, "G": 11, "H": 12, "I": 12, "J": 11,
                 "K": 13, "L": 12, "M": 12, "N": 15, "O": 9, "P": 14, "Q": 12, "R": 9, "S": 26, "T": 34, "U": 9,
                 "V": 24})

    # ── Insights ──
    wi = wb.create_sheet("Insights")
    wi.column_dimensions["A"].width = 36
    for col in "BCDEFGHIJKL":
        wi.column_dimensions[col].width = 15
    r = 1
    put(wi, r, 1, "Insights", bold=True, size=16)
    r += 2
    def section(title, sub=None):
        nonlocal r
        put(wi, r, 1, title, bold=True, size=12, color="1F3A5F")
        r += 1
        if sub:
            put(wi, r, 1, sub, italic=True, color="555555")
            r += 1

    def header(cols):
        nonlocal r
        for i, h in enumerate(cols, 1):
            wi.cell(r, i, h)
        style_header(wi, r, len(cols))
        r += 1

    # 1. by card + interest-free float
    def rng(name):
        return f"Transactions!${TC[name]}$2:${TC[name]}${n_tx}"

    section("By card", "Float days = days from purchase to the statement's payment date, i.e. how long the bank "
                       "lends you the money for free (amount-weighted average).")
    header(["Card", "Spending TRY", "Share", "Purchases", "Avg float days", "Days covered"])
    c0 = r
    for c in cards:
        put(wi, r, 1, c, bold=True)
        put(wi, r, 2, f'=SUMIFS({tx("Spending TRY")},{tx("Card")},A{r})', MONEY)
        put(wi, r, 3, f"=B{r}/SUM($B${c0}:$B${c0+len(cards)-1})", PCT)
        put(wi, r, 4, f'=COUNTIFS({tx("Card")},A{r},{tx("Type")},"Purchase")')
        put(wi, r, 5, f'=IFERROR(SUMPRODUCT(({rng("Card")}=A{r})*({rng("Type")}="Purchase")*{rng("Amount TRY")}'
                      f'*{rng("Float days")})/SUMIFS({tx("Amount TRY")},{tx("Card")},A{r},{tx("Type")},"Purchase"),0)',
            "0.0")
        put(wi, r, 6, f"=SUM(Daily!{card_cov_col[c]}:{card_cov_col[c]})")
        r += 1
    r += 1
    section("Card rotation check",
            f"Your rule: QNB from the 2nd to the 14th, Garanti from the 16th to the 1st (the 15th counts for either). "
            f"QNB closes on the 1st, so on the 1st Garanti gives the longer interest-free period. Checked from "
            f"{ROTATION_START:%d.%m.%Y}; August 2026 was the new-card exception.")
    put(wi, r, 1, "Off-rotation purchases")
    put(wi, r, 2, f'=COUNTIF({tx("Rotation")},"off-rotation*")')
    r += 1
    put(wi, r, 1, "Amount put on the 'wrong' card TRY")
    put(wi, r, 2, f'=SUMIFS({tx("Amount TRY")},{tx("Rotation")},"off-rotation*")', MONEY)
    r += 2

    # 2. trips
    if tags:
        section("Trips & events", "Defined in data/tags.csv. IBAN transfers friends sent you from the trip start "
                                  "until 'reimbursements_until' count as them paying back their share.")
        header(["Tag", "From", "To", "Days", "Purchases", "Spend (orig.)", "Currency", "Gross TRY",
                "Friends paid back TRY", "Net TRY", "Net per day TRY"])
        for g in tags:
            put(wi, r, 1, g["tag"], bold=True)
            put(wi, r, 2, to_date(g["start"]), DATE)
            put(wi, r, 3, to_date(g["end"]), DATE)
            put(wi, r, 4, f"=C{r}-B{r}+1")
            put(wi, r, 5, f'=COUNTIFS({tx("Tag")},A{r},{tx("Type")},"Purchase")')
            put(wi, r, 6, f'=SUMIFS({tx("Amount")},{tx("Tag")},A{r},{tx("Currency")},G{r},'
                          f'{tx("Type")},"<>Reimbursement")', MONEY)
            put(wi, r, 7, g["currency"] or "TRY")
            put(wi, r, 8, f'=SUMIFS({tx("Spending TRY")},{tx("Tag")},A{r},{tx("Type")},"<>Reimbursement")', MONEY)
            put(wi, r, 9, f'=-SUMIFS({tx("Amount TRY")},{tx("Tag")},A{r},{tx("Type")},"Reimbursement")', MONEY)
            put(wi, r, 10, f"=H{r}-I{r}", MONEY, bold=True)
            put(wi, r, 11, f"=IF(D{r}=0,0,J{r}/D{r})", MONEY)
            r += 1
        r += 1

    # 3. top 10 purchases
    section("10 biggest purchases")
    header(["#", "ID", "Date", "Merchant", "Category", "Amount TRY"])
    top = sorted([x for x in rows if x["t"]["type"] == "Purchase"], key=lambda x: -x["try_amt"])[:10]
    for k, x in enumerate(top, 1):
        put(wi, r, 1, k)
        put(wi, r, 2, x["t"]["id"], color="777777")
        put(wi, r, 3, f'=INDEX({tx("Date")},MATCH(B{r},{tx("ID")},0))', DATE)
        put(wi, r, 4, f'=INDEX({tx("Merchant")},MATCH(B{r},{tx("ID")},0))')
        put(wi, r, 5, f'=INDEX({tx("Category")},MATCH(B{r},{tx("ID")},0))')
        put(wi, r, 6, f'=INDEX({tx("Amount TRY")},MATCH(B{r},{tx("ID")},0))', MONEY)
        r += 1
    wi.column_dimensions["B"].width = 30
    r += 1

    # 4. eating out vs groceries
    section("Food: eating out vs. cooking")
    header(["Where the food money goes", "Spending TRY", "Purchases", "Share of food"])
    f0 = r
    food_subs = sorted({x["sub"] for x in spend_rows if x["cat"] == "Food & Dining"},
                       key=lambda s: -sub_tot[("Food & Dining", s)])
    for s in food_subs:
        put(wi, r, 1, f"Food & Dining – {s}")
        wi.cell(r, 13, s)  # hidden helper key (column M)
        put(wi, r, 2, f'=SUMIFS({tx("Spending TRY")},{tx("Category")},"Food & Dining",{tx("Subcategory")},M{r})', MONEY)
        put(wi, r, 3, f'=COUNTIFS({tx("Category")},"Food & Dining",{tx("Subcategory")},M{r},{tx("Type")},"Purchase")')
        r += 1
    put(wi, r, 1, "Groceries")
    put(wi, r, 2, f'=SUMIFS({tx("Spending TRY")},{tx("Category")},"Groceries")', MONEY)
    put(wi, r, 3, f'=COUNTIFS({tx("Category")},"Groceries",{tx("Type")},"Purchase")')
    r += 1
    for k in range(f0, r):
        put(wi, k, 4, f"=IF(SUM($B${f0}:$B${r-1})=0,0,B{k}/SUM($B${f0}:$B${r-1}))", PCT)
    wi.column_dimensions["M"].hidden = True
    put(wi, r, 1, "Eating out per 1 TRY of groceries", bold=True, fill=SUB_FILL)
    put(wi, r, 2, f"=IF(B{r-1}=0,0,SUM(B{f0}:B{r-2})/B{r-1})", "0.0x", bold=True, fill=SUB_FILL)
    r += 2

    # 5. transport
    section("Transport mix")
    header(["Mode", "Spending TRY", "Trips / top-ups", "Avg TRY"])
    for s in sorted({x["sub"] for x in spend_rows if x["cat"] == "Transport"}, key=lambda s: -sub_tot[("Transport", s)]):
        put(wi, r, 1, s)
        put(wi, r, 2, f'=SUMIFS({tx("Spending TRY")},{tx("Category")},"Transport",{tx("Subcategory")},A{r})', MONEY)
        put(wi, r, 3, f'=COUNTIFS({tx("Category")},"Transport",{tx("Subcategory")},A{r},{tx("Type")},"Purchase")')
        put(wi, r, 4, f"=IF(C{r}=0,0,B{r}/C{r})", MONEY)
        r += 1
    r += 1

    # 6. small tickets
    section("Small-ticket habit", "Purchases under the threshold (blue cell, change it freely).")
    put(wi, r, 1, "Threshold TRY")
    put(wi, r, 2, 100, MONEY0, color="0000FF", fill=INPUT_FILL)
    thr = f"$B${r}"
    r += 1
    put(wi, r, 1, "Number of small purchases")
    put(wi, r, 2, f'=COUNTIFS({tx("Amount TRY")},"<"&{thr},{tx("Type")},"Purchase")')
    put(wi, r, 3, f'=B{r}/COUNTIF({tx("Type")},"Purchase")', PCT)
    put(wi, r, 4, "of all purchases", italic=True, color="555555")
    r += 1
    put(wi, r, 1, "Money spent on them TRY")
    put(wi, r, 2, f'=SUMIFS({tx("Amount TRY")},{tx("Amount TRY")},"<"&{thr},{tx("Type")},"Purchase")', MONEY)
    put(wi, r, 3, f'=B{r}/SUMIFS({tx("Amount TRY")},{tx("Type")},"Purchase")', PCT)
    put(wi, r, 4, "of purchase spending", italic=True, color="555555")
    r += 2

    # 7. weekday pattern
    section("Weekday pattern", "Only complete days are counted.")
    header(["Weekday", "Spending TRY", "Covered days", "Avg per day TRY"])
    w0 = r
    for wdn in WEEKDAYS:
        put(wi, r, 1, wdn)
        put(wi, r, 2, f'=SUMIFS(Daily!$E:$E,Daily!$B:$B,A{r},Daily!$D:$D,1)', MONEY)
        put(wi, r, 3, f'=COUNTIFS(Daily!$B:$B,A{r},Daily!$D:$D,1)')
        put(wi, r, 4, f"=IF(C{r}=0,0,B{r}/C{r})", MONEY)
        r += 1
    put(wi, r, 1, "Weekdays (Mon–Fri) avg / day", bold=True, fill=SUB_FILL)
    put(wi, r, 4, f"=IF(SUM(C{w0}:C{w0+4})=0,0,SUM(B{w0}:B{w0+4})/SUM(C{w0}:C{w0+4}))", MONEY, bold=True, fill=SUB_FILL)
    r += 1
    put(wi, r, 1, "Weekend (Sat–Sun) avg / day", bold=True, fill=SUB_FILL)
    put(wi, r, 4, f"=IF(SUM(C{w0+5}:C{w0+6})=0,0,SUM(B{w0+5}:B{w0+6})/SUM(C{w0+5}:C{w0+6}))", MONEY, bold=True, fill=SUB_FILL)
    r += 2

    # 8. cash advance tracker
    def installment(x):
        mm = re.match(r"(\d+)/(\d+) of ([\d.]+)", x["t"]["installment"] or "")
        return (int(mm.group(1)), int(mm.group(2)), float(mm.group(3))) if mm else (1, 1, x["amt"])

    # one line per advance: its installments (1/3, 2/3, …) arrive on successive statements; keep the latest
    adv_groups = defaultdict(list)
    for x in rows:
        if x["t"]["type"] == "Cash advance":
            _, n, tot = installment(x)
            adv_groups[(x["t"]["card"], n, round(tot, 2))].append(x)
    adv = [max(g, key=lambda x: installment(x)[0]) for g in adv_groups.values()]
    section("Cash advance tracker",
            "Cash advances are borrowing, so they are NOT in spending. The principal is filled in from the "
            "bank account when it's there; otherwise enter what you received (yellow) to see what the loan costs.")
    header(["ID", "Date", "Installment TRY", "Paid so far", "Of", "Total repayable", "Remaining TRY",
            "Principal received", "Cost of borrowing"])
    a0 = r
    for x in adv:
        k, n, tot = installment(x)
        taken = add_months(x["date"], -(k - 1))  # installment k is billed k-1 months after the advance
        put(wi, r, 1, x["t"]["id"], color="777777")
        put(wi, r, 2, taken, DATE)
        put(wi, r, 3, f'=INDEX({tx("Amount TRY")},MATCH(A{r},{tx("ID")},0))', MONEY)
        put(wi, r, 4, k, color="0000FF")
        put(wi, r, 5, n, color="0000FF")
        put(wi, r, 6, tot, MONEY, color="0000FF")
        put(wi, r, 7, f"=C{r}*(E{r}-D{r})", MONEY, bold=True)
        got = next((y for y in rows if y["t"]["type"] == "Cash advance received"
                    and abs((y["date"] - taken).days) <= 2), None)
        put(wi, r, 8, -got["amt"] if got else None, MONEY, color="0000FF", fill=None if got else INPUT_FILL)
        put(wi, r, 9, f'=IF(H{r}="","",F{r}-H{r})', MONEY)
        r += 1
    put(wi, r, 1, "Still to be billed", bold=True, fill=SUB_FILL)
    put(wi, r, 7, f"=SUM(G{a0}:G{r-1})" if adv else 0, MONEY, bold=True, fill=SUB_FILL)
    r += 1
    put(wi, r, 1, "Paid/of/total come from the statement text; the installment counter only moves when a new "
                  "statement is imported.", italic=True, color="555555")
    r += 2

    has_account = any(x["t"]["section"] == "Account" for x in rows)

    # 9a. card payments seen in the bank account
    card_groups = defaultdict(list)
    for s in sts:
        if (s.get("kind") or "Card") != "Account":
            card_groups[(s["card"], s["cutoff"])].append(s)
    if has_account:
        section("Card payments seen in your bank account",
                "Matched by the card's last 4 digits, paid after the cutoff and up to the card's next cutoff. "
                "'Days after due' above 0 "
                "means the automatic payment couldn't take the money on the due date: the whole balance then earns "
                "interest from the cutoff, and an unpaid EUR/USD balance is converted at the bank's own rate. "
                "The last column shows the interest that landed on the next statement.")
        header(["Card", "Cutoff", "Due date", "Statement balance", "Paid TRY", "Last payment", "Days after due",
                "TL part unpaid", "Interest on next statement"])
        for (card, cutoff), g in sorted(card_groups.items(), key=lambda kv: (kv[0][1], kv[0][0])):
            # payments after this cutoff up to the card's next cutoff (a later payment settles the next bill)
            nxt = min((c for (cd, c) in card_groups if cd == card and c > cutoff),
                      default=cutoff + dt.timedelta(days=35))
            crit = (f'{tx("Type")},"Payment",{tx("Card")},"*account*",{tx("Description (raw)")},"*{card[-4:]}*",'
                    f'{tx("Date")},">"&B{r},{tx("Date")},"<="&DATE({nxt.year},{nxt.month},{nxt.day})')
            put(wi, r, 1, card, bold=True)
            put(wi, r, 2, cutoff, DATE)
            put(wi, r, 3, g[0]["payment_date"], DATE)
            put(wi, r, 4, " + ".join(f"{float(s['balance_due']):,.2f} {s['currency']}"
                                     for s in sorted(g, key=lambda s: s["currency"] != "TRY")
                                     if s["currency"] == "TRY" or float(s["balance_due"])))
            put(wi, r, 5, f"=SUMIFS({tx('Amount TRY')},{crit})", MONEY)
            put(wi, r, 6, f'=IF(E{r}=0,"",_xlfn.MAXIFS({tx("Date")},{crit}))', DATE)
            put(wi, r, 7, f'=IF(F{r}="",IF(C{r}>=TODAY(),"not due yet","not in account data"),F{r}-C{r})')
            put(wi, r, 8, f'=IF(E{r}=0,"",MAX(0,ROUND(SUMIFS(Statements!$L:$L,Statements!$C:$C,A{r},'
                          f'Statements!$F:$F,B{r},Statements!$D:$D,"TRY")-E{r},2)))', MONEY, bold=True)
            put(wi, r, 9, f'=SUMIFS({tx("Amount TRY")},{tx("Card")},A{r},{tx("Subcategory")},"Interest & tax",'
                          f'{tx("Date")},">"&C{r},{tx("Date")},"<="&(B{r}+40))', MONEY, bold=True)
            r += 1
        r += 1

    # 9b. money with friends
    people_rows = defaultdict(float)
    for x in rows:
        if x["cat"] == "People (IBAN)" and x["t"]["type"] in ("Transfer out", "Reimbursement"):
            people_rows[x["merchant"]] += abs(x["try_amt"])
    if people_rows:
        section("Money with friends (IBAN)",
                "Sent = you paid them (counted as spending). Received = they paid you back (subtracted from "
                "spending). Transfers that came straight back are left out. Label people in data/people.csv.")
        header(["Person", "Sent TRY", "Received TRY", "Net TRY", "Transfers"])
        p0 = r
        for p in sorted(people_rows, key=lambda p: -people_rows[p]):
            put(wi, r, 1, p)
            put(wi, r, 2, f'=SUMIFS({tx("Amount TRY")},{tx("Merchant")},A{r},{tx("Type")},"Transfer out")', MONEY)
            put(wi, r, 3, f'=-SUMIFS({tx("Amount TRY")},{tx("Merchant")},A{r},{tx("Type")},"Reimbursement")', MONEY)
            put(wi, r, 4, f"=B{r}-C{r}", MONEY)
            put(wi, r, 5, f'=COUNTIFS({tx("Merchant")},A{r},{tx("Type")},"Transfer out")'
                          f'+COUNTIFS({tx("Merchant")},A{r},{tx("Type")},"Reimbursement")')
            r += 1
        put(wi, r, 1, "Total", bold=True, fill=SUB_FILL)
        for col in (2, 3, 4, 5):
            L = get_column_letter(col)
            put(wi, r, col, f"=SUM({L}{p0}:{L}{r-1})", MONEY if col < 5 else None, bold=True, fill=SUB_FILL)
        r += 2

    # 9c. income vs typical month
    plan = read_csv(INCOME_PLAN_CSV)
    if has_account or plan:
        section("Income & what's left",
                "Income received, by the month it's for (a salary paid on the 4th for last month counts for last "
                "month). Blue = your current monthly income from data/income_plan.csv. 'Typical month' = the "
                "average spending per normal day (complete, no trips) × 30.")
        header(["Month", "Income received TRY"])
        for k, mo in enumerate(months):
            put(wi, r, 1, mo)
            put(wi, r, 2, f"=Monthly!P{k + 2}", MONEY)
            r += 1
        r += 1
        if plan:
            header(["Current monthly income", "TRY", "From", "Note"])
            p0 = r
            for item in plan:
                put(wi, r, 1, item["item"])
                put(wi, r, 2, float(item["amount"]), MONEY, color="0000FF")
                put(wi, r, 3, item["from"])
                put(wi, r, 4, item["note"], italic=True, color="555555")
                r += 1
            sal = next((p0 + k for k, it in enumerate(plan) if "salary" in it["item"].lower()), p0)
            typical = f"=IFERROR(SUMIFS(Daily!$E:$E,Daily!$D:$D,1,Daily!$I:$I,0)/{NORMAL_DAYS}*30,0)"
            for label, formula, fmt in (
                    ("Total monthly income", f"=SUM(B{p0}:B{r-1})", MONEY),
                    ("Typical month spending (cards + IBAN)", typical, MONEY),
                    ("Left over from salary", f"=B{sal}-B{r+1}", MONEY),
                    ("Share of salary kept", f"=IF(B{sal}=0,0,B{r+2}/B{sal})", PCT)):
                put(wi, r, 1, label, bold=True, fill=SUB_FILL)
                put(wi, r, 2, formula, fmt, bold=True, fill=SUB_FILL)
                r += 1
            put(wi, r, 1, "Meal-card (Pluxee) spending isn't in any bank data, so it's left out of both sides. "
                          "Compulsory BES is taken off before the salary reaches your account.",
                italic=True, color="555555")
            r += 1
        r += 1

    # 9d. money-market fund
    if any(x["t"]["type"] == "Investment" for x in rows):
        section("Money-market fund (GPY 1. Para Piyasa)", "Money moved between the account and the fund; not spending.")
        for label, formula in (
                ("Bought TRY", f'=SUMIFS({tx("Amount TRY")},{tx("Type")},"Investment",{tx("Subcategory")},"Fund buy")'),
                ("Sold TRY", f'=-SUMIFS({tx("Amount TRY")},{tx("Type")},"Investment",{tx("Subcategory")},"Fund sell")'),
                ("Net put into the fund TRY", f"=B{r}-B{r+1}"),
                ("Withholding tax (stopaj) & MKK fees TRY", f'=SUMIFS({tx("Amount TRY")},{tx("Type")},"Investment cost")')):
            put(wi, r, 1, label)
            put(wi, r, 2, formula, MONEY)
            r += 1
        r += 1

    # 9. things to identify
    unknown = [m for m in merch_sorted
               if any(x["merchant"] == m and x["conf"] in ("guess", "none", "unknown") for x in spend_rows)]
    section("Merchants to double-check",
            "'my guess' = category guessed from the name — tell Claude if it's wrong. 'not recognised' = you "
            "didn't recognise it either; the bank app shows each charge's location if you want to check.")
    header(["Merchant", "Category", "Visits", "Total TRY", "Status", "Raw description"])
    for m in unknown:
        x = next(x for x in spend_rows if x["merchant"] == m)
        put(wi, r, 1, m, fill=REVIEW_FILL)
        put(wi, r, 2, m_cat[m])
        put(wi, r, 3, f'=COUNTIFS({tx("Merchant")},A{r},{tx("Type")},"Purchase")')
        put(wi, r, 4, f'=SUMIFS({tx("Spending TRY")},{tx("Merchant")},A{r})', MONEY)
        put(wi, r, 5, "not recognised" if x["conf"] == "unknown" else "my guess")
        put(wi, r, 6, x["t"]["description"], color="777777")
        r += 1

    # ── Rules ──
    wr = wb.create_sheet("Rules")
    rh = ["Pattern (regex)", "Merchant", "Category", "Subcategory", "Confidence", "Note"]
    for i, h in enumerate(rh, 1):
        wr.cell(1, i, h)
    style_header(wr, 1, len(rh))
    for i, rule in enumerate(rules, 2):
        for j, k in enumerate(["pattern", "merchant", "category", "subcategory", "confidence", "note"], 1):
            put(wr, i, j, rule[k], fill=REVIEW_FILL if rule["confidence"] == "guess" and j == 5 else None)
    widths(wr, {"A": 34, "B": 32, "C": 24, "D": 24, "E": 11, "F": 70})
    wr.freeze_panes = "A2"

    # ── README ──
    wrm = wb.create_sheet("README")
    wrm.column_dimensions["A"].width = 120
    lines = [
        ("How this workbook works", True),
        ("This file is regenerated by harcama.py from the CSV files in data/. Don't type data here — it would be "
         "overwritten on the next rebuild. Change data/*.csv or ask Claude.", False),
        ("", False),
        ("Adding statements", True),
        ("1. Drop new ekstre PDFs into ekstreler/ and account exports (Garanti 'hesap hareketleri' .xlsx, QNB "
         "'Hesap Hareketleri' .pdf) into hesap_dokumu/. Overlapping date ranges are fine — rows already imported "
         "are skipped. Your current income lives in data/income_plan.csv.", False),
        ("2. Run:  python harcama.py   — only statements that aren't already in the ledger get imported.", False),
        ("3. The Statements sheet reconciles every statement against its printed balance (Check = OK).", False),
        ("", False),
        ("Definitions", True),
        ("Spending = card purchases + fees − refunds − discounts, plus from the bank account: money sent to "
         "people, ATM cash, transfer fees, minus friends paying you back. Not spending: card payments (the "
         "purchases are already counted), cash advances (borrowing), own-account transfers, fund buys/sells, "
         "salary, and transfers that came straight back.", False),
        ("Amount TRY: TRY statements as printed; EUR/USD statements × TCMB forex selling rate on that "
         "statement's payment date (FX sheet).", False),
        ("Complete day = every active card/account has imported coverage or an explicit bank-account "
         "no-activity confirmation in data/confirmed_no_activity.csv. Accounts listed in "
         "data/excluded_coverage_sources.csv do not gate completeness. Card rotation does not waive coverage. "
         "Daily averages use complete days only; 'normal days' also exclude trip days.", False),
        ("Float days = payment date − purchase date: how long each purchase was financed interest-free. The "
         "Rotation column flags purchases on the 'wrong' card (from Sep 2026 on).", False),
        ("Review = 'check' means the category is my best guess — see Insights ▸ 'Merchants I'm not sure "
         "about'.", False),
        ("", False),
        ("Dashboard", True),
        ("Open Harcamalarim.html (written by the same run) in any browser: pick a month or a range of months and "
         "every number, chart and table follows it. It starts at the first month with card statements; this "
         "workbook keeps every raw row and the History sheet. Monthly budgets per category go in "
         "data/budgets.csv.", False),
        ("", False),
        ("History", True),
        ("Month-by-month spending estimated from bank accounts alone (works for months without card statements): "
         "card payments are moved back to the month of their purchases, minus cash-advance principal, plus money to "
         "friends, ATM cash and fees. Also income, KYK loan, family support, investments and one-offs per month.",
         False),
        ("Upload rule: send each bank's card statement together with that bank's account export covering at least "
         "the same period. Statements ▸ 'Same-bank account covers it' checks this.", False),
        ("", False),
        ("Fixing a category", True),
        ("One transaction: fill category_override / merchant_override in data/transactions.csv. A merchant in "
         "general: edit data/category_rules.csv (first matching pattern wins). Then run  python harcama.py "
         "--build.", False),
    ]
    for i, (text, bold) in enumerate(lines, 1):
        c = put(wrm, i, 1, text, bold=bold, size=12 if bold else 10)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    build_history_sheet(wb, rows, sts)
    write_dashboard(OUT_HTML, build_payload(rows, sts, fx, cards, card_days, complete, trip_ranges, plan,
                                            load_budgets(cats_sorted), ROTATION_TEXT,
                                            extra_categories=rule_categories(rules), rotation_start=ROTATION_START,
                                            confirmations=confirmations, exclusions=exclusions))

    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT_XLSX)
    return rows


def recalc_with_excel(path: Path):
    """Let Excel compute every formula and store the values (so previews / pandas see numbers)."""
    ps = (f"$x=New-Object -ComObject Excel.Application;$x.DisplayAlerts=$false;"
          f"$w=$x.Workbooks.Open('{path}');$x.CalculateFull();$w.Save();$w.Close();$x.Quit()")
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, timeout=180,
                       capture_output=True)
        return True
    except Exception as e:
        print(f"  (Excel recalculation skipped: {e})")
        return False


def main():
    for d in (DATA, DOCS_DIR):
        d.mkdir(parents=True, exist_ok=True)
    txs, sts, fx_rows = read_csv(TX_CSV), read_csv(ST_CSV), read_csv(FX_CSV)
    reimport = "--reimport" in sys.argv
    if reimport:
        # parse every file again from scratch (e.g. after a parser fix); manual overrides / notes are kept by row id
        kept = {t["id"]: {k: t.get(k) or "" for k in ("merchant_override", "category_override",
                                                    "subcategory_override", "note")} for t in txs}
        txs, sts = [], []
    if "--build" not in sys.argv:
        print("Importing statements…")
        added = import_new(txs, sts, fresh=reimport)
        if reimport:
            for t in txs:
                t.update({k: v for k, v in kept.get(t["id"], {}).items() if v})
        print(f"  {added} new statement(s)")
        write_csv(TX_CSV, TX_FIELDS, txs)
        write_csv(ST_CSV, ST_FIELDS, sts)
    fx_rows = ensure_fx(sts, fx_rows)
    fx_rows = derive_fx_from_transfers(sts, txs, fx_rows)
    write_csv(FX_CSV, FX_FIELDS, fx_rows)
    print("Building workbook…")
    try:
        rows = build_workbook(txs, sts, fx_rows, load_rules(), load_people())
    except PermissionError:
        # the dashboard is written before the workbook is saved, so it's already up to date
        print(f"  ! {OUT_XLSX.name} is open in Excel, so it wasn't updated (the dashboard was) — close it and run "
              f"again to refresh the Excel ledger.")
        print(f"Done → {OUT_HTML.name}")
        return
    if "--no-recalc" not in sys.argv:
        recalc_with_excel(OUT_XLSX)
    print(f"Done → {OUT_HTML.name} (open it in your browser) + {OUT_XLSX.name} ({len(rows)} transactions)")


if __name__ == "__main__":
    main()
