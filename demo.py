# -*- coding: utf-8 -*-
"""A fake 'home' full of made-up statements, for developing features without touching real data.

    python demo.py          # (re)creates demo/ and builds demo/Harcamalarim.html
    python app.py --demo    # the local app on top of it

Everything here is invented: two cards (Garanti •1111 with a TL and a EUR statement, QNB •2222), two accounts,
a salary, friends, a trip, a cash advance and an unknown card, so every panel of the dashboard has something to show.
Dates follow today's date, so the demo never goes stale."""
from __future__ import annotations

import calendar
import csv
import datetime as dt
import os
import random
import shutil
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
CODE = Path(__file__).resolve().parent
HOME = CODE / "demo"
R = random.Random(7)
HOLDER = "ADA DEMIR"

TX_FIELDS = ["id", "statement_id", "bank", "card", "date", "description", "bank_category", "section", "type",
             "currency", "amount", "orig_amount", "orig_currency", "installment", "bonus", "payment_date",
             "source_file", "merchant_override", "category_override", "subcategory_override", "note",
             "ext_id", "counterparty", "memo"]
ST_FIELDS = ["statement_id", "bank", "card", "currency", "ekstre_no", "period_start", "cutoff", "payment_date",
             "prev_balance", "spending", "fees_interest", "payments_credits", "balance_due", "min_payment",
             "source_file", "kind"]

# description, category rule (pattern, merchant, category, subcategory, confidence) or None, amount range, weight
SHOPS = [
    ("MIGROS SANAL MARKET ISTANBUL", ("MIGROS", "Migros", "Groceries", "Supermarket", "high"), (250, 1900), 7),
    ("A101 KADIKOY", ("A101", "A101", "Groceries", "Supermarket", "high"), (60, 520), 6),
    ("GETIR ISTANBUL", ("GETIR", "Getir", "Groceries", "Quick delivery", "high"), (140, 760), 4),
    ("STARBUCKS MODA", ("STARBUCKS", "Starbucks", "Food & Dining", "Coffee", "high"), (115, 270), 7),
    ("KRONOTROP CIHANGIR", ("KRONOTROP", "Kronotrop", "Food & Dining", "Coffee", "high"), (110, 240), 3),
    ("YEMEKSEPETI.COM", ("YEMEKSEPETI", "Yemeksepeti", "Food & Dining", "Delivery", "high"), (260, 950), 6),
    ("CIYA SOFRASI KADIKOY", ("CIYA", "Çiya Sofrası", "Food & Dining", "Restaurant", "high"), (420, 1600), 2),
    ("BURGER KING BOSTANCI", ("BURGER KING", "Burger King", "Food & Dining", "Fast food", "high"), (240, 520), 3),
    ("KOFTECI YUSUF ATASEHIR", ("KOFTECI YUSUF", "Köfteci Yusuf", "Food & Dining", "Restaurant", "high"),
     (300, 780), 2),
    ("ISTANBULKART DOLUM", ("ISTANBULKART", "İstanbulkart", "Transport", "Public transport", "high"), (200, 600), 5),
    ("BITAKSI", ("BITAKSI", "BiTaksi", "Transport", "Taxi", "high"), (160, 520), 3),
    ("MARTI TEKNOLOJI", ("MARTI", "Martı", "Transport", "Scooter", "guess"), (45, 140), 3),
    ("TRENDYOL.COM", ("TRENDYOL", "Trendyol", "Shopping", "Online", "high"), (280, 2600), 3),
    ("ZARA KADIKOY", ("ZARA", "Zara", "Shopping", "Clothes", "high"), (900, 3200), 1),
    ("DECATHLON ATASEHIR", ("DECATHLON", "Decathlon", "Shopping", "Sports", "high"), (400, 2400), 1),
    ("ECZANE SIFA", ("ECZANE", "Pharmacy", "Health", "Pharmacy", "high"), (90, 650), 1),
    ("BILET BLOK KONSER", ("BILET", "Biletix / Bubilet", "Entertainment", "Events", "guess"), (600, 1800), 1),
    ("PAYTR *GSH TICARET 34", None, (180, 1400), 1),       # unmatched → shows up in "to review"
    ("IYZICO *NOVA DIJITAL", None, (90, 450), 1),
]
MONTHLY = [  # day of month, description, rule, amount
    (5, "NETFLIX.COM", ("NETFLIX", "Netflix", "Subscriptions & Digital", "Streaming", "high"), 229.99),
    (12, "SPOTIFY P3A8F", ("SPOTIFY", "Spotify", "Subscriptions & Digital", "Music", "high"), 59.99),
    (8, "APPLE.COM/BILL", ("APPLE\\.COM", "Apple (iCloud)", "Subscriptions & Digital", "Cloud", "high"), 39.99),
    (20, "TURKCELL FATURA", ("TURKCELL", "Turkcell", "Bills & Utilities", "Phone", "high"), 459.00),
    (22, "IGDAS DOGALGAZ", ("IGDAS", "İGDAŞ", "Bills & Utilities", "Gas", "high"), None),
    (27, "SPOR SALONU MACFIT", ("MACFIT", "MACFit", "Health", "Gym", "high"), 1450.00),
]
TRIP = [("PASTEIS DE BELEM LISBOA", ("PASTEIS", "Pastéis de Belém", "Travel", "Food abroad", "high"), (6, 28)),
        ("UBER *TRIP LISBOA", ("UBER", "Uber", "Travel", "Local transport", "high"), (7, 22)),
        ("PINGO DOCE LISBOA", ("PINGO DOCE", "Pingo Doce", "Travel", "Groceries abroad", "high"), (9, 45)),
        ("TIME OUT MARKET LISBOA", ("TIME OUT", "Time Out Market", "Travel", "Food abroad", "high"), (18, 60))]
FRIENDS = ["DENIZ YILMAZ", "ECE KAYA", "MERT OZTURK", "SELIN ARSLAN"]


def add_months(d, n):
    y, m = d.year + (d.month - 1 + n) // 12, (d.month - 1 + n) % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def month_end(d):
    return d.replace(day=calendar.monthrange(d.year, d.month)[1])


def days(a, b):
    while a <= b:
        yield a
        a += dt.timedelta(days=1)


def write(path, fields, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})


def main():
    today = dt.date.today()
    # the last three Garanti cutoffs (15th) and QNB cutoffs (month end) at least 3 days in the past
    g_cut = [c for c in (add_months(today.replace(day=15), -i) for i in range(0, 5)) if c <= today - dt.timedelta(3)][:3]
    q_cut = [c for c in (month_end(add_months(today.replace(day=1), -i)) for i in range(0, 5))
             if c <= today - dt.timedelta(3)][:3]
    g_cut.reverse(), q_cut.reverse()
    trip_start = g_cut[1] - dt.timedelta(days=26)
    trip_end = trip_start + dt.timedelta(days=4)

    txs, sts, rules = [], [], {}
    seq = {}

    def add_tx(st, d, desc, amount, typ="Purchase", **kw):
        seq[st["statement_id"]] = seq.get(st["statement_id"], 0) + 1
        row = dict(id=f"{st['statement_id']}-{seq[st['statement_id']]:03d}", statement_id=st["statement_id"],
                   bank=st["bank"], card=st["card"], date=d.isoformat(), description=desc,
                   section=kw.pop("section", "Domestic"), type=typ, currency=st["currency"],
                   amount=f"{amount:.2f}", payment_date=st["payment_date"], source_file=st["source_file"])
        row.update(kw)
        txs.append(row)
        return row

    def rule(r):
        if r:
            rules[r[0]] = dict(zip(["pattern", "merchant", "category", "subcategory", "confidence"], r))

    def card_card(bank, card, cuts, first_start, owner_days, cur="TRY", tag=""):
        prev_bal, out = 0.0, []
        for i, cut in enumerate(cuts):
            start = cuts[i - 1] + dt.timedelta(1) if i else first_start
            due = cut + dt.timedelta(days=10)
            sid = f"{tag or bank.upper()}-{card[-4:]}-{cur}-{cut}"
            st = dict(statement_id=sid, bank=bank, card=card, currency=cur, ekstre_no=f"D{R.randint(10**6, 10**7)}",
                      period_start=start.isoformat() if bank.startswith("Garanti") else "",
                      cutoff=cut.isoformat(), payment_date=due.isoformat(), source_file=f"demo_{sid}.pdf")
            n0 = len(txs)
            if prev_bal and cur == "TRY":
                add_tx(st, cuts[i - 1] + dt.timedelta(days=10), "ODEMENIZ ICIN TESEKKUR EDERIZ", -prev_bal, "Payment")
            if cur == "TRY":
                for d in days(start, cut):
                    if not owner_days(d):
                        continue
                    for _ in range(R.choices([0, 1, 2, 3, 4], [2, 4, 4, 2, 1])[0]):
                        desc, r, (lo, hi), _w = R.choices(SHOPS, [s[3] for s in SHOPS])[0]
                        rule(r)
                        add_tx(st, d, desc, round(R.uniform(lo, hi), 2))
                    for day, desc, r, amt in MONTHLY:
                        if d.day == day:
                            rule(r)
                            add_tx(st, d, desc, amt or round(R.uniform(280, 900), 2))
                if R.random() < .5:
                    d = start + dt.timedelta(days=R.randint(2, 20))
                    add_tx(st, d, "TRENDYOL.COM IADE", -round(R.uniform(150, 600), 2), "Refund")
            rows = txs[n0:]
            spend = sum(float(r["amount"]) for r in rows if float(r["amount"]) > 0)
            pays = sum(float(r["amount"]) for r in rows if float(r["amount"]) < 0)
            bal = round(prev_bal + spend + pays, 2)
            st.update(prev_balance=f"{prev_bal:.2f}", spending=f"{spend:.2f}", fees_interest="0.00",
                      payments_credits=f"{pays:.2f}", balance_due=f"{bal:.2f}", min_payment=f"{bal * .4:.2f}")
            sts.append(st)
            out.append(st)
            prev_bal = bal
        return out

    # rotation: QNB from the 2nd to the 14th, Garanti from the 16th to the 1st (the 15th: either)
    garanti = card_card("Garanti BBVA", "Garanti Bonus •1111", g_cut, g_cut[0] - dt.timedelta(days=29),
                        lambda d: (d.day >= 15 or d.day == 1 or R.random() < .06) and not trip_start <= d <= trip_end)
    qnb = card_card("QNB", "QNB Fix •2222", q_cut, q_cut[0].replace(day=1),
                    lambda d: 2 <= d.day <= 14 and not trip_start <= d <= trip_end)

    # a EUR statement for the trip (same cutoff as the TL one it belongs to)
    cut = g_cut[1]
    eur = dict(statement_id=f"GARANTI-1111-EUR-{cut}", bank="Garanti BBVA", card="Garanti Bonus •1111", currency="EUR",
               ekstre_no="D-EUR", period_start=(g_cut[0] + dt.timedelta(1)).isoformat(), cutoff=cut.isoformat(),
               payment_date=(cut + dt.timedelta(days=10)).isoformat(), source_file=f"demo_GARANTI-1111-EUR-{cut}.pdf")
    n0 = len(txs)
    add_tx(eur, trip_start, "HOTEL ALFAMA LISBOA", 318.40, section="Foreign")
    rule(("HOTEL ALFAMA", "Hotel Alfama", "Travel", "Accommodation", "high"))
    for d in days(trip_start, trip_end):
        for desc, r, (lo, hi) in R.sample(TRIP, 3):
            rule(r)
            add_tx(eur, d, desc, round(R.uniform(lo, hi), 2), section="Foreign")
    eur_total = round(sum(float(r["amount"]) for r in txs[n0:]), 2)
    eur.update(prev_balance="0.00", spending=f"{eur_total:.2f}", fees_interest="0.00", payments_credits="0.00",
               balance_due=f"{eur_total:.2f}", min_payment="0.00")
    sts.append(eur)
    eur_rate = 48.3721

    # a 3-installment cash advance on QNB, and a card fee
    adv_day = q_cut[0] - dt.timedelta(days=1)
    for i, st in enumerate(qnb):
        add_tx(st, dt.date.fromisoformat(st["cutoff"]), "NAKIT AVANS TAKSIT", 3150.00, "Cash advance",
               installment=f"{i + 1}/3 of 9450.00")
    add_tx(garanti[-1], dt.date.fromisoformat(garanti[-1]["cutoff"]), "YILLIK KART UCRETI", 650.00, "Fee")
    for st in garanti + qnb:  # the extra rows change the balance
        extra = sum(float(r["amount"]) for r in txs if r["statement_id"] == st["statement_id"]
                    and r["type"] in ("Cash advance", "Fee"))
        st["balance_due"] = f"{float(st['balance_due']) + extra:.2f}"

    # ── accounts ──
    acc_start, acc_end = q_cut[0].replace(day=1), today - dt.timedelta(days=2)
    acct = {}
    for bank, card, tag in (("Garanti BBVA", "Garanti account •3333", "GARANTI-ACCT"),
                            ("QNB", "QNB account •4444", "QNB-ACCT")):
        sid = f"{tag}-{card[-4:]}-{acc_start}-{acc_end}"
        acct[bank] = dict(statement_id=sid, bank=bank, card=card, currency="TRY", ekstre_no="Hesap demo",
                          period_start=acc_start.isoformat(), cutoff=acc_end.isoformat(),
                          payment_date=acc_end.isoformat(), source_file=f"demo_{tag.lower()}.xlsx", kind="Account",
                          prev_balance="0.00", spending="0.00", fees_interest="0.00", payments_credits="0.00",
                          balance_due="0.00")
        sts.append(acct[bank])
    ext = iter(range(10**6, 10**7))

    def acc(bank, d, desc, amount, typ, **kw):
        if d > acc_end:
            return
        st = acct[bank]
        add_tx(dict(st, payment_date=d.isoformat()), d, desc, amount, typ, section="Account",
               ext_id=f"demo-{next(ext)}", **kw)

    tr_months = ["OCAK", "SUBAT", "MART", "NISAN", "MAYIS", "HAZIRAN", "TEMMUZ", "AGUSTOS", "EYLUL", "EKIM",
                 "KASIM", "ARALIK"]
    m = acc_start
    while m <= acc_end:
        prev = add_months(m, -1)
        acc("QNB", m.replace(day=4), f"{prev.year}-{tr_months[prev.month - 1]} MAAS ODEMESI", -60000.00, "Income")
        acc("QNB", m.replace(day=20), f"Transfer Islemleri - Alici : {HOLDER.title()}", 25000.00, "Own transfer",
            counterparty=HOLDER)
        acc("Garanti BBVA", m.replace(day=20), f"{HOLDER} GOND: FAST", -25000.00, "Own transfer", counterparty=HOLDER)
        acc("Garanti BBVA", m.replace(day=12), "KREDI KARTI ODEME ****9999", round(R.uniform(2200, 4200), 2), "Payment")
        acc("Garanti BBVA", m.replace(day=R.randint(3, 26)), "ATM PARA CEKME KADIKOY", 1000.00, "Cash withdrawal")
        for _ in range(R.randint(2, 4)):
            who = R.choice(FRIENDS)
            acc("Garanti BBVA", m.replace(day=R.randint(1, 27)), f"{who} -FAST", round(R.uniform(250, 1500), -1),
                "Transfer out", counterparty=who)
        for _ in range(R.randint(1, 3)):
            who = R.choice(FRIENDS)
            acc("Garanti BBVA", m.replace(day=R.randint(1, 27)), f"{who} -FAST", -round(R.uniform(200, 1200), -1),
                "Reimbursement", counterparty=who)
        acc("Garanti BBVA", m.replace(day=R.randint(1, 27)), "EFT UCRETI", 8.50, "Fee")
        m = add_months(m, 1)
    acc("QNB", adv_day, "NAKIT AVANS ****2222 MOBIL-3", -9000.00, "Cash advance received")
    for st in garanti:
        extra = sum(float(s["balance_due"]) * eur_rate for s in sts if s["currency"] == "EUR"
                    and s["cutoff"] == st["cutoff"])
        acc("Garanti BBVA", dt.date.fromisoformat(st["payment_date"]), "KK OTOMATIK ODEME ****1111",
            round(float(st["balance_due"]) + extra, 2), "Payment")
    for st in qnb:
        acc("QNB", dt.date.fromisoformat(st["payment_date"]), "KREDI KARTI ODEME ****2222",
            float(st["balance_due"]), "Payment")

    # ── write the home ──
    if HOME.exists():
        shutil.rmtree(HOME)
    data = HOME / "data"
    write(data / "transactions.csv", TX_FIELDS, txs)
    write(data / "statements.csv", ST_FIELDS, sts)
    write(data / "fx_rates.csv", ["key", "currency", "payment_date", "rate", "rate_date", "source", "status"],
          [dict(key=f"EUR|{eur['payment_date']}", currency="EUR", payment_date=eur["payment_date"],
                rate=f"{eur_rate}", rate_date=eur["payment_date"], source="demo", status="derived")])
    write(data / "category_rules.csv", ["pattern", "merchant", "category", "subcategory", "confidence", "note"],
          list(rules.values()))
    write(data / "tags.csv", ["tag", "start", "end", "currency", "card", "reimbursements_until", "description"],
          [dict(tag="Lisbon trip", start=trip_start.isoformat(), end=trip_end.isoformat(), description="demo")])
    write(data / "people.csv", ["pattern", "label", "category", "subcategory", "note"], [])
    write(data / "income_plan.csv", ["item", "amount", "from", "note"],
          [dict(item="Salary", amount="60000"), dict(item="Meal card", amount="4000")])
    write(data / "budgets.csv", ["category", "monthly_budget"],
          [dict(category=c, monthly_budget=b) for c, b in (("Food & Dining", 9000), ("Groceries", 8000),
                                                           ("Transport", 2500), ("Shopping", 5000),
                                                           ("Subscriptions & Digital", 500))])
    (HOME / "Belgeler").mkdir(parents=True, exist_ok=True)
    print(f"demo home → {HOME} ({len(txs)} transactions, {len(sts)} statements)")
    env = dict(os.environ, HARCAMA_HOME=str(HOME))
    subprocess.run([sys.executable, str(CODE / "harcama.py"), "--build", "--no-recalc"], env=env, check=True)


if __name__ == "__main__":
    main()
