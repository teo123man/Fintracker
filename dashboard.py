# -*- coding: utf-8 -*-
"""Harcamalarim.html — a self-contained, offline dashboard (no internet, no libraries).

harcama.py hands over its enriched rows; build_payload() turns them into a compact JSON payload and
write_dashboard() drops it into dashboard_template.html. Only months from the first card statement on are
included — earlier history stays in the Excel ledger (History sheet)."""
from __future__ import annotations

import calendar
import datetime as dt
import json
import re
from collections import defaultdict
from pathlib import Path

TEMPLATE = Path(__file__).with_name("dashboard_template.html")
REVIEW = {"guess": "check", "none": "check", "unknown": "unknown"}
# not spending, but shown: money coming in (the flow chart) and cash advances (the Cards page)
INCLUDED_NON_SPENDING = {"Income", "Other income", "Family support", "Cash advance"}


def _add_months(d: dt.date, n: int) -> dt.date:
    y = d.year + (d.month - 1 + n) // 12
    m = (d.month - 1 + n) % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def _card_no(desc: str) -> str:
    m = re.search(r"\*+\s*(\d{4})", desc) or re.match(r"^(\d{4})-\d{6,}", desc)
    return m.group(1) if m else ""


def _tl(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".") + " ₺"


def build_payload(rows, sts, fx, sources, source_days, complete, trip_ranges, income_plan, budgets, rotation_text,
                  extra_categories=None, rotation_start=None, confirmations=None, exclusions=None):
    card_sts = [s for s in sts if (s.get("kind") or "Card") != "Account"]
    start = min(s["window_start"] for s in card_sts).replace(day=1)
    start_m = start.strftime("%Y-%m")
    last = max(max(r["date"] for r in rows), max(s["cutoff"] for s in sts))
    today = dt.date.today()

    days, d = [], start
    while d <= last:
        days.append({"d": d.isoformat(), "c": int(complete(d)), "t": int(any(a <= d <= b for a, b in trip_ranges))})
        d += dt.timedelta(days=1)

    tx = []
    for r in rows:
        t = r["t"]
        if r["month"] < start_m or (not r["counts"] and t["type"] not in INCLUDED_NON_SPENDING):
            continue
        on_card = t["type"] == "Purchase" and t["section"] != "Account" and r["pay"]
        tx.append({"id": t["id"], "d": r["date"].isoformat(), "m": r["month"], "card": t["card"],
                   "merch": r["merchant"], "cat": r["cat"], "sub": r["sub"], "type": t["type"],
                   "amt": r["try_amt"], "sp": r["counts"], "ch": r["channel"], "tag": r["tag"],
                   "fl": (r["pay"] - r["date"]).days if on_card else None, "rot": int(bool(r["rotation"])),
                   "rev": REVIEW.get(r["conf"], ""), "desc": t["description"],
                   "bankCat": t.get("bank_category") or ""})

    # statements (grouped per card & cutoff: Garanti's TL/EUR/USD parts are one bill) + what was paid for them
    pays = [r for r in rows if r["t"]["section"] == "Account" and r["t"]["type"] == "Payment"]
    groups = defaultdict(list)
    for s in card_sts:
        groups[(s["card"], s["cutoff"])].append(s)
    statements = []
    for (card, cutoff), g in sorted(groups.items(), key=lambda kv: (kv[0][1], kv[0][0])):
        due = g[0]["payment_date"]
        # payments after this cutoff, up to the card's next cutoff (a later payment settles the next bill)
        nxt = min((c for (cd, c) in groups if cd == card and c > cutoff), default=cutoff + dt.timedelta(days=35))
        paid_rows = [p for p in pays if card[-4:] in p["t"]["description"] and cutoff < p["date"] <= nxt]
        paid_on = max((p["date"] for p in paid_rows), default=None)
        tl_part = sum(float(s["balance_due"]) for s in g if s["currency"] == "TRY")
        paid = sum(p["try_amt"] for p in paid_rows)
        interest = sum(r["try_amt"] for r in rows if r["t"]["card"] == card and r["sub"] == "Interest & tax"
                       and due < r["date"] <= cutoff + dt.timedelta(days=40))
        bal_try = sum(float(s["balance_due"]) * (1.0 if s["currency"] == "TRY"
                                                 else fx.get(f"{s['currency']}|{s['payment_date']}", 0.0)) for s in g)
        statements.append({"card": card, "cutoff": cutoff.isoformat(), "due": due.isoformat(),
                           "balance": " + ".join(f"{float(s['balance_due']):,.2f} {s['currency']}"
                                                 for s in sorted(g, key=lambda s: s["currency"] != "TRY")
                                                 if s["currency"] == "TRY" or float(s["balance_due"])),
                           "balTry": round(bal_try, 2), "paid": round(paid, 2),
                           "paidOn": paid_on.isoformat() if paid_on else None,
                           "tlUnpaid": round(max(0.0, tl_part - paid), 2) if paid_rows else None,
                           "interest": round(interest, 2)})

    # cash advances: one line per advance (its installments arrive on successive statements)
    adv_groups = defaultdict(list)
    for r in rows:
        if r["t"]["type"] == "Cash advance":
            m = re.match(r"(\d+)/(\d+) of ([\d.]+)", r["t"]["installment"] or "")
            k, n, tot = (int(m.group(1)), int(m.group(2)), float(m.group(3))) if m else (1, 1, r["amt"])
            adv_groups[(r["t"]["card"], n, round(tot, 2))].append((k, r))
    advances = []
    for (card, n, tot), lst in adv_groups.items():
        k, r = max(lst, key=lambda kr: kr[0])
        taken = _add_months(r["date"], -(k - 1))
        principal = next((-y["amt"] for y in rows if y["t"]["type"] == "Cash advance received"
                          and abs((y["date"] - taken).days) <= 2), None)
        advances.append({"card": card, "taken": taken.isoformat(), "k": k, "n": n, "inst": r["try_amt"],
                         "total": tot, "remaining": round(r["try_amt"] * (n - k), 2), "principal": principal,
                         "cost": round(tot - principal, 2) if principal else None})

    # things that make the numbers incomplete
    alerts = []
    known = {s["card"][-4:] for s in card_sts}
    unknown_paid = defaultdict(list)
    for p in pays:
        no = _card_no(p["t"]["description"])
        if no and no not in known:
            unknown_paid[no].append(p)
    for no, lst in unknown_paid.items():
        last_day = max(p["date"] for p in lst)
        if (today - last_day).days > 60:
            continue
        last_amt = sum(p["try_amt"] for p in lst if p["date"] == last_day)  # a day can have several payment rows
        recent = sum(p["try_amt"] for p in lst if (today - p["date"]).days <= 92)
        alerts.append({"level": "critical", "title": f"Card •{no} is missing",
                       "text": f"You pay it from your {lst[0]['t']['bank']} account ({_tl(last_amt)} on "
                               f"{last_day:%d.%m.%Y}, {_tl(recent)} in the last 3 months), but none of its "
                               f"statements are imported — its spending is in none of the numbers here."})
    missing = set()
    for s in card_sts:
        if s["currency"] != "TRY":
            continue
        for r in rows:
            if r["t"]["statement_id"] == s["statement_id"] and r["t"]["type"] == "FX balance transfer":
                m = re.search(r"\b(EUR|USD|GBP)\b", r["t"]["description"].upper())
                if m and not any(x["card"] == s["card"] and x["currency"] == m.group(1) and x["cutoff"] < s["cutoff"]
                                 for x in card_sts):
                    missing.add((s["card"], m.group(1), s["cutoff"]))
    for card, cur, cutoff in sorted(missing):
        alerts.append({"level": "warning", "title": f"{card}: a {cur} statement is missing",
                       "text": f"The statement of {cutoff:%d.%m.%Y} carries an unpaid {cur} balance from the "
                               f"previous {cur} statement, which isn't imported — its purchases are missing."})

    cats = defaultdict(set)
    for r in rows:
        if r["counts"]:
            cats[r["cat"]].add(r["sub"])
    for c, subs in (extra_categories or {}).items():
        cats[c] |= set(subs)

    # overlapping account exports are de-duplicated row by row; this is the safety net in case a bank rewrote a row
    # between two exports (same account & amount, ≤2 days apart, coming from different files)
    by_amt = defaultdict(list)
    for r in rows:
        t = r["t"]
        if t["section"] == "Account" and r["month"] >= start_m:
            by_amt[(t["card"], t["amount"])].append(r)
    suspects = set()
    for lst in by_amt.values():
        for i, a in enumerate(lst):
            for b in lst[i + 1:]:
                if (a["t"]["source_file"] != b["t"]["source_file"] and abs((a["date"] - b["date"]).days) <= 2
                        and a["t"]["description"][:12] == b["t"]["description"][:12]):
                    suspects.add(b["t"]["id"])
    if suspects:
        alerts.append({"level": "warning", "title": f"{len(suspects)} possible duplicate row(s)",
                       "text": "Two overlapping account exports contain rows with the same account, amount and "
                               "description a day or two apart. Check them in Transactions — if they're the same "
                               "transaction, one can be removed from data/transactions.csv."})

    salary = next((float(i["amount"]) for i in income_plan if "salary" in i["item"].lower()), 0.0)
    return {
        "meta": {"built": today.isoformat(), "start": start_m, "rotation": rotation_text,
                 "rotationStart": rotation_start.isoformat() if rotation_start else None,
                 "noActivity": confirmations or [],
                 "coverageExclusions": exclusions or [],
                 "sources": [{"name": n, "last": max(source_days[n]).isoformat(),
                              "account": "account" in n} for n in sources],
                 "plan": [{"item": i["item"], "amount": float(i["amount"])} for i in income_plan],
                 "salary": salary, "budgets": budgets, "alerts": alerts,
                 "categories": {c: sorted(s) for c, s in sorted(cats.items())}},
        "days": days, "tx": tx, "statements": statements, "advances": advances,
    }


def write_dashboard(out_path: Path, payload: dict):
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = TEMPLATE.read_text(encoding="utf-8").replace("/*__DATA__*/null", data)
    out_path.write_text(html, encoding="utf-8")
