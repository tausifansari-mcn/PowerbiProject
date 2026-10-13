#!/usr/bin/env python3
"""
Downloads Rohan's daily "GNC Dashboard Data" mail (rohan@teammas.in) - two workbooks:
  Order <Mon>.xlsx           (Sheet1 = sale raw, Sheet2 = Zecpe)    -> Order.xlsx
  Allocation <dd.mm.yyyy>.xlsx (allocation of the report date)      -> Allocation.xlsx
into --out. The mail whose Allocation file is for --date (default yesterday) is taken; if the
names carry no date, the newest mail received the day after --date is used.
Gmail login: the same email.env.txt used by the BVO export / Clovia (read-only IMAP).
"""
from __future__ import annotations
import argparse, email, imaplib, re, sys
from datetime import date, datetime, timedelta
from email.header import decode_header, make_header
from pathlib import Path

SENDER = "rohan@teammas.in"
SUBJECT = "GNC Dashboard Data"
ENV_FILES = [Path("D:\\MIS\\Bella Vitta\\Oct'26\\automate") / n for n in ("gmail.env", "email.env", "email.env.txt")]


def load_env() -> dict[str, str]:
    p = next((p for p in ENV_FILES if p.exists()), None)
    if not p:
        sys.exit("No Gmail login file found (email.env.txt)")
    env = {}
    for line in p.read_text(encoding="utf-8").splitlines():
        if "=" in line and not line.strip().startswith("#"):
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').replace(" ", "")
    return env


def text(h) -> str:
    return str(make_header(decode_header(h or "")))


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="report date YYYY-MM-DD (default yesterday)")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    d = datetime.strptime(a.date, "%Y-%m-%d").date() if a.date else date.today() - timedelta(days=1)
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    env = load_env()
    im = imaplib.IMAP4_SSL("imap.gmail.com")
    im.login(env["GMAIL_USER"], env["GMAIL_APP_PASSWORD"])
    try:
        im.select('"[Gmail]/All Mail"', readonly=True)
        typ, ids = im.search(None, f'(FROM "{SENDER}" SUBJECT "{SUBJECT}" SINCE {d:%d-%b-%Y} BEFORE {d + timedelta(days=3):%d-%b-%Y})')
        cands = []
        for mid in ids[0].split():
            m = email.message_from_bytes(im.fetch(mid, "(BODY.PEEK[])")[1][0][1])
            atts = {text(p.get_filename()): p.get_payload(decode=True) for p in m.walk() if p.get_filename()}
            order = next((n for n in atts if n.lower().startswith("order") and n.lower().endswith(".xlsx")), None)
            if not order:
                continue
            when = email.utils.parsedate_to_datetime(m.get("Date"))
            # one mail can carry several allocations (Monday: "Allocation 11.10.2026-.xlsx" + "ALL Allocation 10.10.2026.csv")
            for alloc in (n for n in atts if "allocation" in n.lower() and n.lower().endswith((".xlsx", ".csv"))):
                mdate = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", alloc)
                adate = date(int(mdate[3]), int(mdate[2]), int(mdate[1])) if mdate else None
                cands.append((adate == d, adate is None and when.date() == d + timedelta(days=1), when, order, alloc, atts, text(m.get("Subject"))))
        good = [c for c in cands if c[0]] or [c for c in cands if c[1]]
        if not good:
            sys.exit(f"no '{SUBJECT}' mail from {SENDER} with the Order + Allocation workbooks for {d:%d-%b-%Y} yet")
        c = max(good, key=lambda c: c[2])
        for n in ("Order.xlsx", "Allocation.xlsx", "Allocation.csv"):
            (out / n).unlink(missing_ok=True)
        aname = "Allocation.csv" if c[4].lower().endswith(".csv") else "Allocation.xlsx"
        (out / "Order.xlsx").write_bytes(c[5][c[3]])
        (out / aname).write_bytes(c[5][c[4]])
        print(f"Mail: {c[6]} ({c[2]:%d-%b %H:%M}) | {c[3]} -> Order.xlsx | {c[4]} -> {aname}")
    finally:
        im.logout()


if __name__ == "__main__":
    main()
