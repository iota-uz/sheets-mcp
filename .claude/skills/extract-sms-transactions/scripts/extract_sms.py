#!/usr/bin/env python3
"""Extract Tengebank/Uzumbank transaction SMS from macOS Messages chat.db.

Output: JSON list to stdout, one record per real transaction.
Filters out: OTP codes, declined/failed, surveys/promos, top-ups, cashback.

Usage:
  extract_sms.py --since 2026-04-01            # all banks since date
  extract_sms.py --since 2026-04-01 --bank tenge
  extract_sms.py --since 2026-04-01 --bank uzum
"""
import argparse, json, re, sqlite3, sys, os
import typedstream

DB = os.path.expanduser("~/Library/Messages/chat.db")

TENGE_SENDERS = ("TENGEBANK", "tengebank")
UZUM_SENDERS = ("UzumBank", "uzumbank", "Uzum_Bank", "UzumBank_uz", "uzumbank_uz")

# Card mask -> account name in sheet
CARD_MAP = {
    "3734": "Tenge Mastercard",
    "9407": "Uzum Visa UZS",
}


def decode_body(text, blob):
    if text:
        return text
    if not blob:
        return None
    try:
        for c in typedstream.unarchive_from_data(blob).contents:
            v = getattr(c, "value", None)
            if v is not None and "String" in type(v).__name__:
                return getattr(v, "value", str(v))
    except Exception:
        return None
    return None


def apple_to_iso(ns):
    # Apple epoch ns -> ISO local
    import datetime
    sec = ns / 1e9 + 978307200
    return datetime.datetime.fromtimestamp(sec).strftime("%Y-%m-%d %H:%M:%S")


# Tenge: "Pokupka 25.35 USD, po karte 518100******3734, 2026-04-30 20:44:59, USA. Dostupno: ..."
TENGE_PURCHASE = re.compile(
    r"^Pokupka\s+([\d\s.,]+?)\s+(USD|EUR|UZS|RUB|SOM)"
    r"(?:\s*\+\s*CASH BACK [\d\s.,]+\s+\w+)?"
    r",\s*po karte\s+\d+\*+(\d{4}),\s*(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})"
)

# Uzum (RU): "Spisanie/Pokupka, karta ****9407: 149.00 USD, AHREFS.COM, SG. Dostupno: ..."
# Uzum (EN): "Withdrawal/Purchase, card *9407: 500,00 USD, GOOGLE*ADS..., IE. Available: ..."
UZUM_PURCHASE = re.compile(
    r"^(?:Spisanie|Pokupka|Withdrawal|Purchase),?\s+(?:karta|card)\s+\*+(\d{4}):\s*"
    r"([\d\s.,]+?)\s+(USD|EUR|UZS|RUB|SOM|sum)"
    r",\s*([^,.]+?)(?:,\s*[A-Z]{2})?\.\s*(?:Dostupno|Available)",
    re.IGNORECASE,
)

# Reject patterns
SKIP_PATTERNS = [
    r"^Otkaz",                       # Tenge declined
    r"^Transaction failed",          # Uzum failed (en)
    r"^Payment failed",
    r"CASH-?BACK",                   # cashback (income)
    r"^Top-?up",                     # incoming
    r"Popolnenie",
    r"Confirm the transfer",         # OTP
    r"Do not share the code",
    r"opros",                        # survey
    r"^Pozhaluysta, proydite",
    r"One-time code",
    r"confirmation code",
    r"^TENGE24",
    r"Vam zachislen",                # cashback ru
]
SKIP_RE = re.compile("|".join(SKIP_PATTERNS), re.IGNORECASE)


def parse_amount(s):
    # "500,00" / "1 342 844,08" / "1342844.08" / "149.00"
    s = s.replace(" ", "").replace(" ", "")
    if "," in s and "." in s:
        # ambiguous; use rightmost as decimal
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        # comma is decimal sep
        s = s.replace(",", ".")
    return float(s)


def normalize_currency(c):
    c = c.upper()
    return "SOM" if c in ("UZS", "SOM", "SUM") else c


def extract(since_iso, banks):
    senders = []
    if "tenge" in banks:
        senders += list(TENGE_SENDERS)
    if "uzum" in banks:
        senders += list(UZUM_SENDERS)

    db = sqlite3.connect(DB)
    cur = db.cursor()
    placeholders = ",".join("?" * len(senders))
    since_apple_ns = int((__import__("time").mktime(__import__("time").strptime(since_iso, "%Y-%m-%d")) - 978307200) * 1e9)
    cur.execute(
        f"""SELECT m.date, h.id, m.text, m.attributedBody
            FROM message m JOIN handle h ON m.handle_id=h.ROWID
            WHERE h.id IN ({placeholders}) AND m.date >= ?
            ORDER BY m.date""",
        (*senders, since_apple_ns),
    )

    out = []
    for date_ns, sender, text, blob in cur.fetchall():
        body = decode_body(text, blob)
        if not body:
            continue
        if SKIP_RE.search(body):
            continue

        rec = None
        if sender in TENGE_SENDERS:
            m = TENGE_PURCHASE.match(body)
            if m:
                amt, cur_, card, d, t = m.groups()
                rec = {
                    "bank": "tenge",
                    "card": card,
                    "account": CARD_MAP.get(card, f"card-{card}"),
                    "amount": parse_amount(amt),
                    "currency": normalize_currency(cur_),
                    "datetime": f"{d} {t}",
                    "date": ".".join(reversed(d.split("-"))),  # ДД.ММ.ГГГГ
                    "merchant": None,
                    "raw": body,
                    "sms_ts": apple_to_iso(date_ns),
                }
        elif sender in UZUM_SENDERS:
            m = UZUM_PURCHASE.match(body)
            if m:
                card, amt, cur_, merch = m.groups()
                rec = {
                    "bank": "uzum",
                    "card": card,
                    "account": CARD_MAP.get(card, f"card-{card}"),
                    "amount": parse_amount(amt),
                    "currency": normalize_currency(cur_),
                    "datetime": apple_to_iso(date_ns),
                    "date": apple_to_iso(date_ns)[8:10] + "." + apple_to_iso(date_ns)[5:7] + "." + apple_to_iso(date_ns)[0:4],
                    "merchant": merch.strip(),
                    "raw": body,
                    "sms_ts": apple_to_iso(date_ns),
                }

        if rec:
            out.append(rec)

    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", required=True, help="YYYY-MM-DD")
    ap.add_argument("--bank", choices=["tenge", "uzum", "all"], default="all")
    args = ap.parse_args()
    banks = ["tenge", "uzum"] if args.bank == "all" else [args.bank]
    print(json.dumps(extract(args.since, banks), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
