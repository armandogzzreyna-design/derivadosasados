from __future__ import annotations

import re
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import pdfplumber

from .excel_utils import parse_money


REF_DATA = [
    ("SCOTIA", "SIEFORE INVERCAP 95", "1008375", "BCSCOT95F"),
    ("SCOTIA", "SIEFORE INVERCAP 60", "1008793", "BCSCOT60F"),
    ("SCOTIA", "SIEFORE INVERCAP 65", "10081002", "BCSCOT65F"),
    ("SCOTIA", "SIEFORE INVERCAP 70", "10024618", "BCSCOT70F"),
    ("SCOTIA", "SIEFORE INVERCAP 75", "1001717", "BCSCOT75F"),
    ("SCOTIA", "SIEFORE INVERCAP 80", "10086617", "BCSCOT80F"),
    ("SCOTIA", "SIEFORE INVERCAP 85", "1003533", "BCSCOT85F"),
    ("SCOTIA", "SIEFORE INVERCAP 90", "10091375", "BCSCOT90F"),
    ("SCOTIA", "SIEFORE INVERCAP IN", "10096697", "BCSCOTINF"),
    ("SCOTIA", "SIEFORE INVERCAP 60", "8793", "BCSCOT60"),
    ("SCOTIA", "SIEFORE INVERCAP 65", "81002", "BCSCOT65"),
    ("SCOTIA", "SIEFORE INVERCAP 70", "24618", "BCSCOT70"),
    ("SCOTIA", "SIEFORE INVERCAP 75", "1717", "BCSCOT75"),
    ("SCOTIA", "SIEFORE INVERCAP 80", "86617", "BCSCOT80"),
    ("BCSANT", "SIEFORE INVERCAP 95", "08375SIC1C", "BCSANT95F"),
    ("BCSANT", "SIEFORE INVERCAP 60", "08793SIC2C", "BCSANT60F"),
    ("BCSANT", "SIEFORE INVERCAP 65", "81002SIC6C", "BCSANT65F"),
    ("BCSANT", "SIEFORE INVERCAP 70", "24618SIC7C", "BCSANT70F"),
    ("BCSANT", "SIEFORE INVERCAP 75", "01717SIC3C", "BCSANT75F"),
    ("BCSANT", "SIEFORE INVERCAP 80", "86617SIC8C", "BCSANT80F"),
    ("BCSANT", "SIEFORE INVERCAP 85", "03533SIC4C", "BCSANT85F"),
    ("BCSANT", "SIEFORE INVERCAP 90", "91375SIC9C", "BCSANT90F"),
    ("BCSANT", "SIEFORE INVERCAP IN", "96697SICBC", "BCSANTINF"),
    ("BCSANT", "SIEFORE INVERCAP 65", "810022IC6I", "BCSANT65"),
    ("BCSANT", "SIEFORE INVERCAP 60", "087932IC2I", "BCSANT60"),
    ("BCSANT", "SIEFORE INVERCAP 70", "246182IC7I", "BCSANT70"),
    ("BCSANT", "SIEFORE INVERCAP 75", "017172IC3I", "BCSANT75"),
    ("BCSANT", "SIEFORE INVERCAP 80", "866172IC8I", "BCSANT80"),
    ("BCSANT", "SIEFORE INVERCAP 85", "035332IC4I", "BCSANT85"),
    ("BCSANT", "SIEFORE INVERCAP 90", "913752IC9I", "BCSANT90"),
    ("BCSANT", "SIEFORE INVERCAP IN", "966972ICBI", "BCSANTIN"),
    ("BCSANT", "SIEFORE INVERCAP 95", "083752IC1I", "BCSANT95"),
    ("SANTANDER CME", "SIEFORE INVERCAP BP", "UKN99", "BSANCME10"),
    ("SANTANDER CME", "SIEFORE INVERCAP 95", "UKJ99", "BSANCME195"),
    ("SANTANDER CME", "SIEFORE INVERCAP 60", "UKK99", "BSANCME160"),
    ("SANTANDER CME", "SIEFORE INVERCAP 65", "VLM99", "BSANCME165"),
    ("SANTANDER CME", "SIEFORE INVERCAP 70", "VLN99", "BSANCME170"),
    ("SANTANDER CME", "SIEFORE INVERCAP 75", "UKL99", "BSANCME175"),
    ("SANTANDER CME", "SIEFORE INVERCAP 80", "VLO99", "BSANCME180"),
    ("SANTANDER CME", "SIEFORE INVERCAP 85", "UKM99", "BSANCME185"),
    ("SANTANDER CME", "SIEFORE INVERCAP 90", "VLP99", "BSANCME190"),
    ("SANTANDER CME", "SIEFORE INVERCAP IN", "VLQ99", "BSANCME1IN"),
    ("GOLDMAN", "SIEFORE INVERCAP BP", "191086", ""),
    ("GOLDMAN", "SIEFORE INVERCAP 95", "191089", "GOLDMAN95"),
    ("GOLDMAN", "SIEFORE INVERCAP 60", "191091", "GOLDMAN60"),
    ("GOLDMAN", "SIEFORE INVERCAP 65", "105729", "GOLDMAN65"),
    ("GOLDMAN", "SIEFORE INVERCAP 70", "105732", "GOLDMAN70"),
    ("GOLDMAN", "SIEFORE INVERCAP 75", "191087", "GOLDMAN75"),
    ("GOLDMAN", "SIEFORE INVERCAP 80", "105733", "GOLDMAN80"),
    ("GOLDMAN", "SIEFORE INVERCAP 85", "191090", "GOLDMAN85"),
    ("GOLDMAN", "SIEFORE INVERCAP 90", "105735", "GOLDMAN90"),
    ("GOLDMAN", "SIEFORE INVERCAP IN", "105736", "GOLDMANIN"),
]


def last_business_day(day: date) -> date:
    return day - timedelta(days={0: 3, 5: 1, 6: 2}.get(day.weekday(), 0))


def _pages(pdf_path: Path) -> list[str]:
    with pdfplumber.open(pdf_path) as pdf:
        return [page.extract_text() or "" for page in pdf.pages]


def _normal_number(text: str) -> float:
    text = re.sub(r"(?<=\d) (?=\d)", "", str(text))
    return parse_money(text)


def parse_goldman_cme(paths: list[Path]) -> list[dict]:
    rows = []
    for path in paths:
        for text in _pages(path):
            if not re.search(r"JOURNAL ENTRIES", text, re.IGNORECASE):
                continue
            if re.search(r"NO ACTIVITY", text, re.IGNORECASE) and not re.search(r"FUNDS\s+(PAID|RECEIVED)", text, re.IGNORECASE):
                continue
            if "PROCESSING" in text.upper():
                continue
            siefore = re.search(r"SIEFORE INVERCAP BASICA\s+([\w][\w\-]*)", text, re.IGNORECASE)
            account = re.search(r"ACCOUNT NUMBER[:\s]+([\d\s]+)", text, re.IGNORECASE)
            received = bool(re.search(r"FUNDS\s+RECEIVED", text, re.IGNORECASE))
            paid = bool(re.search(r"FUNDS\s+PAID", text, re.IGNORECASE))
            amount = re.search(r"USD\s+([\(]?[\d,\.]+[\)]?)", text)
            if not (siefore and amount and (received or paid)):
                continue
            monto = abs(_normal_number(amount.group(1)))
            if received:
                monto = -monto
            rows.append({
                "Counterparty": "GOLDMAN",
                "siefore_raw": siefore.group(0).strip().upper(),
                "account_raw": re.sub(r"\s+", "", account.group(1)).split("\n")[0].strip() if account else "",
                "Currency": "USD",
                "Monto": monto,
            })
    return rows


def parse_santander_cme(paths: list[Path]) -> list[dict]:
    by_account = {}
    for path in paths:
        for text in _pages(path):
            if not re.search(r"Saldo Libre", text, re.IGNORECASE):
                continue
            account = re.search(r"CUENTA:\s*([A-Z0-9]+)", text)
            descr = re.search(r"DESCR\.\s*CTA:\s*(SIEFORE INVERCAP[^\n]+)", text, re.IGNORECASE)
            saldo = re.search(r"Saldo Libre disposici[o\xf3n]+\s+([\-\d\.\,]+)", text, re.IGNORECASE)
            if not (descr and saldo):
                continue
            siefore = re.sub(r"\s*SA DE CV.*$", "", descr.group(1).strip().upper()).strip()
            row = {
                "Counterparty": "SANTANDER CME",
                "siefore_raw": siefore,
                "account_raw": account.group(1).strip() if account else "",
                "Currency": "USD",
                "Monto": _normal_number(saldo.group(1)),
            }
            by_account[row["account_raw"]] = row
    return list(by_account.values())


def parse_santander_mexder(paths: list[Path]) -> list[dict]:
    rows = []
    for path in paths:
        text = "\n".join(_pages(path))
        siefore = re.search(r"(SIEFORE INVERCAP BASICA[\s\w]+?)(?:\s+SA DE CV|\s+RFC|\s*\n)", text, re.IGNORECASE)
        account = re.search(r"ACCOUNT NO\.?:\s*([A-Z0-9]+)", text, re.IGNORECASE)
        value = re.search(r"VALOR TOTAL DE LA CUENTA[^\n]*?([\-]?[\d,\.]+)\s*$", text, re.IGNORECASE | re.MULTILINE)
        if not (siefore and value):
            continue
        rows.append({
            "Counterparty": "SANTANDER MEXDER",
            "siefore_raw": siefore.group(1).strip().upper(),
            "account_raw": account.group(1).strip() if account else "",
            "Currency": "MXN",
            "Monto": _normal_number(value.group(1)),
        })
    return rows


def parse_scotia_mexder(paths: list[Path]) -> list[dict]:
    rows = []
    for path in paths:
        text = "\n".join(_pages(path))
        siefore = re.search(r"SIEFORE INVERCAP BASICA,?\s+[\w][\w\-\s]*", text, re.IGNORECASE)
        account = re.search(r"ACCOUNT NUMBER[:\s]+([\d]+)", text, re.IGNORECASE)
        margin = re.search(r"MARGIN\s+DEFAULT/EXCESS\s+([\d,\.\s]+?)\s+(CR|DR)", text, re.IGNORECASE)
        if not (siefore and margin):
            continue
        siefore_raw = re.sub(r"\s+(AV|COL|RFC|CP|SA|DEL|LORENZO|CUIDAD)\b.*", "", siefore.group(0).strip().upper(), flags=re.IGNORECASE).strip()
        monto = abs(_normal_number(margin.group(1)))
        if margin.group(2).upper() == "DR":
            monto = -monto
        rows.append({
            "Counterparty": "SCOTIA",
            "siefore_raw": siefore_raw,
            "account_raw": account.group(1).strip() if account else "",
            "Currency": "MXN",
            "Monto": monto,
        })
    return rows


def _normal_account(account: object) -> str:
    return re.sub(r"^0+", "", str(account).strip().upper())


def enrich(rows: list[dict], valuation_date: date) -> pd.DataFrame:
    ref = pd.DataFrame(REF_DATA, columns=["Counterparty", "Portfolio", "Account", "Portafolio"])
    output = []
    aliases = {
        "SANTANDER MEXDER": "BCSANT",
        "GOLDMAN CME": "GOLDMAN",
        "SANTANDER CME": "SANTANDER CME",
        "SCOTIA": "SCOTIA",
        "GOLDMAN": "GOLDMAN",
    }
    for row in rows:
        cp_ref = aliases.get(row["Counterparty"].upper(), row["Counterparty"].upper())
        sub = ref[ref["Counterparty"].str.upper() == cp_ref]
        match = sub[sub["Account"].str.upper() == str(row["account_raw"]).upper()]
        if match.empty:
            match = sub[sub["Account"].apply(_normal_account) == _normal_account(row["account_raw"])]
        if not match.empty:
            found = match.iloc[0]
            portfolio, account, portafolio = found["Portfolio"], found["Account"], found["Portafolio"]
        else:
            portfolio, account, portafolio = row["siefore_raw"], row["account_raw"], ""
        output.append({
            "Counterparty": row["Counterparty"],
            "Portfolio": portfolio,
            "Date": valuation_date.strftime("%d/%m/%Y"),
            "Currency": row["Currency"],
            "Account": account,
            "Monto": row["Monto"],
            "Portafolio": portafolio,
        })
    return pd.DataFrame(output, columns=["Counterparty", "Portfolio", "Date", "Currency", "Account", "Monto", "Portafolio"])


def parse_statement_uploads(
    goldman_cme: list[Path] | None = None,
    santander_cme: list[Path] | None = None,
    santander_mexder: list[Path] | None = None,
    scotia_mexder: list[Path] | None = None,
    valuation_date: date | None = None,
) -> pd.DataFrame:
    valuation_date = valuation_date or last_business_day(date.today())
    rows: list[dict] = []
    rows += parse_goldman_cme(goldman_cme or [])
    rows += parse_santander_cme(santander_cme or [])
    rows += parse_santander_mexder(santander_mexder or [])
    rows += parse_scotia_mexder(scotia_mexder or [])
    return enrich(rows, valuation_date)
