import re
import logging
import pandas as pd

log = logging.getLogger("remittance")
CLAIM_RE = re.compile(r"^\d{7}\.\d{3}$")


def _to_number(value):
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def _claim_text(value):
    if isinstance(value, float):
        return f"{value:.3f}"
    return str(value).strip()


def _pick_sheet(xls: pd.ExcelFile, wanted: str) -> str:
    names = xls.sheet_names
    log.info("Sheets found: %s", names)
    for name in names:
        if wanted.lower() in name.lower():
            return name
    log.warning("Sheet '%s' not found, scanning all sheets for debtor rows", wanted)
    best, best_count = names[0], -1
    for name in names:
        raw = xls.parse(name, header=None)
        count = sum(
            1 for v in raw.iloc[:, :3].values.flatten()
            if pd.notna(v) and CLAIM_RE.match(_claim_text(v))
        )
        if count > best_count:
            best, best_count = name, count
    return best


def read_report(path: str, sheet_name: str, columns: list) -> pd.DataFrame:
    log.info("Reading file: %s", path)
    xls = pd.ExcelFile(path)
    sheet = _pick_sheet(xls, sheet_name)
    log.info("Using sheet: %s", sheet)

    raw = xls.parse(sheet, header=None)
    n = len(columns)
    records = []

    for _, row in raw.iterrows():
        vals = [v for v in row.tolist() if pd.notna(v) and str(v).strip() != ""]
        if not vals:
            continue
        first = _claim_text(vals[0])

        if CLAIM_RE.match(first):
            if len(vals) < n + 2:
                log.warning("Skipping short row for claim %s: %s", first, vals)
                continue
            numbers = [_to_number(v) for v in vals[-n:]]
            if None in numbers:
                log.warning("Skipping row with non-numeric values, claim %s", first)
                continue
            record = {
                "Claim": first,
                "Debtor Name": str(vals[1]).strip(),
                "Ref Acct #": str(vals[2]).strip() if len(vals) > n + 2 else "",
            }
            record.update(dict(zip(columns, numbers)))
            record["Plaintiff"] = ""
            records.append(record)

        elif first.startswith("Plaintiff:") and records:
            records[-1]["Plaintiff"] = first.replace("Plaintiff:", "").strip()

    df = pd.DataFrame(records)
    log.info("Debtor rows read: %d", len(df))
    return df