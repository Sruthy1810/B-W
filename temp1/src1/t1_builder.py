import pandas as pd

PAYMENT_COLUMNS = [
    "DEFENDANT'S NAME", "REF NO.", "AMOUNT PAID", "DIRECT PAYMENT",
    "CLIENT SHARE", "BW, LLP SHARE", "Notes",
]

COST_COLUMNS = ["DEFENDANT'S NAME", "REF NO.", "AMOUNT", "COST REASON", "Date", "Notes"]

def build_payment_sheet(data, client_pct):
    out = pd.DataFrame(columns=PAYMENT_COLUMNS)
    if data.empty:
        return out

    amount = (data["principal"] + data["interest"]).round(2)
    client = (amount * client_pct).round(2)          # Amount Paid x client % from the master

    out["DEFENDANT'S NAME"] = data["name"]
    out["REF NO."] = data["ref_no"]
    out["AMOUNT PAID"] = amount
    out["DIRECT PAYMENT"] = data["direct_payment"]
    out["CLIENT SHARE"] = client                      # <-- was ""
    out["BW, LLP SHARE"] = (amount - client).round(2)  # <-- was "" (remainder, so the two add up)
    out["Notes"] = ""
    return out


def build_cost_sheet(data):
    out = pd.DataFrame(columns=COST_COLUMNS)
    if data.empty:
        return out
    out["DEFENDANT'S NAME"] = data["name"]
    out["REF NO."] = data["ref_no"]
    out["AMOUNT"] = data["amount"]
    out["COST REASON"] = data["reason"]
    out["Date"] = data["date"]
    out["Notes"] = data["note"]
    return out