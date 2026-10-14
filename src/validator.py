import logging
import pandas as pd

log = logging.getLogger("remittance")


def check_totals(df: pd.DataFrame, col_a: int, col_b: int,
                 expected_a: float, expected_b: float) -> bool:
    got_a = round(df.iloc[:, col_a].sum(), 2)
    got_b = round(df.iloc[:, col_b].sum(), 2)
    ok_a = abs(got_a - expected_a) < 0.01
    ok_b = abs(got_b - expected_b) < 0.01

    log.info("Column %d: expected %.2f, got %.2f -> %s",
             col_a, expected_a, got_a, "PASS" if ok_a else "FAIL")
    log.info("Column %d: expected %.2f, got %.2f -> %s",
             col_b, expected_b, got_b, "PASS" if ok_b else "FAIL")
    return ok_a and ok_b