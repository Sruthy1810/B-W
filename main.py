from pathlib import Path

from src.config_loader import load_config
from src.logger_setup import setup_logger
from src.reader import read_report
from src.validator import check_totals
from src.cleaner import keep_valid_rows


def main():
    config = load_config()
    log = setup_logger(config["log_file"], config["log_level"])
    log.info("===== Run started =====")

    try:
        df = read_report(config["input_file"], config["sheet_name"],
                         config["report_columns"])

        totals_ok = check_totals(df, config["principal_col_index"],
                                 config["interest_col_index"],
                                 config["expected_principal_total"],
                                 config["expected_interest_total"])
        if not totals_ok:
            log.warning("Totals do not match the report. Some rows may have been missed.")

        valid = keep_valid_rows(df, config["principal_col_index"],
                                config["interest_col_index"], config["rule_mode"])

        Path(config["output_file"]).parent.mkdir(parents=True, exist_ok=True)
        valid.to_excel(config["output_file"], index=False)
        log.info("Saved valid data to %s", config["output_file"])

        print(valid.head(20).to_string())
        print(f"\nValid rows: {len(valid)}")
    except Exception:
        log.exception("Run failed")
        raise
    finally:
        log.info("===== Run finished =====")


if __name__ == "__main__":
    main()