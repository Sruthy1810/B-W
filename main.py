import logging
import config
from core.writer import append_datatable
from processes import REGISTRY


def main():
    config.LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.FileHandler(config.LOG_FILE, encoding="utf-8"),
                                  logging.StreamHandler()])
    log = logging.getLogger("main")

    input_folder = config.get_input_file()
    output_file = config.get_output_file()

    for name in config.ENABLED_PROCESSES:
        proc = REGISTRY[name]()
        log.info("=== %s ===", name)
        dt = proc.build_datatable(input_folder)               # 1) DataTable first
        if dt.empty:
            log.warning("[%s] nothing to write", name)
            continue
        append_datatable(output_file, proc.sheet_name, dt,    # 2) single write to Excel
                         proc.out_columns, proc.sum_col_idx)
    log.info("Done.")


if __name__ == "__main__":
    main()