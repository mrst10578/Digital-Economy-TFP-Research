import json
import pandas as pd

from src.config import RAW_DATA, PROCESSED_DATA, TABLE_DIR
from src.data_check import audit_data, validate_minimum_structure
from src.prepare_data import prepare_data
from src.dashboard_export import write_dashboard_snapshot


def main():
    if not RAW_DATA.exists():
        raise FileNotFoundError(
            f"فایل داده پیدا نشد: {RAW_DATA}\n"
            "فایل واقعی را با نام dataset.csv در data/raw قرار بده."
        )

    raw = pd.read_csv(RAW_DATA)
    clean = prepare_data(raw)
    validate_minimum_structure(clean)
    audit = audit_data(clean)

    PROCESSED_DATA.parent.mkdir(parents=True, exist_ok=True)
    TABLE_DIR.mkdir(parents=True, exist_ok=True)

    clean.to_csv(PROCESSED_DATA, index=False)

    with open(TABLE_DIR / "data_audit.json", "w", encoding="utf-8") as f:
        json.dump(audit, f, ensure_ascii=False, indent=2)

    print("\n=== DATA AUDIT ===")
    print(json.dumps(audit, ensure_ascii=False, indent=2))

    desc = clean.describe(include="all").transpose()
    desc.to_csv(TABLE_DIR / "descriptive_statistics.csv")

    numeric = clean.select_dtypes(include="number").copy()
    if "year" in numeric.columns:
        numeric = numeric.drop(columns=["year"])

    if numeric.shape[1] >= 2:
        numeric.corr().to_csv(TABLE_DIR / "correlation_matrix.csv")

    dashboard_path = write_dashboard_snapshot()

    print("\nپیش‌پردازش و کنترل اولیه تمام شد.")
    print(f"داده تمیزشده: {PROCESSED_DATA}")
    print(f"گزارش کنترل داده: {TABLE_DIR / 'data_audit.json'}")
    print(f"داده داشبورد: {dashboard_path}")
    print("مدل نهایی عمداً اجرا نشد چون دستور دقیق استاد هنوز دریافت نشده.")


if __name__ == "__main__":
    main()
