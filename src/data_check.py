import pandas as pd
from .config import EXPECTED_COLUMNS, ID_COLUMN, TIME_COLUMN


def audit_data(df: pd.DataFrame) -> dict:
    report = {
        "n_rows": int(len(df)),
        "n_columns": int(len(df.columns)),
        "columns": list(df.columns),
        "missing_expected_columns": [
            c for c in EXPECTED_COLUMNS if c not in df.columns
        ],
        "missing_values": {
            k: int(v) for k, v in df.isna().sum().to_dict().items()
        },
        "dtypes": {k: str(v) for k, v in df.dtypes.items()},
    }

    if ID_COLUMN in df.columns and TIME_COLUMN in df.columns:
        report["duplicate_id_time"] = int(
            df.duplicated(subset=[ID_COLUMN, TIME_COLUMN]).sum()
        )
        report["missing_id"] = int(df[ID_COLUMN].isna().sum())
        report["missing_time"] = int(df[TIME_COLUMN].isna().sum())

        counts = (
            df.dropna(subset=[ID_COLUMN, TIME_COLUMN])
            .groupby(ID_COLUMN)[TIME_COLUMN]
            .nunique()
        )
        report["n_entities"] = int(counts.shape[0])
        report["min_periods_per_entity"] = int(counts.min()) if not counts.empty else 0
        report["max_periods_per_entity"] = int(counts.max()) if not counts.empty else 0
        report["balanced_panel_candidate"] = (
            bool(counts.nunique() == 1) if not counts.empty else False
        )
    else:
        report["duplicate_id_time"] = None
        report["missing_id"] = None
        report["missing_time"] = None
        report["n_entities"] = None
        report["min_periods_per_entity"] = None
        report["max_periods_per_entity"] = None
        report["balanced_panel_candidate"] = None

    return report


def validate_minimum_structure(df: pd.DataFrame) -> None:
    required = [ID_COLUMN, TIME_COLUMN]
    missing = [c for c in required if c not in df.columns]

    if missing:
        raise ValueError(
            f"ستون‌های پایه برای ساختار داده موجود نیستند: {missing}. "
            "نام ستون‌ها را با config.py تطبیق بده."
        )

    if df[ID_COLUMN].isna().all():
        raise ValueError("ستون شناسه واحدها کاملاً خالی است.")

    if df[TIME_COLUMN].isna().all():
        raise ValueError("ستون زمان/سال کاملاً خالی است.")
