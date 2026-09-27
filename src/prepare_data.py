import re
import pandas as pd
from .config import ID_COLUMN, TIME_COLUMN


def normalize_column_name(name: object) -> str:
    name = str(name).strip().lower()
    name = re.sub(r"[^a-z0-9_]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name


def prepare_data(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [normalize_column_name(c) for c in out.columns]
    out = out.dropna(how="all")

    if TIME_COLUMN in out.columns:
        out[TIME_COLUMN] = pd.to_numeric(out[TIME_COLUMN], errors="coerce")

    if ID_COLUMN in out.columns:
        out[ID_COLUMN] = out[ID_COLUMN].astype("string").str.strip()

    if ID_COLUMN in out.columns and TIME_COLUMN in out.columns:
        out = out.sort_values(
            [ID_COLUMN, TIME_COLUMN], na_position="last"
        ).reset_index(drop=True)

    return out
