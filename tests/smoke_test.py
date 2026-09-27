from pathlib import Path
import sys
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.prepare_data import prepare_data
from src.data_check import validate_minimum_structure, audit_data


def run():
    df = pd.DataFrame({
        " Country ": ["A", "A", "B", "B"],
        "Year": [2020, 2021, 2020, 2021],
        "TFP": [1.0, 1.1, 0.9, 1.0],
    })

    clean = prepare_data(df)
    validate_minimum_structure(clean)
    report = audit_data(clean)

    assert clean.columns.tolist() == ["country", "year", "tfp"]
    assert report["duplicate_id_time"] == 0
    assert report["n_entities"] == 2
    assert report["balanced_panel_candidate"] is True

    print("SMOKE TEST PASSED")


if __name__ == "__main__":
    run()
