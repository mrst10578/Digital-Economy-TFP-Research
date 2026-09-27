from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
RAW_DATA = BASE_DIR / "data" / "raw" / "dataset.csv"
PROCESSED_DATA = BASE_DIR / "data" / "processed" / "dataset_clean.csv"
TABLE_DIR = BASE_DIR / "outputs" / "tables"
FIGURE_DIR = BASE_DIR / "outputs" / "figures"

ID_COLUMN = "country"
TIME_COLUMN = "year"

EXPECTED_COLUMNS = [
    "country",
    "year",
    "tfp",
    "digital_economy",
    "ai",
    "human_capital",
    "investment",
    "trade_openness",
    "fdi",
    "institutional_quality",
    "energy",
]

MODEL_NAME = "TO_BE_DEFINED_AFTER_PROFESSOR_INSTRUCTIONS"
