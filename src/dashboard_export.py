from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from .config import BASE_DIR, EXPECTED_COLUMNS

OUTPUT_PATH = BASE_DIR / "site" / "data" / "project.json"
AUDIT_PATH = BASE_DIR / "outputs" / "tables" / "data_audit.json"


def _read_json(path: Path):
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None


def _period_label(audit: dict | None) -> str | None:
    # Exact time range is not inferred unless it exists in an analysis output.
    if not audit:
        return None
    return audit.get("period_label")


def build_dashboard_snapshot() -> dict:
    audit = _read_json(AUDIT_PATH)

    now = datetime.now(timezone.utc)
    snapshot = {
        "project": {
            "title": "Digital Economy × TFP Research",
            "language": "fa",
        },
        "stage": {
            "short": "Model Specification",
            "title": "مدل نهایی در انتظار دستور دقیق استاد",
            "description": (
                "زیرساخت داده و اجرای Python آماده است. "
                "برآورد نهایی فقط بعد از دریافت specification استاد انجام می‌شود."
            ),
        },
        "generated_at": now.isoformat(timespec="seconds"),
        "generated_at_local": now.strftime("%Y-%m-%d %H:%M UTC"),
        "expected_columns": EXPECTED_COLUMNS,
        "period_label": _period_label(audit),
        "audit": audit or {
            "n_rows": None,
            "n_columns": None,
            "missing_values": None,
            "duplicate_id_time": None,
            "n_entities": None,
            "min_periods_per_entity": None,
            "max_periods_per_entity": None,
            "balanced_panel_candidate": None,
        },
        "model": {
            "specified": False,
            "name": None,
            "description": None,
        },
        "results": {
            "coefficients": [],
            "diagnostics": [],
        },
    }
    return snapshot


def write_dashboard_snapshot() -> Path:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    snapshot = build_dashboard_snapshot()
    OUTPUT_PATH.write_text(
        json.dumps(snapshot, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return OUTPUT_PATH


if __name__ == "__main__":
    path = write_dashboard_snapshot()
    print(f"Dashboard snapshot written to: {path}")
