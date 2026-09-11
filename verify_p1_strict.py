from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from gti_quality_contract import OPERATIVE_STAGES, verified_mapping


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    df = pd.read_excel(args.workbook, sheet_name="Executive Radar")
    required = {
        "ImpactStatus", "PolicyMateriality", "SamsungExposure",
        "EvidenceConfidence", "ActionUrgency", "SamsungDirectFlag",
        "OfficialSourceFlag", "Policy Stage", "EventKey",
    }
    missing = sorted(required - set(df.columns))
    if missing:
        raise RuntimeError(f"P1 columns missing: {missing}")
    if df["EventKey"].fillna("").duplicated().any():
        raise RuntimeError("Duplicate EventKey remains in Executive Radar")

    violations = []
    for idx, row in df[df["SamsungDirectFlag"].eq("Y")].iterrows():
        stage = str(row.get("Policy Stage", "")).upper().replace(" ", "_")
        if row.get("OfficialSourceFlag") != "Y" or stage not in OPERATIVE_STAGES or not verified_mapping(row):
            violations.append(int(idx) + 2)
    if violations:
        raise RuntimeError(f"Strict Direct Gate violation at Excel rows: {violations}")

    print(f"P1 STRICT VERIFY OK: rows={len(df)} / direct={int(df['SamsungDirectFlag'].eq('Y').sum())}")
    print("ImpactStatus:", df["ImpactStatus"].value_counts().to_dict())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
