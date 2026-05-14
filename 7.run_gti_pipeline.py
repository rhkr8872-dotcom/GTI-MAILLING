```python
# -*- coding: utf-8 -*-

from __future__ import annotations

import sys
import subprocess
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).resolve().parent

SCRIPTS = [
    "1.collect_gti_news.py",
    "3.merge_news.py",
    "5.GTI Mail Engine.py",
]

def log(msg):
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)

def run(script):

    path = ROOT / script

    if not path.exists():
        log(f"⚠️ FILE NOT FOUND: {script}")
        return False

    log("=" * 70)
    log(f"RUN START: {script}")
    log("=" * 70)

    try:

        result = subprocess.run(
            [sys.executable, str(path)],
            cwd=str(ROOT),
            text=True,
            capture_output=True,
            timeout=3600
        )

        if result.stdout:
            print(result.stdout)

        if result.stderr:
            print(result.stderr)

        if result.returncode != 0:

            log(f"⚠️ FAILED BUT CONTINUE: {script}")
            log(f"RETURN CODE: {result.returncode}")

            return False

        log(f"✅ COMPLETE: {script}")

        return True

    except Exception as e:

        log(f"⚠️ EXCEPTION BUT CONTINUE: {script}")
        log(str(e))

        return False


def main():

    log("#" * 80)
    log("GTI PIPELINE FINAL START")
    log("#" * 80)

    results = {}

    for script in SCRIPTS:

        ok = run(script)

        results[script] = ok

    log("#" * 80)
    log("GTI PIPELINE RESULT")
    log("#" * 80)

    for k, v in results.items():

        status = "OK" if v else "FAILED"

        log(f"{k} : {status}")

    log("✅ GTI PIPELINE FINISHED")
    

if __name__ == "__main__":
    main()
```
