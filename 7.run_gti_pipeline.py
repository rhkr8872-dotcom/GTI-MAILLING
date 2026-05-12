# -*- coding: utf-8 -*-
"""Run GTI STEP4 analysis and STEP5 mail generation in order."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent


def run(script: str) -> None:
    subprocess.run([sys.executable, str(ROOT / script)], check=True)


if __name__ == "__main__":
    run("4.policy_ai_analyzer.py")
    run("5.GTI_Mail_Engine.py")
