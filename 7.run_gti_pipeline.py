# -*- coding: utf-8 -*-
"""
GTI PIPELINE v43 - REGULATION / NEWS FULLY SEPARATED
====================================================

REGULATION BRANCH
    1.site_crawler.py
        -> 1-1.regulation_raw.xlsx
    3-1.regulation_merge.py
        -> 3-1.regulation_summary.xlsx
        -> 3-1.regulation_article_summary.xlsx
        -> 3-1.regulation_cumulative.xlsx
    4-1.regulation_ai_analysis.py
        -> 4-1.regulation_ai_summary.xlsx
        -> 4-1.regulation_ai_cumulative.xlsx

NEWS BRANCH
    2-1.NAVER_news_collector.py
    2-2.google_news_collector.py
    2-3.rss_news_raw.py
        -> news raw files
    3-2.news_merge.py
        -> 3-2.news_summary.xlsx
        -> 3-2.news_cumulative.xlsx
    4-2.news_ai_analysis.py
        -> 4-2.news_ai_summary.xlsx
        -> 4-2.news_ai_cumulative.xlsx

FINAL MAIL
    4-1.regulation_ai_summary.xlsx
        +
    4-2.news_ai_summary.xlsx
        -> 5.GTI_Mail_Engine.py
        -> GTI Radar HTML / XLSX

Critical rules
--------------
1. Regulation and News never depend on each other before Step5.
2. 3-1 never creates/checks news files.
3. 3-2 never creates/checks regulation files.
4. 4-1 reads regulation output only.
5. 4-2 reads news output only.
6. Step5 runs only when BOTH current-run AI outputs are freshly generated.
7. Old AI outputs are never mixed into a new mail.
"""

from __future__ import annotations

import argparse
import os
import queue
import shutil
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd


BASE_DIR = Path(os.getenv("GTI_BASE_DIR", r"C:\Temp"))
PYTHON_EXE = Path(
    os.getenv(
        "GTI_PYTHON_EXE",
        r"C:\Users\KCH\AppData\Local\Programs\Python\Python312\python.exe",
    )
)

LOG_DIR = BASE_DIR / "logs"
ARCHIVE_DIR = BASE_DIR / "archive"
LOG_FILE = LOG_DIR / "gti_pipeline_v43_split.log"

MAIL_OUTPUT_DIR = BASE_DIR / "12345" / "c_type_outputs"


PIPELINE_ENV_DEFAULTS = {
    # Regulation collection can be wider; new-event control is handled by 3-1.
    "GTI_STEP1_HOURS_BACK": "72",

    # News: broad collection -> strict recent report candidates.
    "GTI_LOOKBACK_HOURS": "72",
    "GTI_STEP3_RECENT_HOURS": "24",
    "GTI_STEP4_NEWS_MAX_AGE_HOURS": "24",

    # Step5 final hard guard.
    "GTI_MAIL_NEWS_HOURS": "24",

    # Gemini.
    "GTI_GEMINI_MODEL": "gemini-2.5-flash-lite",
    "GTI_GEMINI_TIMEOUT": "20",

    # Network.
    "GTI_ARTICLE_FETCH_TIMEOUT": "12",
    "GTI_RSS_FETCH_TIMEOUT": "15",

    # News output.
    "GTI_STEP3_TARGET_MAX": "300",
    "GTI_STEP4_NEWS_TARGET_MAX": "30",
    "GTI_STRICT_NEWS_TARGET_MAX": "30",

    # URL.
    "GTI_STEP2_RESOLVE_ORIGINAL_URL": "N",
    "GTI_STEP3_RESOLVE_URL": "Y",
}


@dataclass(frozen=True)
class Step:
    name: str
    script: str
    required: bool = True
    expected_outputs: tuple[str, ...] = field(default_factory=tuple)
    min_rows: dict[str, int] = field(default_factory=dict)
    args: tuple[str, ...] = field(default_factory=tuple)
    timeout_sec: int | None = None


# =============================================================================
# REGULATION BRANCH
# =============================================================================

REGULATION_STEPS = [
    Step(
        name="REG_1_CRAWL",
        script="1.site_crawler.py",
        expected_outputs=("1-1.regulation_raw.xlsx",),
        timeout_sec=1800,
    ),
    Step(
        name="REG_3_1_SELECT",
        script="3-1.regulation_merge.py",
        expected_outputs=(
            "3-1.regulation_summary.xlsx",
            "3-1.regulation_article_summary.xlsx",
            "3-1.regulation_cumulative.xlsx",
        ),
        timeout_sec=3600,
    ),
    Step(
        name="REG_4_1_AI",
        script="4-1.regulation_ai_analysis.py",
        expected_outputs=(
            "4-1.regulation_ai_summary.xlsx",
            "4-1.regulation_ai_cumulative.xlsx",
        ),
        timeout_sec=3600,
    ),
]


# =============================================================================
# NEWS BRANCH
# =============================================================================

NEWS_COLLECTOR_STEPS = [
    Step(
        name="NEWS_2_1_NAVER",
        script="2-1.NAVER_news_collector.py",
        required=False,
        expected_outputs=("2-1.naver_news_raw.xlsx",),
        timeout_sec=1200,
    ),
    Step(
        name="NEWS_2_2_GOOGLE",
        script="2-2.google_news_collector.py",
        required=False,
        expected_outputs=("2-2.google_news_raw.xlsx",),
        timeout_sec=1800,
    ),
    Step(
        name="NEWS_2_3_RSS_SITE",
        script="2-3.rss_news_raw.py",
        required=False,
        expected_outputs=("2-3.rss_news_raw.xlsx",),
        timeout_sec=1800,
    ),
]

NEWS_POST_STEPS = [
    Step(
        name="NEWS_3_2_MERGE",
        script="3-2.news_merge.py",
        expected_outputs=(
            "3-2.news_summary.xlsx",
            "3-2.news_cumulative.xlsx",
        ),
        min_rows={"3-2.news_summary.xlsx": 1},
        timeout_sec=1800,
    ),
    Step(
        name="NEWS_4_2_AI",
        script="4-2.news_ai_analysis.py",
        expected_outputs=(
            "4-2.news_ai_summary.xlsx",
            "4-2.news_ai_cumulative.xlsx",
        ),
        min_rows={"4-2.news_ai_summary.xlsx": 1},
        timeout_sec=3600,
    ),
]


# =============================================================================
# FINAL MAIL
# =============================================================================

MAIL_STEP = Step(
    name="FINAL_5_MAIL",
    script="5.GTI_Mail_Engine.py",
    expected_outputs=(),
    args=(
        "--regulation-input",
        str(BASE_DIR / "4-1.regulation_ai_summary.xlsx"),
        "--news-input",
        str(BASE_DIR / "4-2.news_ai_summary.xlsx"),
        "--output-dir",
        str(MAIL_OUTPUT_DIR),
    ),
    timeout_sec=1800,
)


ARCHIVE_TARGETS = [
    # Regulation
    "1-1.regulation_raw.xlsx",
    "3-1.regulation_summary.xlsx",
    "3-1.regulation_article_summary.xlsx",
    "3-1.regulation_cumulative.xlsx",
    "3-1.regulation_audit.xlsx",
    "3-1.regulation_excluded.xlsx",
    "3-1.regulation_cumulative_removed.xlsx",
    "4-1.regulation_ai_summary.xlsx",
    "4-1.regulation_ai_cumulative.xlsx",
    "4-1.regulation_ai_excluded.xlsx",
    "4-1.regulation_ai_diagnostic.xlsx",

    # News
    "2-1.naver_news_raw.xlsx",
    "2-2.google_news_raw.xlsx",
    "2-3.rss_news_raw.xlsx",
    "3-2.news_summary.xlsx",
    "3-2.news_cumulative.xlsx",
    "3-2.news_excluded.xlsx",
    "3-2.news_cluster_audit.xlsx",
    "4-2.news_ai_summary.xlsx",
    "4-2.news_ai_cumulative.xlsx",
    "4-2.news_ai_audit_candidates.xlsx",
    "4-2.news_ai_excluded.xlsx",

    # Final
    "4.gti_mail_input.xlsx",
    "GTI_Radar.xlsx",
    "mail_cumulative.xlsx",
]


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str = "") -> None:
    line = f"[{now()}] {msg}"
    print(line)
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def ensure_dirs() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    MAIL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def apply_env_defaults() -> None:
    for k, v in PIPELINE_ENV_DEFAULTS.items():
        os.environ.setdefault(k, v)


def get_python() -> str:
    candidates = [
        str(PYTHON_EXE),
        sys.executable,
        shutil.which("python"),
        shutil.which("py"),
    ]
    for x in candidates:
        if x and Path(x).exists():
            return x
    return sys.executable


def excel_row_count(path: Path) -> int:
    try:
        if not path.exists() or path.stat().st_size == 0:
            return -1
        return len(pd.read_excel(path))
    except Exception:
        return -1


def valid_output(path: Path, started_at: float, min_rows: int | None = None) -> tuple[bool, str]:
    if not path.exists():
        return False, "missing"

    if path.stat().st_size == 0:
        return False, "empty file"

    # Must be generated/refreshed by this run.
    if path.stat().st_mtime + 1 < started_at:
        return False, "not refreshed"

    if path.suffix.lower() in {".xlsx", ".xls"}:
        try:
            df = pd.read_excel(path)
        except Exception as exc:
            return False, f"excel read error:{type(exc).__name__}"

        # Regulation can legitimately be zero rows. A valid header is enough.
        if len(df.columns) == 0:
            return False, "no columns"

        if min_rows is not None and len(df) < min_rows:
            return False, f"rows={len(df)} < {min_rows}"

    return True, ""


def validate_step_outputs(step: Step, started_at: float) -> tuple[bool, list[str]]:
    bad = []

    for filename in step.expected_outputs:
        path = BASE_DIR / filename
        ok, reason = valid_output(
            path,
            started_at,
            step.min_rows.get(filename),
        )
        if not ok:
            bad.append(f"{filename} ({reason})")

    return not bad, bad


def run_step(step: Step, python_exe: str, dry_run: bool = False) -> str:
    script = BASE_DIR / step.script

    log("=" * 80)
    log(f"{step.name} START : {step.script}")
    log("=" * 80)

    if not script.exists():
        log(f"FILE NOT FOUND : {script}")
        return "FAILED" if step.required else "SKIPPED"

    cmd = [python_exe, str(script), *step.args]
    log("COMMAND : " + " ".join(f'"{x}"' if " " in x else x for x in cmd))

    if dry_run:
        return "DRY_RUN"

    started = time.time()

    env = os.environ.copy()
    env.setdefault("PYTHONIOENCODING", "utf-8")
    env.setdefault("PYTHONUTF8", "1")

    proc = subprocess.Popen(
        cmd,
        cwd=str(BASE_DIR),
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    assert proc.stdout is not None

    q: queue.Queue[str | None] = queue.Queue()

    def reader():
        try:
            for line in proc.stdout:
                q.put(line)
        finally:
            q.put(None)

    thread = threading.Thread(target=reader, daemon=True)
    thread.start()

    reader_done = False
    timed_out = False

    while True:
        try:
            item = q.get(timeout=0.5)
            if item is None:
                reader_done = True
            else:
                log("  " + item.rstrip())
        except queue.Empty:
            pass

        if (
            step.timeout_sec
            and proc.poll() is None
            and time.time() - started > step.timeout_sec
        ):
            timed_out = True
            log(f"TIMEOUT : {step.timeout_sec} sec")
            try:
                if os.name == "nt":
                    subprocess.run(
                        ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
                else:
                    proc.kill()
            except Exception:
                pass
            break

        if proc.poll() is not None and reader_done:
            break

    try:
        thread.join(timeout=2)
    except Exception:
        pass

    while not q.empty():
        item = q.get_nowait()
        if item:
            log("  " + item.rstrip())

    rc = proc.wait()
    elapsed = round(time.time() - started, 2)

    if timed_out:
        log(f"{step.name} FAILED : timeout / {elapsed}s")
        return "FAILED" if step.required else "WARNING"

    if rc != 0:
        log(f"{step.name} FAILED : return_code={rc} / {elapsed}s")
        return "FAILED" if step.required else "WARNING"

    ok, bad = validate_step_outputs(step, started)
    if not ok:
        log(f"{step.name} OUTPUT CHECK FAILED : {' / '.join(bad)}")
        return "FAILED" if step.required else "WARNING"

    log(f"{step.name} COMPLETE : {elapsed}s")
    return "OK"


def archive_previous() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = ARCHIVE_DIR / stamp
    folder.mkdir(parents=True, exist_ok=True)

    count = 0
    for filename in ARCHIVE_TARGETS:
        src = BASE_DIR / filename
        if not src.exists():
            continue
        try:
            shutil.copy2(src, folder / filename)
            count += 1
        except Exception as exc:
            log(f"ARCHIVE WARN : {filename} / {exc}")

    log(f"ARCHIVE COMPLETE : {count} files -> {folder}")


def run_regulation_branch(
    python_exe: str,
    dry_run: bool,
    results: list[tuple[str, str, str]],
) -> bool:
    log("")
    log("#" * 80)
    log("REGULATION BRANCH START")
    log("#" * 80)

    for step in REGULATION_STEPS:
        status = run_step(step, python_exe, dry_run)
        results.append((step.name, step.script, status))

        if status == "FAILED":
            log(f"REGULATION BRANCH STOP : {step.name}")
            return False

    log("REGULATION BRANCH COMPLETE")
    return True



NEWS_RAW_BY_STEP = {
    "NEWS_2_1_NAVER": BASE_DIR / "2-1.naver_news_raw.xlsx",
    "NEWS_2_2_GOOGLE": BASE_DIR / "2-2.google_news_raw.xlsx",
    "NEWS_2_3_RSS_SITE": BASE_DIR / "2-3.rss_news_raw.xlsx",
}

def purge_news_raw_before_collectors() -> None:
    """
    Remove previous-run raw files after archiving.
    This prevents a missing/failed collector from leaking yesterday's raw data
    into today's 3-2 merge.
    """
    for path in NEWS_RAW_BY_STEP.values():
        if path.exists():
            try:
                path.unlink()
                log(f"STALE RAW PURGED : {path.name}")
            except Exception as exc:
                log(f"STALE RAW PURGE FAILED : {path.name} / {exc}")


def run_news_branch(
    python_exe: str,
    dry_run: bool,
    results: list[tuple[str, str, str]],
) -> bool:
    log("")
    log("#" * 80)
    log("NEWS BRANCH START")
    log("#" * 80)

    if not dry_run:
        purge_news_raw_before_collectors()

    collector_ok = 0

    for step in NEWS_COLLECTOR_STEPS:
        status = run_step(step, python_exe, dry_run)
        results.append((step.name, step.script, status))

        if status in {"OK", "DRY_RUN"}:
            collector_ok += 1
        elif not dry_run:
            stale = NEWS_RAW_BY_STEP.get(step.name)
            if stale and stale.exists():
                try:
                    stale.unlink()
                    log(f"FAILED COLLECTOR RAW REMOVED : {stale.name}")
                except Exception as exc:
                    log(f"FAILED COLLECTOR RAW REMOVE WARN : {stale.name} / {exc}")

    if collector_ok == 0:
        log("NEWS BRANCH STOP : no collector available")
        return False

    for step in NEWS_POST_STEPS:
        status = run_step(step, python_exe, dry_run)
        results.append((step.name, step.script, status))

        if status == "FAILED":
            log(f"NEWS BRANCH STOP : {step.name}")
            return False

    log("NEWS BRANCH COMPLETE")
    return True


def ai_outputs_current_run_ready(run_started_at: float) -> tuple[bool, list[str]]:
    checks = [
        BASE_DIR / "4-1.regulation_ai_summary.xlsx",
        BASE_DIR / "4-2.news_ai_summary.xlsx",
    ]

    bad = []

    for path in checks:
        if not path.exists():
            bad.append(f"{path.name}:missing")
            continue

        if path.stat().st_mtime + 1 < run_started_at:
            bad.append(f"{path.name}:stale")
            continue

        try:
            df = pd.read_excel(path)
            if len(df.columns) == 0:
                bad.append(f"{path.name}:no columns")
        except Exception as exc:
            bad.append(f"{path.name}:{type(exc).__name__}")

    return not bad, bad


def run_mail(
    python_exe: str,
    dry_run: bool,
    run_started_at: float,
    regulation_ok: bool,
    news_ok: bool,
    results: list[tuple[str, str, str]],
) -> bool:
    log("")
    log("#" * 80)
    log("FINAL COMBINED MAIL START")
    log("#" * 80)

    if dry_run:
        results.append((MAIL_STEP.name, MAIL_STEP.script, "DRY_RUN"))
        return True

    if not regulation_ok or not news_ok:
        log(
            "MAIL BLOCKED : current-run branch failure "
            f"/ regulation_ok={regulation_ok} / news_ok={news_ok}"
        )
        results.append((MAIL_STEP.name, MAIL_STEP.script, "SKIPPED"))
        return False

    ready, bad = ai_outputs_current_run_ready(run_started_at)
    if not ready:
        log("MAIL BLOCKED : stale/missing AI output / " + " / ".join(bad))
        results.append((MAIL_STEP.name, MAIL_STEP.script, "SKIPPED"))
        return False

    status = run_step(MAIL_STEP, python_exe, dry_run=False)
    results.append((MAIL_STEP.name, MAIL_STEP.script, status))
    return status == "OK"


def print_result(results: list[tuple[str, str, str]]) -> None:
    log("")
    log("#" * 80)
    log("GTI PIPELINE v43 SPLIT RESULT")
    log("#" * 80)

    counts = {}
    for name, script, status in results:
        log(f"{name} / {script} : {status}")
        counts[status] = counts.get(status, 0) + 1

    log("-" * 80)
    for key in ["OK", "WARNING", "SKIPPED", "DRY_RUN", "FAILED"]:
        log(f"{key:<8}: {counts.get(key, 0)}")
    log("#" * 80)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="GTI v42 fully separated regulation/news pipeline"
    )
    p.add_argument("--no-archive", action="store_true")
    p.add_argument("--skip-mail", action="store_true")
    p.add_argument("--regulation-only", action="store_true")
    p.add_argument("--news-only", action="store_true")
    p.add_argument("--dry-run", action="store_true")
    return p.parse_args()


def main() -> int:
    args = parse_args()

    ensure_dirs()
    apply_env_defaults()

    run_started_at = time.time()
    python_exe = get_python()

    log("#" * 80)
    log("GTI PIPELINE v43 - REGULATION / NEWS FULLY SEPARATED START")
    log("#" * 80)
    log(f"BASE_DIR : {BASE_DIR}")
    log(f"PYTHON   : {python_exe}")

    if not args.no_archive and not args.dry_run:
        archive_previous()

    results: list[tuple[str, str, str]] = []

    regulation_ok = False
    news_ok = False

    if not args.news_only:
        regulation_ok = run_regulation_branch(
            python_exe,
            args.dry_run,
            results,
        )

    if not args.regulation_only:
        news_ok = run_news_branch(
            python_exe,
            args.dry_run,
            results,
        )

    mail_ok = True

    if args.skip_mail:
        log("FINAL MAIL SKIPPED BY OPTION")
    elif args.regulation_only or args.news_only:
        log("FINAL MAIL SKIPPED : single-branch mode")
    else:
        mail_ok = run_mail(
            python_exe,
            args.dry_run,
            run_started_at,
            regulation_ok,
            news_ok,
            results,
        )

    print_result(results)

    failed = any(status == "FAILED" for _, _, status in results)

    if (
        not failed
        and (args.regulation_only or regulation_ok)
        and (args.news_only or news_ok)
        and mail_ok
    ):
        log("GTI PIPELINE v43 FINISHED")
        return 0

    log("GTI PIPELINE v43 FINISHED WITH ERROR")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
