# 7.run_gti_pipeline.py
# -*- coding: utf-8 -*-

import os
import sys
import time
import shutil
import subprocess
from datetime import datetime

BASE_DIR = r"C:\temp"

PYTHON_EXE = r"C:\Users\KCH\AppData\Local\Programs\Python\Python312\python.exe"

LOG_DIR = os.path.join(BASE_DIR, "logs")
ARCHIVE_DIR = os.path.join(BASE_DIR, "archive")
LOG_FILE = os.path.join(LOG_DIR, "gti_pipeline_run.log")

STEPS = [
    ("STEP1_SITE_CRAWLER", "1.site_crawler.py", False),
    ("STEP2_NAVER", "2-1.naver_news_collector.py", False),
    ("STEP2_GOOGLE", "2-2.google_news_collector.py", False),
    ("STEP2_RSS", "2-3._rss_news_raw.py", False),
    ("STEP3_MERGE", "3.merge_news.py", True),
    ("STEP4_ANALYZE", "4.policy_ai_analyzer.py", True),
    ("STEP5_MAIL", "5.GTI_Mail_Engine.py", True),
]

ARCHIVE_TARGETS = [
    "1.site_news_raw.xlsx",
    "2-1.naver_news_raw.xlsx",
    "2-2.google_news_raw.xlsx",
    "2-3.rss_news_raw.xlsx",
    "3.news_ai_summary.xlsx",
    "news_raw.xlsx",
    "GTI_Radar_2026-05-23_Top25.xlsx",
    "GTI_Radar_2026-05-23_Top25_Email.html",
]


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_dirs():
    os.makedirs(LOG_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_DIR, exist_ok=True)


def log(msg=""):
    line = f"[{now()}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def archive_outputs():
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target_dir = os.path.join(ARCHIVE_DIR, stamp)
    os.makedirs(target_dir, exist_ok=True)

    copied = 0

    for filename in ARCHIVE_TARGETS:
        src = os.path.join(BASE_DIR, filename)
        if os.path.exists(src):
            try:
                shutil.copy2(src, os.path.join(target_dir, filename))
                copied += 1
                log(f"ARCHIVE OK : {filename}")
            except Exception as e:
                log(f"ARCHIVE FAIL : {filename} / {e}")

    log(f"ARCHIVE COMPLETE : {copied} files")


def get_python_exe():
    candidates = [
        PYTHON_EXE,
        sys.executable,
        shutil.which("python"),
        shutil.which("py"),
    ]

    for exe in candidates:
        if exe and os.path.exists(exe):
            return exe

    return sys.executable


def run_script(step_name, script_file, required, python_exe):
    script_path = os.path.join(BASE_DIR, script_file)

    log("=" * 80)
    log(f"RUN START : {script_file}")
    log("=" * 80)

    if not os.path.exists(script_path):
        log(f"⚠️ FILE NOT FOUND : {script_file}")

        if required:
            log(f"❌ REQUIRED STEP FAILED : {script_file}")
            return "FAILED"
        else:
            log(f"⏭ OPTIONAL STEP SKIPPED : {script_file}")
            return "SKIPPED"

    start = time.time()

    try:
result = subprocess.run(
    [python_exe, script_path],
    cwd=BASE_DIR,
    capture_output=True,
    text=True,
    encoding="utf-8",
    errors="replace"
)

if result.stdout:
    log("[STDOUT]")
    log(result.stdout)

if result.stderr:
    log("[STDERR]")
    log(result.stderr)

        elapsed = round(time.time() - start, 2)

        if result.returncode == 0:
            log(f"✅ COMPLETE : {script_file} / {elapsed} sec")
            return "OK"
        else:
            log(f"❌ FAILED : {script_file} / RETURN CODE {result.returncode}")
            return "FAILED"

    except Exception as e:
        log(f"❌ EXCEPTION : {script_file} / {e}")
        return "FAILED"


def print_result(results):
    log("#" * 80)
    log("GTI PIPELINE RESULT")
    log("#" * 80)

    for script_file, status in results:
        log(f"{script_file} : {status}")

    ok_count = sum(1 for _, status in results if status == "OK")
    skip_count = sum(1 for _, status in results if status == "SKIPPED")
    fail_count = sum(1 for _, status in results if status == "FAILED")

    log("-" * 80)
    log(f"OK      : {ok_count}")
    log(f"SKIPPED : {skip_count}")
    log(f"FAILED  : {fail_count}")
    log("#" * 80)


def main():
    ensure_dirs()

    log("#" * 80)
    log("GTI PIPELINE FINAL START")
    log("#" * 80)

    python_exe = get_python_exe()
    log(f"PYTHON : {python_exe}")
    log(f"BASE_DIR : {BASE_DIR}")

    archive_outputs()

    results = []

    for step_name, script_file, required in STEPS:
        status = run_script(step_name, script_file, required, python_exe)
        results.append((script_file, status))

        if required and status == "FAILED":
            log(f"🛑 PIPELINE STOPPED AT : {script_file}")
            break

    print_result(results)

    if any(status == "FAILED" for _, status in results):
        log("❌ GTI PIPELINE FINISHED WITH ERROR")
    else:
        log("✅ GTI PIPELINE FINISHED")


if __name__ == "__main__":
    main()
