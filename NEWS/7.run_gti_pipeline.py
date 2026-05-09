# 7.run_gti_pipeline.py
# -*- coding: utf-8 -*-

import os
import sys
import time
import shutil
import subprocess
from datetime import datetime

BASE_DIR = r"C:\temp"

STEP_FILES = [
    ("STEP1 site crawler", os.path.join(BASE_DIR, "1.site_crawler.py"), True),
    ("STEP2-1 naver news", os.path.join(BASE_DIR, "2-1.naver_news_collector.py"), True),
    ("STEP2-2 google news", os.path.join(BASE_DIR, "2-2.google_news_collector.py"), True),

    # 실제 사용자 파일명 기준 수정
    ("STEP2-3 rss news", os.path.join(BASE_DIR, "2-3._rss_news_raw.py"), True),

    ("STEP3 merge", os.path.join(BASE_DIR, "3.merge_news.py"), True),
    ("STEP4 analyze", os.path.join(BASE_DIR, "4.policy_ai_analyzer.py"), True),
    ("STEP5 mail", os.path.join(BASE_DIR, "5.GTI Mail Engine.py"), True),

    # 선택 실행
    ("STEP6 dashboard", os.path.join(BASE_DIR, "dashboard_generator.py"), False),
]

EXPECTED_OUTPUTS = {
    "STEP1 site crawler": ["1.site_news_raw.xlsx"],
    "STEP2-1 naver news": ["2-1.naver_news_raw.xlsx"],
    "STEP2-2 google news": ["2-2.google_news_raw.xlsx"],
    "STEP2-3 rss news": ["2-3.rss_news_raw.xlsx"],
    "STEP3 merge": ["news_master_raw.xlsx"],
    "STEP4 analyze": ["news_raw.xlsx"],
    "STEP5 mail": [],
    "STEP6 dashboard": [],
}

LOG_DIR = os.path.join(BASE_DIR, "logs")
ARCHIVE_DIR = os.path.join(BASE_DIR, "archive")
RUN_LOG_FILE = os.path.join(LOG_DIR, "gti_pipeline_run.log")


def ensure_dirs():
    os.makedirs(BASE_DIR, exist_ok=True)
    os.makedirs(LOG_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_DIR, exist_ok=True)


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def write_log(msg=""):
    line = f"[{now_str()}] {msg}"
    print(line)
    with open(RUN_LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def detect_python():
    candidates = [
        r"C:\Users\KCH\AppData\Local\Programs\Python\Python312\python.exe",
        sys.executable,
        shutil.which("python"),
        shutil.which("py"),
    ]

    for c in candidates:
        if c and os.path.exists(c):
            return c

    return sys.executable


def check_required_files():
    missing_required = []
    missing_optional = []

    for step_name, script_path, required in STEP_FILES:
        if not os.path.exists(script_path):
            if required:
                missing_required.append((step_name, script_path))
            else:
                missing_optional.append((step_name, script_path))

    return missing_required, missing_optional


def archive_existing_outputs():
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_archive_dir = os.path.join(ARCHIVE_DIR, stamp)
    os.makedirs(run_archive_dir, exist_ok=True)

    targets = [
        "1.site_news_raw.xlsx",
        "2-1.naver_news_raw.xlsx",
        "2-2.google_news_raw.xlsx",
        "2-3.rss_news_raw.xlsx",
        "news_master_raw.xlsx",
        "news_raw.xlsx",
        "news_cumulative.xlsx",
        "mail_cumulative.xlsx",
    ]

    copied = 0

    for name in targets:
        src = os.path.join(BASE_DIR, name)
        if os.path.exists(src):
            try:
                shutil.copy2(src, os.path.join(run_archive_dir, name))
                write_log(f"ARCHIVE OK: {name}")
                copied += 1
            except Exception as e:
                write_log(f"ARCHIVE FAIL: {name} / {e}")

    write_log(f"ARCHIVE COMPLETE: {copied} files")


def check_outputs(step_name):
    expected = EXPECTED_OUTPUTS.get(step_name, [])
    if not expected:
        return True

    ok = True
    for filename in expected:
        path = os.path.join(BASE_DIR, filename)
        if os.path.exists(path):
            size = os.path.getsize(path)
            write_log(f"OUTPUT OK: {filename} / {size:,} bytes")
        else:
            write_log(f"OUTPUT MISSING: {filename}")
            ok = False

    return ok


def run_step(step_name, script_path, python_exe):
    write_log("=" * 70)
    write_log(f"START: {step_name}")
    write_log(f"SCRIPT: {script_path}")

    if not os.path.exists(script_path):
        write_log(f"SKIP: script not found")
        return "SKIP"

    start = time.time()

    try:
        result = subprocess.run(
            [python_exe, script_path],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        elapsed = round(time.time() - start, 2)

        write_log(f"RETURN CODE: {result.returncode}")
        write_log(f"ELAPSED: {elapsed} sec")

        if result.stdout:
            for line in result.stdout.splitlines():
                write_log(f"STDOUT: {line}")

        if result.stderr:
            for line in result.stderr.splitlines():
                write_log(f"STDERR: {line}")

        if result.returncode != 0:
            return "FAIL"

        output_ok = check_outputs(step_name)
        if not output_ok:
            return "FAIL_OUTPUT_MISSING"

        return "SUCCESS"

    except Exception as e:
        write_log(f"EXCEPTION: {step_name} / {e}")
        return "FAIL"


def print_summary(results):
    write_log("=" * 70)
    write_log("GTI PIPELINE SUMMARY")

    success = 0
    fail = 0
    skip = 0

    for step_name, status in results:
        write_log(f"{step_name}: {status}")

        if status == "SUCCESS":
            success += 1
        elif status == "SKIP":
            skip += 1
        else:
            fail += 1

    write_log("-" * 70)
    write_log(f"SUCCESS: {success}")
    write_log(f"FAIL   : {fail}")
    write_log(f"SKIP   : {skip}")
    write_log("=" * 70)


def main():
    ensure_dirs()

    write_log()
    write_log("#" * 70)
    write_log("GTI PIPELINE START")
    write_log("#" * 70)

    missing_required, missing_optional = check_required_files()

    if missing_optional:
        write_log("OPTIONAL SCRIPT FILES MISSING")
        for step_name, script_path in missing_optional:
            write_log(f"OPTIONAL MISSING: {step_name} / {script_path}")

    if missing_required:
        write_log("REQUIRED SCRIPT FILES MISSING")
        for step_name, script_path in missing_required:
            write_log(f"MISSING: {step_name} / {script_path}")
        write_log("GTI PIPELINE STOPPED")
        return

    python_exe = detect_python()
    write_log(f"PYTHON: {python_exe}")

    archive_existing_outputs()

    results = []

    for step_name, script_path, required in STEP_FILES:
        status = run_step(step_name, script_path, python_exe)
        results.append((step_name, status))

        if required and status not in ["SUCCESS"]:
            write_log(f"PIPELINE STOPPED AT: {step_name}")
            break

    print_summary(results)
    write_log("GTI PIPELINE END")


if __name__ == "__main__":
    main()