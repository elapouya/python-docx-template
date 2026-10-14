#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""CLI coverage for --validate / --report exit codes."""

import json
import os
import subprocess
import sys
import tempfile


def run_cli(*args):
    return subprocess.run(
        [sys.executable, "-m", "docxtpl"] + list(args),
        capture_output=True,
        text=True,
    )


def check(name, result, expected_code):
    ok = result.returncode == expected_code
    status = "PASS" if ok else "FAIL"
    print(
        "%s: %s (exit %s, expected %s)"
        % (status, name, result.returncode, expected_code)
    )
    if not ok and result.stdout:
        print("  stdout: %s" % result.stdout.strip())
    if not ok and result.stderr:
        print("  stderr: %s" % result.stderr.strip())
    return ok


def main():
    os.chdir(os.path.dirname(__file__))
    print("Testing CLI --validate / --report ...")
    print("=" * 50)

    results = []

    results.append(
        check(
            "valid template + --validate",
            run_cli("templates/module_execute_tpl.docx", "--validate", "-q"),
            0,
        )
    )

    results.append(
        check(
            "invalid Jinja template + --validate",
            run_cli("templates/template_error_tpl.docx", "--validate", "-q"),
            1,
        )
    )

    results.append(
        check(
            "missing template + --validate",
            run_cli("missing.docx", "--validate", "-q"),
            1,
        )
    )

    report_path = os.path.join(tempfile.gettempdir(), "docxtpl_cli_validate_report.json")
    results.append(
        check(
            "--report without --validate",
            run_cli("templates/module_execute_tpl.docx", "--report", report_path),
            1,
        )
    )

    # Light sanity check: valid template writes {"valid": true}
    report_path = os.path.join(
        tempfile.gettempdir(), "docxtpl_cli_validate_ok_report.json"
    )
    if os.path.exists(report_path):
        os.unlink(report_path)
    result = run_cli(
        "templates/module_execute_tpl.docx",
        "--validate",
        "--report",
        report_path,
        "-q",
    )
    report_ok = False
    if result.returncode == 0 and os.path.isfile(report_path):
        with open(report_path) as f:
            payload = json.load(f)
        report_ok = payload == {"valid": True}
    status = "PASS" if report_ok else "FAIL"
    print("%s: valid template --report JSON" % status)
    results.append(report_ok)

    print("=" * 50)
    if all(results):
        print("ALL TESTS PASSED!")
        return 0
    print("SOME TESTS FAILED!")
    return 1


if __name__ == "__main__":
    sys.exit(main())
