import pandas as pd
import requests
import json
import re
import time
import os
from typing import Dict, Any, Optional

BASE_URL = "https://jira.honda-ivi.com/rest/atm/1.0/"
TOKEN = "OTI4ODIzNTkxMDU3OqXjifLIZb4PU1slc2DBCcLaeTM2"  # <-- Put your token here or read from env

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "content-type": "application/json"
}

SEARCH_PAGE_SIZE = 100
RESULTS_PAGE_SIZE = 200
SLEEP_BETWEEN_CALLS = 0.15  # polite delay to avoid rate limits

def urljoin(*parts: str) -> str:
    return "/".join(p.strip("/").rstrip("/") for p in parts if p)

def safe_filename(name: str, suffix: str = ".xlsx") -> str:
    """Windows-safe filename from arbitrary string."""
    cleaned = re.sub(r'[\\/:\*\?"<>\|]+', '_', name.strip())
    cleaned = re.sub(r"_+", "_", cleaned).strip(" .")
    cleaned = cleaned[:150]
    return f"{cleaned}{suffix}"

# ---------------------------
# API: Search test runs
# ---------------------------
def fetch_test_cycles_in_folder(project_key: str, folder_path: str):
    """
    Fetch test runs for a project and folder using GET /testrun/search?query=...
    If your tenant prefers POST, we can switch to POST with JSON body.
    """
    query = f'projectKey = "{project_key}" AND folder = "{folder_path}"'
    search_url = f"{BASE_URL}testrun/search"

    all_runs = []
    start_at = 0

    while True:
        params = {"query": query, "startAt": start_at, "maxResults": SEARCH_PAGE_SIZE}
        try:
            resp = requests.get(search_url, headers=HEADERS, params=params, timeout=60)
            resp.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"Error fetching test cycles: {e}")
            return []

        data = resp.json()

        # Normalize potential shapes
        if isinstance(data, dict):
            values = data.get("values") or data.get("items") or data.get("results") or []
            total = data.get("total") or data.get("totalResults")
        elif isinstance(data, list):
            values, total = data, None
        else:
            values, total = [], None

        all_runs.extend(values)

        if total is not None:
            start_at += len(values)
            if start_at >= int(total):
                break
        else:
            if len(values) < SEARCH_PAGE_SIZE:
                break
            start_at += len(values)

    return all_runs

# ---------------------------
# API: Test results for a run
# ---------------------------
def iter_test_results_for_run(test_run_key: str):
    """
    Iterate all test execution results for a given test run.
    GET /testrun/{testRunKey}/testresults
    """
    base = urljoin(BASE_URL, "testrun", test_run_key, "testresults")
    start_at = 0

    while True:
        params = {"startAt": start_at, "maxResults": RESULTS_PAGE_SIZE}
        try:
            resp = requests.get(base, headers=HEADERS, params=params, timeout=60)
            resp.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"[testresults] {test_run_key} error: {e}")
            break

        data = resp.json()

        if isinstance(data, dict):
            results = data.get("values") or data.get("items") or data.get("results") or data.get("testResults") or []
            total = data.get("total") or data.get("totalResults")
        elif isinstance(data, list):
            results, total = data, None
        else:
            results, total = [], None

        if not results:
            break

        for r in results:
            yield r

        if total is not None:
            start_at += len(results)
            if start_at >= int(total):
                break
        else:
            if len(results) < RESULTS_PAGE_SIZE:
                break
            start_at += len(results)

        time.sleep(SLEEP_BETWEEN_CALLS)

# ---------------------------
# API: Fetch test case name
# ---------------------------
def get_test_case_name(test_case_key: str, cache: Dict[str, str]) -> Optional[str]:
    """
    Returns the test case name from cache or by calling GET /testcase/{testCaseKey}.
    Caches results to avoid repeated calls.
    """
    if not test_case_key:
        return None
    if test_case_key in cache:
        return cache[test_case_key]

    url = urljoin(BASE_URL, "testcase", test_case_key)
    try:
        resp = requests.get(url, headers=HEADERS, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        # Common shapes: {"key": "...", "name": "..."} or deeper nesting
        name = data.get("name") or data.get("testCase", {}).get("name")
        if name:
            cache[test_case_key] = name
            return name
    except requests.exceptions.RequestException as e:
        print(f"[testcase] {test_case_key} error: {e}")
    return None

# ---------------------------
# Flatten row
# ---------------------------
def flatten_result(run: Dict[str, Any], result: Dict[str, Any], tc_name_cache: Dict[str, str]) -> Dict[str, Any]:
    run_key   = run.get("key") or run.get("testRunKey")
    run_name  = run.get("name")  # <-- Test Cycle Name

    # Pull keys & names
    tc_obj    = result.get("testCase") or {}
    test_case_key = result.get("testCaseKey") or tc_obj.get("key")
    test_case_name = tc_obj.get("name")

    # If name missing, fetch via /testcase/{key} and cache
    if not test_case_name and test_case_key:
        test_case_name = get_test_case_name(test_case_key, tc_name_cache)

    row = {
        "testRunKey": run_key,
        "TestCycleName": run_name,        # <-- added explicit column
        "testCaseKey": test_case_key,
        "testCaseName": test_case_name,
        "status": result.get("status"),
        "comment": result.get("comment"),
        "executedBy": result.get("executedBy") or result.get("executedByUser") or result.get("executedById"),
        "executedOn": result.get("executedOn") or result.get("executionDate"),
        "assignedTo": result.get("assignedTo") or result.get("assignee"),
        "environment": result.get("environment"),
        "Variant Code": json.dumps(result.get("Variant Code")),
        "Execution Method": json.dumps(result.get("Execution Method")),
        "executionTime": result.get("executionTime") or result.get("duration"),
        "defects_json": json.dumps(result.get("defects") or result.get("issueLinks") or [], ensure_ascii=False),
        "evidences_json": json.dumps(result.get("evidences") or result.get("attachments") or [], ensure_ascii=False),
        "raw_result_json": json.dumps(result, ensure_ascii=False)
    }
    return row

# ---------------------------
# Main
# ---------------------------
def main():
    PROJECT_KEY = "CNTKPIT"
    FOLDER_PATH = "/PF 26/A15/BFA1509R0.260603.1/"
    #FOLDER_PATH = "/PF 26/A15/BFA1509R0.260603.1/iPhone 16 Pro/"
    #FOLDER_PATH = "/PF 26/A15/BFA1509R0.260603.1/Google Pixel 9/"
    #FOLDER_PATH = "/PF 26/A15/BFA1509R0.260603.1/Samsung S25 Ultra"

    print(f"Fetching test runs from folder: '{FOLDER_PATH}'...")
    runs = fetch_test_cycles_in_folder(PROJECT_KEY, FOLDER_PATH)
    if not runs:
        print("No test runs found or an error occurred. Exiting.")
        return

    print(f"Found {len(runs)} run(s). Fetching test results...")
    rows = []
    total_results = 0
    tc_name_cache: Dict[str, str] = {}

    for i, run in enumerate(runs, start=1):
        run_key = run.get("key")
        run_name = run.get("name")
        print(f"[{i}/{len(runs)}] {run_key} | {run_name}")

        count = 0
        for res in iter_test_results_for_run(run_key):
            rows.append(flatten_result(run, res, tc_name_cache))
            count += 1
            total_results += 1

        print(f"   -> {count} result(s)")
        time.sleep(SLEEP_BETWEEN_CALLS)

    if not rows:
        print("No test results to export.")
        return

    df = pd.DataFrame(rows)

    # Optional: enforce a clear column order (with Test Cycle Name up front)
    preferred_cols = [
        "testRunKey", "TestCycleName",
        "testCaseKey", "testCaseName",
        "status", "comment",
        "executedBy", "executedOn", "assignedTo", "environment", "Variant Code", "Execution Method", "executionTime",
        "defects_json", "evidences_json", "raw_result_json"
    ]
    # Keep any extra columns at the end (if API adds fields later)
    cols = [c for c in preferred_cols if c in df.columns] + [c for c in df.columns if c not in preferred_cols]
    df = df[cols]

    out_name = safe_filename(f"Zephyr_{FOLDER_PATH}_dump")
    out_path = os.path.abspath(out_name)

    df.to_excel(out_path, index=False, engine="openpyxl")
    print(f"\nSaved {total_results} test result(s) to: {out_path}")

if __name__ == "__main__":
    main()