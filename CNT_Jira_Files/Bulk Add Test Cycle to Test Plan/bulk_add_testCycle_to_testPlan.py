import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ==============================
# 🔹 CONFIGURATION
# ==============================
base_url = "https://jira.honda-ivi.com"
TOKEN = "MDE3NTE2MTg1MjQ5Oqc4yGWxGSApiFdsaSE/F4GPXrwK"   # 🔴 Replace for security
test_plan_key = "CNTKPIT-P12"

# API endpoint
link_api = f"{base_url}/rest/atm/1.0/testplan/{test_plan_key}/testcycle"

# Headers
headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json"
}

# ==============================
# 🔹 SESSION WITH RETRY
# ==============================
session = requests.Session()

retries = Retry(
    total=3,
    backoff_factor=1,
    status_forcelist=[429, 500, 502, 503, 504]
)

session.mount("https://", HTTPAdapter(max_retries=retries))

# ==============================
# 🔹 READ EXCEL
# ==============================
print("📄 Reading Excel file...")

try:
    df = pd.read_excel("cycles.xlsx", engine="openpyxl")
except Exception as e:
    print(f"❌ Failed to read Excel: {e}")
    exit()

if "cycleKey" not in df.columns:
    print("❌ Excel must contain 'cycleKey' column")
    exit()

# Clean and collect cycle keys
cycle_keys = (
    df["cycleKey"]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
    .tolist()
)

if not cycle_keys:
    print("❌ No valid cycle keys found")
    exit()

print(f"✅ Found {len(cycle_keys)} cycle(s)")

# ==============================
# 🔹 LINK TEST CYCLES
# ==============================
print("\n🔗 Linking test cycles to test plan...\n")

payload = {
    "testCycleKeys": cycle_keys   # ✅ Correct key (plural)
}

try:
    response = session.post(
        link_api,
        headers=headers,
        json=payload
    )

    print(f"📡 Status Code: {response.status_code}")

    if response.status_code in [200, 201]:
        print("✅ All test cycles linked successfully!")

    else:
        print("❌ Failed to link cycles")
        print("Response:")
        print(response.text)

        # Helpful hints
        if response.status_code == 404:
            print("\n⚠️ POSSIBLE ISSUE:")
            print("- Incorrect API endpoint")
            print("- Zephyr Scale API not installed")
            print("- Wrong Jira base URL")

        elif response.status_code == 401:
            print("\n⚠️ AUTH ERROR:")
            print("- Invalid / expired token")

        elif response.status_code == 403:
            print("\n⚠️ PERMISSION ERROR:")
            print("- No access to Test Plan or Test Cycles")

except Exception as e:
    print(f"❌ Request failed: {e}")

# ==============================
# 🔹 DONE
# ==============================
print("\n🏁 Script completed.")