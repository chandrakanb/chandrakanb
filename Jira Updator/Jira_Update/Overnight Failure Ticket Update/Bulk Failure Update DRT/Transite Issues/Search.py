import requests
from requests.auth import HTTPBasicAuth

# ================= CONFIG =================
SERVER_URL = "https://kpithondajapan.atlassian.net"
EMAIL = "chandrakanb@kpit.com"
API_TOKEN = "ATATT3xFfGF0ov91hsvxUfwrLyB2M-bUttAuaA3wjjindxG1NlvCmYfCsY4KWj-my90Wmig_q6uWhHLDpIEFlqbvBpyh8_l8Q8kfthpL4g3-YVFtbNyByNdgTmEZHGETM40vdjZr7RGpurLRGMtO4wMkXjG4z8m9f5PWksr5l-90w7Gh3NDoaIw=3705F98B"

JQL_QUERY = 'project = DRT AND type = Overnight-Issues AND "tl name[people]" = 712020:7ebd6444-898e-4d1e-ba7a-2710a262ffa7 ORDER BY created DESC'

# ================= REQUEST =================
url = f"{SERVER_URL}/rest/api/3/search/jql"

headers = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}

payload = {
    "jql": JQL_QUERY,
    "maxResults": 50,
    "fields": ["summary", "status", "assignee"]
}

response = requests.post(
    url,
    headers=headers,
    json=payload,
    auth=HTTPBasicAuth(EMAIL, API_TOKEN)
)

# ================= RESPONSE =================
if response.status_code == 200:
    data = response.json()
    issues = data.get("issues", [])

    print(f"Total issues found: {len(issues)}\n")

    for issue in issues:
        key = issue["key"]
        summary = issue["fields"]["summary"]
        status = issue["fields"]["status"]["name"]
        assignee = issue["fields"]["assignee"]
        assignee_name = assignee["displayName"] if assignee else "Unassigned"

        print(f"{key} | {status} | {assignee_name} | {summary}")
else:
    print("Failed to fetch issues")
    print(response.status_code)
    print(response.text)
