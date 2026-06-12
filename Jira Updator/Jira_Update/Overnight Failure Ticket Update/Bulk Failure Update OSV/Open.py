import webbrowser

# JIRA base URL
jira_base_url = "https://kpithondajapan.atlassian.net/browse/"

# List of Jira IDs
jira_ids = [
    "OSV-2489", "OSV-2612", "OSV-2487", "OSV-2610", "OSV-2607", "OSV-2609",
    "OSV-2494", "OSV-2611", "OSV-2491", "OSV-2485", "OSV-2613", "OSV-2608",
    "OSV-2496", "OSV-2479", "OSV-2480", "OSV-2481", "OSV-2484", "OSV-2495",
    "OSV-2483", "OSV-2493", "OSV-2492", "OSV-2478", "OSV-2490", "OSV-2482", "OSV-2488"
]

# Open each Jira ticket in the default web browser
for ticket_id in jira_ids:
    url = f"{jira_base_url}{ticket_id}"
    webbrowser.open_new_tab(url)
