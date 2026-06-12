from jira import JIRA

# Configuration
JIRA_URL = 'https://kpithondajapan.atlassian.net'
USERNAME = 'chandrakanb@kpit.com'
API_TOKEN = 'ATATT3xFfGF00OVb1BHWgz1klgLPAPkq8W8nFtC2j6k_ZYhvBxAWBFwEnCRk18fPD1ctA4TWpWJ79K4s4OPXJOiskD0GInEa14A0ex8puHBIQRLaGOsmNkVRY4ij9rKJr1q2JgK9D4zR-nK0rpl4QevUi5_g2PWiSM5HB7qEZAOMUaWUirk1vD4=54745E1F'
JQL_QUERY = 'issuetype = Overnight-Issues AND labels = Dev_2025-04-22 AND "Bench Number[Checkboxes]" = "Bench07_HATC(T)_DEV" order by created DESC'
LABELS_TO_REMOVE = {'Dev_2025-04-22'}

# Connect to Jira
jira = JIRA(server=JIRA_URL, basic_auth=(USERNAME, API_TOKEN))

# Fetch matching issues
issues = jira.search_issues(JQL_QUERY, maxResults=False)

for issue in issues:
    current_labels = set(issue.fields.labels)
    print(f"\n📌 Issue: {issue.key}")
    print(f"   Current Labels: {current_labels}")

    labels_to_be_removed = current_labels & LABELS_TO_REMOVE
    if not labels_to_be_removed:
        print("   No matching labels to remove. Skipping.")
        continue

    confirm = input(f"   Remove labels {labels_to_be_removed} from {issue.key}? (y/n): ").strip().lower()
    if confirm == 'y':
        updated_labels = list(current_labels - labels_to_be_removed)
        issue.update(fields={'labels': updated_labels})
        print(f"   ✅ Removed labels from {issue.key}.")
    else:
        print(f"   ❌ Skipped {issue.key}.")

print("\n✔️ Finished processing all issues.")
