from jira import JIRA

# --- CONFIGURATION ---
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"
password = "ATATT3xFfGF00OVb1BHWgz1klgLPAPkq8W8nFtC2j6k_ZYhvBxAWBFwEnCRk18fPD1ctA4TWpWJ79K4s4OPXJOiskD0GInEa14A0ex8puHBIQRLaGOsmNkVRY4ij9rKJr1q2JgK9D4zR-nK0rpl4QevUi5_g2PWiSM5HB7qEZAOMUaWUirk1vD4=54745E1F"
ticket_key = "DRT-3948"        # Your Jira ticket
#target_transition = "Done"        # Name of the transition to perform

# --- CONNECT TO JIRA ---
jira = JIRA(server=jira_server, basic_auth=(username, password))

# --- GET AVAILABLE TRANSITIONS ---
transitions = jira.transitions(ticket_key)
print(f"Available transitions for {ticket_key}:\n")
for t in transitions:
    print(f"- {t['name']} (ID: {t['id']})")

"""
# --- FIND TRANSITION ID ---
transition_id = None
for t in transitions:
    if t['name'].lower() == target_transition.lower():
        transition_id = t['id']
        break

# --- APPLY TRANSITION ---
if transition_id:
    jira.transition_issue(ticket_key, transition_id)
    print(f"Issue {ticket_key} transitioned to '{target_transition}'.")
else:
    print(f"Transition '{target_transition}' not found for {ticket_key}.")
"""