from jira import JIRA
import os

def add_jira_comment(server_url, email, api_token, issue_key, comment_text):
    """
    Adds a comment to a specific Jira ticket.
    
    :param server_url: The base URL of your Jira instance (e.g., https://yourcompany.atlassian.net)
    :param email: The email address associated with your Jira account
    :param api_token: Your Jira API token
    :param issue_key: The ID of the ticket (e.g., 'PROJ-123')
    :param comment_text: The text content of the comment
    """
    try:
        # Authenticate with Jira
        jira_options = {'server': server_url}
        jira = JIRA(options=jira_options, basic_auth=(email, api_token))

        # Fetch the issue
        issue = jira.issue(issue_key)

        # Add the comment
        jira.add_comment(issue, comment_text)
        
        print(f"Successfully added comment to {issue_key}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # --- CONFIGURATION ---
    # Replace these values with your actual Jira credentials
    
    JIRA_SERVER = "https://kpithondajapan.atlassian.net"
    USER_EMAIL = "chandrakanb@kpit.com"

    API_TOKEN = os.getenv("JIRA_API_TOKEN")    

    COMMENT = "As discussed with [Saurabh Arvind Kulkarni|~accountid:712020:b8695750-73a1-4b24-9a64-4b6acf7ccc2d] san, Assigning this ticket to Current Bench Owner([Sayali Raut|~accountid:712020:1dbc52f4-7b49-4147-8e25-bc64c9ef4b37] ), Reports were attached by [Pratham Chaturvedi|~accountid:712020:2510e08b-eecd-4a64-85cc-8b9a4a15e0d9], Hence assign this to him post fixation."
    # ---------------------

    List = ["OSV-2396", "OSV-2345", "OSV-2344", "OSV-2343", "OSV-2342", "OSV-2341", "OSV-2339", "OSV-2338", "OSV-1964", "OSV-1685", "OSV-1682", "OSV-1677", "OSV-1676", "OSV-1675", "OSV-1674", "OSV-1673", "OSV-1670"]

    for TICKET_ID in List:
        
        add_jira_comment(JIRA_SERVER, USER_EMAIL, API_TOKEN, TICKET_ID, COMMENT)