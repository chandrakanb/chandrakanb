from jira import JIRA

# --- Configuration ---
JIRA_SERVER = "https://jira.honda-ivi.com"
# It is recommended to use an API Token instead of a password
JIRA_USER = "JIRAUSER21210"
JIRA_PASSWORD = "MDE3NTE2MTg1MjQ5Oqc4yGWxGSApiFdsaSE/F4GPXrwK"
ISSUE_KEY = "CNTKPIT-C1213"

# The new data extracted from your previous request
new_expected_result = "After step 1, CarPlay projection screen should be displayed"
new_actual_result = "After step 1, CarPlay projection screen is displayed"

def update_jira_test_step():
    try:
        # Connect to Jira
        options = {'server': JIRA_SERVER}
        jira = JIRA(options, basic_auth=(JIRA_USER, JIRA_PASSWORD))
        
        print(f"Connected to {JIRA_SERVER}. Fetching issue {ISSUE_KEY}...")
        issue = jira.issue(ISSUE_KEY)

        # NOTE: In Jira, Test Steps are often stored in Custom Fields.
        # You must find the custom field ID (e.g., 'customfield_12345') for your specific Jira setup.
        # If you are simply updating the Description, use 'description'.
        
        # Example 1: Updating the Description (Common for simple cases)
        # issue.update(description="Updated description text here")

        # Example 2: Updating a Custom Field (Common for Xray/Zephyr test steps)
        # Replace 'customfield_XXXXX' with your actual field ID
        update_fields = {
            'customfield_10101': new_expected_result, # Replace with actual Expected Result field ID
            'customfield_10102': new_actual_result    # Replace with actual Actual Result field ID
        }

        print("Updating issue...")
        issue.update(fields=update_fields)
        print(f"Successfully updated {ISSUE_KEY}!")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    update_jira_test_step()

# This is AI generated code, please refer KPIT AI Policy before using this in your projects