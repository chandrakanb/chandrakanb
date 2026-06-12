import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import sys

def read_api_token(file_path):
    with open(file_path, "r") as file:
        return file.read().strip()

def post_comment_to_jira(jira_url, username, api_token, branch, test_script, issue_key, comment_text, max_lines):
    api_endpoint = f"{jira_url}/rest/api/3/issue/{issue_key}/comment"
    
    comment_lines = comment_text.split("\n")
    table_rows = []
    
    for i in range(0, len(comment_lines), max_lines):
        text_content = "\n".join(comment_lines[i:i+max_lines]).strip()
        table_rows.append({
            "type": "tableRow",
            "content": [
                {
                    "type": "tableCell",
                    "content": [
                        {
                            "type": "paragraph",
                            "content": [
                                {
                                    "type": "text",
                                    "text": text_content
                                }
                            ]
                        }
                    ]
                }
            ]
        })
    
    payload = {
        "body": {
            "type": "doc",
            "version": 1,
            "content": [
                {
                    "type": "table",
                    "content": table_rows
                }
            ]
        }
    }
    
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    response = requests.post(
        api_endpoint,
        headers=headers,
        json=payload,
        auth=HTTPBasicAuth(username, api_token)
    )
    
    if response.status_code == 201:
        print(f"Comment added successfully to Jira Ticket: {issue_key} for Test Script: {test_script} for {branch} Branch!")
    else:
        print(f"Failed to add comment to {issue_key}: {response.status_code} {response.text}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python script_name.py <excel_filename> <sheet_name>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    sheet_name = sys.argv[2]
    
    jira_url = "https://kpithondajapan.atlassian.net"
    username = "chandrakanb@kpit.com"
    token_path = "api_token.txt"
    
    api_token = read_api_token(token_path)
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    for _, row in df.iterrows():
        branch = row["Branch"] if pd.notna(row["Branch"]) else ""
        test_script = row["Test Script"] if pd.notna(row["Test Script"]) else ""
        issue_key = row["Issue key"] if pd.notna(row["Issue key"]) else ""
        comment_text = row["Comment"] if pd.notna(row["Comment"]) else ""
        max_lines = row["Max lines"] if pd.notna(row["Max lines"]) else ""
        post_comment_to_jira(jira_url, username, api_token, branch, test_script, issue_key, comment_text, max_lines)