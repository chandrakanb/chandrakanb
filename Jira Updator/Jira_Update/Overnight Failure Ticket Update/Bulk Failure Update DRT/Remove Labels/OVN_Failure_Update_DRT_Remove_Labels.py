import os
import sys
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from jira import JIRA
from datetime import datetime
import openpyxl

# --- CONFIG ---
excel_path = "input.xlsx"
sheet_name = "Sheet1"
jira_server = "https://kpithondajapan.atlassian.net"
username = "chandrakanb@kpit.com"

# Read the actual API token from the file
with open("api_token.txt", "r") as f:
    JIRA_API_TOKEN = f.read().strip()

# --- CONNECT TO JIRA ---
jira = JIRA(server=jira_server, basic_auth=(username, JIRA_API_TOKEN))

def get_labels_for_jira_ids(excel_path, sheet_name, label_to_remove):
    """
    Reads Jira issue IDs from an Excel file and fetches their labels.  Removes the specified
    label from each issue and updates Jira.

    Args:
        excel_path (str): The path to the Excel file.
        sheet_name (str): The name of the worksheet containing the Jira IDs.
        label_to_remove (str): The label to remove from the issue.

    Returns:
        dict: A dictionary where keys are Jira issue IDs and values are lists of labels.
              Returns an empty dictionary if an error occurs.
    """
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook[sheet_name]

        header_row = [cell.value for cell in sheet[1]]  # Get header row values
        issue_id_labels = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            row_dict = dict(zip(header_row, row))  # Create a dictionary for the row
            issue_id = row_dict['Jira ID']  # Access by column name
            #test_script_id = row_dict['Test Script ID'] #Unused

            try:
                issue = jira.issue(issue_id)
                existing_labels = issue.fields.labels

                # Remove the specified label
                if label_to_remove in existing_labels:
                    updated_labels = [label for label in existing_labels if label != label_to_remove]
                    print("Updated labels (before API call):", updated_labels)  # Print the updated labels

                    # Update the Jira issue with the new labels *only if labels changed*
                    if updated_labels != existing_labels:
                        issue.update(fields={"labels": updated_labels})
                        print(f"Successfully removed '{label_to_remove}' from issue {issue_id} and updated labels.")
                    else:
                        print(f"Label '{label_to_remove}' found on issue {issue_id}, but no changes were made because the labels are identical.")

                else:
                    print(f"Label '{label_to_remove}' not found on issue {issue_id}. No changes made.")

                issue_id_labels[issue_id] = existing_labels  # Store original labels for reporting

            except Exception as e:
                print(f"Error fetching/updating labels for issue {issue_id}: {e}")
                issue_id_labels[issue_id] = []  # Store an empty list if labels couldn't be fetched

        return issue_id_labels

    except FileNotFoundError:
        print(f"Error: Excel file '{excel_path}' not found.")
        return {}
    except KeyError as e:
        print(f"Error: Column '{e}' not found in Excel file.")  # More specific error message
        return {}
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return {}

if len(sys.argv) != 2:
        print("Usage: python script.py <label_to_remove>")
        sys.exit(1)

label_to_remove = sys.argv[1]

# Example Usage
issue_labels = get_labels_for_jira_ids(excel_path, sheet_name, label_to_remove)

if issue_labels:
    for issue_id, labels in issue_labels.items():
        print(f"Issue {issue_id}: {labels}")
else:
    print("No labels fetched or an error occurred.")