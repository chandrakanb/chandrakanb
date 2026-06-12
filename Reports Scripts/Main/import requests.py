import requests

def send_table_to_teams(webhook_url, table_data, message_header="Execution Report Table"):
    """
    Sends a table to Microsoft Teams via webhook.

    Args:
        webhook_url (str): Teams webhook URL.
        table_data (str): Formatted table as a string.
        message_header (str): Header message for the table.
    """
    # Combine header and table
    message_body = f"**{message_header}**\n\n```\n{table_data}\n```"
    
    # Prepare payload
    payload = {
        "text": message_body
    }
    
    # Send the request to Teams
    response = requests.post(webhook_url, json=payload)
    
    if response.status_code == 200:
        print("Table sent to Teams successfully!")
    else:
        print(f"Failed to send table. Status code: {response.status_code}. Response: {response.text}")

# Example Usage
if __name__ == "__main__":
    # Teams webhook URL
    teams_webhook_url = 'https://kpitc.webhook.office.com/webhookb2/0e77225f-5e51-4ff3-bddd-87d2f5da5fc1@3539451e-b46e-4a26-a242-ff61502855c7/IncomingWebhook/3e007dda8f48458cadabf9daa593adfd/fbaa0227-8d16-4af9-aaf0-17c8889162fe/V2_GOSGEaXLJ7pFA9uEvjfOThvUNEcw-B_bfdz2Dd8-541'

    # Example table data
    table_string = [['Postcondition', '12', '0', '0', '12'],
        ['Precondition', '35', '0', '0', '35'],
        ['Test Script', '20', '2', '0', '22'],
        ['Total', '67', '2', '0', '69']]
    
    # Send the table
    send_table_to_teams(teams_webhook_url, table_string, "Execution Summary")
