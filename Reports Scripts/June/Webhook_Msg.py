import requests
import json

def send_to_teams(webhook_url, message):
    headers = {
        "Content-Type": "application/json"
    }
    payload = {
        "text": message
    }
    response = requests.post(webhook_url, headers=headers, data=json.dumps(payload))
    if response.status_code == 200:
        print("Message sent successfully")
    else:
        print(f"Failed to send message: {response.status_code}")

if __name__ == "__main__":
    # Replace with your actual Teams webhook URL
    YOUR_TEAMS_WEBHOOK_URL = "https://kpitc.webhook.office.com/webhookb2/0e77225f-5e51-4ff3-bddd-87d2f5da5fc1@3539451e-b46e-4a26-a242-ff61502855c7/IncomingWebhook/bb555a09021747a696f5702d10233e92/fbaa0227-8d16-4af9-aaf0-17c8889162fe"
    
    # Replace with your message
    message = "Hello, this is a test message from the Python script!"
    
    # Send the message to Teams
    send_to_teams(YOUR_TEAMS_WEBHOOK_URL, message)
