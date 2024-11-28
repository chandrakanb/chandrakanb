import requests
import json

# Teams webhook URL
#webhook_url = "https://kpitc.webhook.office.com/webhookb2/9269d31d-1feb-4fdc-9969-0ab2da2adcc7@3539451e-b46e-4a26-a242-ff61502855c7/IncomingWebhook/d06a3aa25c674c12ae21c2cf796762f9/85ba79cd-deb2-408d-8c24-664ce7ccb367"

webhook_url = 'https://kpitc.webhook.office.com/webhookb2/0e77225f-5e51-4ff3-bddd-87d2f5da5fc1@3539451e-b46e-4a26-a242-ff61502855c7/IncomingWebhook/bb555a09021747a696f5702d10233e92/fbaa0227-8d16-4af9-aaf0-17c8889162fe'


# Message to send
message = {
    "text": "https://www.youtube.com/watch?v=3NFVuyGwb3Y"
}

# Send the message to Teams
response = requests.post(
    webhook_url, 
    headers={"Content-Type": "application/json"},
    data=json.dumps(message)
)

# Check for a successful response
if response.status_code == 200:
    print("Message sent successfully!")
else:
    print(f"Failed to send message. Status code: {response.status_code}")
