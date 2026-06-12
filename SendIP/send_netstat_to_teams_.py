import socket
import requests
import time
import subprocess

# Function to check if the network is connected
def is_network_connected():
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except OSError:
        return False

# Function to get netstat output filtered for port 5900
def get_netstat_5900():
    try:
        result = subprocess.run(["netstat", "-an"], capture_output=True, text=True)
        lines = [line.strip() for line in result.stdout.split("\n") if "5900" in line]
        return lines if lines else ["No active connections on port 5900."]
    except Exception as e:
        return [f"Error retrieving netstat output: {e}"]

# Function to send data to a Teams channel
def send_to_teams(data, webhook_url):
    formatted_data = "\n\n".join(data)  # Double newline for Teams message formatting
    message = {
        "text": f"**Netstat output for port 5900:**\n\n{formatted_data}"
    }
    try:
        response = requests.post(webhook_url, json=message)
        if response.status_code == 200:
            print("Message successfully sent to Teams!")
        else:
            print(f"Failed to send message. Status code: {response.status_code}, Response: {response.text}")
    except Exception as e:
        print(f"Error sending message to Teams: {e}")

# Main logic
if __name__ == "__main__":
    TEAMS_WEBHOOK_URL = "https://kpitc.webhook.office.com/webhookb2/0e77225f-5e51-4ff3-bddd-87d2f5da5fc1@3539451e-b46e-4a26-a242-ff61502855c7/IncomingWebhook/b8a0c20d149c48048d0d21f6535a4934/fbaa0227-8d16-4af9-aaf0-17c8889162fe/V2V6ipm5mIJaAiCKhkix0ToqOMqYdyyhxpHneCvxmxNbo1"

    print("Waiting for network connection...")
    while not is_network_connected():
        time.sleep(5)

    print("Network connected!")
    netstat_output = get_netstat_5900()
    send_to_teams(netstat_output, TEAMS_WEBHOOK_URL)
