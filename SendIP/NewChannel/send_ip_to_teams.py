import socket
import requests
import time
import psutil

# Function to check if the network is connected
def is_network_connected():
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except OSError:
        return False

# Function to get all IPv4 addresses of the device
def get_all_ip_addresses():
    ip_addresses = []
    try:
        for iface, addrs in psutil.net_if_addrs().items():
            for addr in addrs:
                if addr.family == socket.AF_INET:  # IPv4 addresses
                    ip_addresses.append(addr.address)
    except Exception as e:
        return [f"Error retrieving IP addresses: {e}"]
    return ip_addresses

# Function to send the IP addresses to a Teams channel
def send_to_teams(ip_addresses, webhook_url):
    formatted_addresses = "\n\n".join(ip_addresses)  # Double newline for Teams formatting
    message = {
        "text": f"**The current IPv4 addresses of the local device:**\n\n{formatted_addresses}"
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
    TEAMS_WEBHOOK_URL = "https://kpitc.webhook.office.com/webhookb2/e86c9616-4701-4464-83fe-b6b554dbd88a@3539451e-b46e-4a26-a242-ff61502855c7/IncomingWebhook/4586a2b966e840b592e8dd4d9e2b395c/fbaa0227-8d16-4af9-aaf0-17c8889162fe/V2voshUJLDchVtG6oVoyU5nVcSLM9Rc6eGdbi-kQfZTmg1"

    print("Waiting for network connection...")
    while not is_network_connected():
        time.sleep(5)

    print("Network connected!")
    ip_addresses = get_all_ip_addresses()
    send_to_teams(ip_addresses, TEAMS_WEBHOOK_URL)
