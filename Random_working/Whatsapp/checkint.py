import urllib.request

try:
    urllib.request.urlopen("https://www.google.com", timeout=5)
    print("Internet connection is active.")
except Exception as e:
    print("No internet connection:", e)
