"""Read emails from Outlook."""

import os

import requests
from dotenv import load_dotenv
from msal import PublicClientApplication

load_dotenv()

# App Registration Details
CLIENT_ID = os.getenv("MSFT_CLIENT_ID")
TENANT_ID = os.getenv("MSFT_TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.ReadWrite"]  # Adjust based on your needs

# Authenticate using interactive login
app = PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)
result = app.acquire_token_interactive(scopes=SCOPES)

if "access_token" in result:
    access_token = result["access_token"]
    headers = {"Authorization": f"Bearer {access_token}"}

    # Fetch the latest 10 emails
    response = requests.get(
        "https://graph.microsoft.com/v1.0/me/messages",
        headers=headers,
        params={"$top": 10},  # Fetch the latest 10 emails
        timeout=100,
    )

    if response.status_code == requests.codes.ok:
        emails = response.json().get("value", [])
        for email in emails:
            print(f"Subject: {email['subject']}, Received: {email['receivedDateTime']}")
    else:
        print(f"Error fetching emails: {response.status_code} - {response.text}")
else:
    print("Failed to acquire token.")
