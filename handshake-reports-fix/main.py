"""Read emails from Outlook."""

import os

import win32com.client
from dotenv import load_dotenv


def parse_env_vars() -> None:
    """Parse configuration from environment variables and `.env` file."""
    # Initialize the `python-dotenv` ".env" file, if present...
    load_dotenv()
    # and read the environment variables:

    # 1. `chdir` to the network path specified by `HANDSHAKE_NETWORK_PATH`, if present.
    if network_path := os.getenv("HANDSHAKE_NETWORK_PATH"):
        # Change the current working directory to the network path specified by the
        # variable `HANDSHAKE_NETWORK_PATH` in the .env file, if present.
        os.chdir(network_path)
    # 2. Handle other environment variables in the future?


def load_emails() -> None:
    """Read emails from the "handshake-reports" folder in Outlook."""
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Access the "Handshake-Reports" folder
    root_folder = outlook.Folders[2]
    handshake_reports_folder = root_folder.Folders["Handshake-Reports"]

    # Get the items (emails) in the folder
    messages = handshake_reports_folder.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by received time (newest first)

    # Process the emails
    for i, message in enumerate(messages):
        # 1. Find messages in the right folder with the right subject.
        if "BILAKE" not in message.Subject or "error" in message.Subject.casefold():
            # Wrong subject. Skip this email
            continue
        print(f"{i}. Subject: {message.Subject}, Received: {message.ReceivedTime}")

        # 2. Download attachments.
        # 3. Mark the email as read.
        # 4. Find emails older than $n$ days. Ensure their attachments have been
        #    downloaded, then delete the email.


if __name__ == "__main__":
    parse_env_vars()
    load_emails()
