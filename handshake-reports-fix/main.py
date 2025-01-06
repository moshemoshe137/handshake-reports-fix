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

    print("Latest emails in handshake-reports:")
    for i, message in enumerate(messages, start=1):
        print(f"{i}. Subject: {message.Subject}, Received: {message.ReceivedTime}")
        if i >= 20:  # Stop after the first 20 emails  # noqa: PLR2004
            break


if __name__ == "__main__":
    parse_env_vars()
    load_emails()
