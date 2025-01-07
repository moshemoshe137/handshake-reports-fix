"""Read emails from Outlook."""

import os
from pathlib import Path

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

    # Determine the Current User's email address
    users_email = (
        # Resolve the exchange email address
        outlook.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        # if they are an Exchange user.
        if outlook.CurrentUser.AddressEntry.GetExchangeUser()
        # Otherwise, use the email address from the `CurrentUser` object.
        else outlook.CurrentUser.Address
    )

    # Access the "Handshake-Reports" folder
    root_folder = outlook.Folders[users_email]
    handshake_reports_folder = root_folder.Folders["Handshake-Reports"]

    # Get the items (emails) in the folder
    messages = handshake_reports_folder.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by received time (newest first)

    # Process the emails
    for i, message in enumerate(messages):
        # 1. Find unread messages with the right subject.
        if (
            "BILAKE" not in message.Subject  # If the subject doesn't contain "BILAKE"
            or "error" in message.Subject.casefold()  # or if contains "error"
            or not message.UnRead  # or if it's already read
        ):
            # Skip this email! It's got the wrong subject or is already read.
            continue

        print(f"{i}. Subject: {message.Subject}, Received: {message.ReceivedTime}")

        # 2. Download attachments.
        if message.Attachments.Count == 0:
            # We expect *ALL* emails to have an attachment!
            msg = f"Email {message.Subject} ({i:,}) has no attachments!"
            raise ValueError(msg)

        for attachment in message.Attachments:
            file_name = Path(attachment.FileName).resolve()
            attachment.SaveAsFile(file_name)
            print(f'Saved attachment {attachment.FileName} to "{file_name}"')

        # 3. Mark the email as read.
        message.UnRead = False
        message.Save()

        # 4. Find emails older than $n$ days. Ensure their attachments have been
        #    downloaded, then delete the email.


if __name__ == "__main__":
    parse_env_vars()
    load_emails()
