"""Read emails from Outlook."""

import datetime
import os
from collections.abc import Generator
from pathlib import Path
from typing import TypeAlias

import win32com.client
from dotenv import load_dotenv
from tqdm.auto import tqdm

N_DAYS_KEPT = 100  # Number of days to keep emails before deleting them.

HS_EMAIL = "handshake@notifications.joinhandshake.com"
included_folders = ("Inbox", "Handshake-Reports")


# Custom types for clearer type hints.
OutlookNamespace: TypeAlias = win32com.client.CDispatch
OutlookFolder: TypeAlias = win32com.client.CDispatch
OutlookMailItem: TypeAlias = win32com.client.CDispatch
OutlookItems: TypeAlias = win32com.client.CDispatch


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


def get_messages(
    root_folder: OutlookFolder,
    sort: bool = True,
    include_dirs: tuple[str, ...] = included_folders,
) -> Generator[OutlookMailItem, None, None]:
    """Get messages from the specified folders in Outlook."""
    for folder in root_folder.Folders:
        if folder.Name in include_dirs:
            messages: OutlookItems = folder.Items
            if sort:
                messages.Sort("[ReceivedTime]", True)

            yield from tqdm(messages, desc=f'Scanning "{folder.Name}"')


def load_emails() -> None:
    """Read emails from the "handshake-reports" folder in Outlook."""
    # Connect to Outlook
    outlook: OutlookNamespace
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

    root_folder = outlook.Folders[users_email]

    # Process the emails
    for i, message in enumerate(get_messages(root_folder)):
        # 1. Skip messages that aren't relevant.
        if (
            "BILAKE" not in message.Subject  # If the subject doesn't contain "BILAKE"
            or "error" in message.Subject.casefold()  # or if it contains "error"
            or not message.UnRead  # or if it's already read
            or message.SenderEmailAddress != HS_EMAIL  # or it's not from Handshake
        ):
            # Skip this email! It's got the wrong subject, wrong sender, or is
            # already read.
            continue

        tqdm.write(f"{i}. Subject: {message.Subject}, Received: {message.ReceivedTime}")

        # 2. Download attachments.
        if message.Attachments.Count == 0:
            # We expect *ALL* emails to have an attachment!
            msg = f"Email {message.Subject} ({i:,}) has no attachments!"
            raise ValueError(msg)

        for attachment in message.Attachments:
            file_name = Path(attachment.FileName).resolve()
            attachment.SaveAsFile(file_name)
            tqdm.write(f'Saved attachment {attachment.FileName} to "{file_name}"')

        # 3. Mark the email as read.
        message.UnRead = False
        message.Save()

        # 4. Find emails older than $n$ days. Ensure their attachments have been
        #    downloaded, then delete the email.
        today = datetime.datetime.now(datetime.UTC)
        if today - message.ReceivedTime > datetime.timedelta(days=N_DAYS_KEPT):
            message.Delete()
            tqdm.write(f"Deleted email {message.Subject} ({i:,})")


if __name__ == "__main__":
    parse_env_vars()
    load_emails()
