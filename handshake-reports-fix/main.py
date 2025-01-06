"""Read emails from Outlook."""

import win32com.client


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
    load_emails()
