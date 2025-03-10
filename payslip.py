import base64
import os
import pickle
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryFile

import click
from git_root import git_root
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from pypdf import PdfReader, PdfWriter

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.file",
]


def _creds():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds


def _gmail_service():
    return build("gmail", "v1", credentials=_creds())


def _google_drive_service():
    return build("drive", "v3", credentials=_creds())


def search_messages(service, query: str) -> list:
    """Search for messages that match the query."""
    result = service.users().messages().list(userId="me", q=query).execute()
    messages = []

    if "messages" in result:
        messages.extend(result["messages"])

    while "nextPageToken" in result:
        page_token = result["nextPageToken"]
        result = (
            service.users()
            .messages()
            .list(userId="me", q=query, pageToken=page_token)
            .execute()
        )
        if "messages" in result:
            messages.extend(result["messages"])

    return messages


def get_attachments(service, message_id: str) -> tuple[list, list]:
    """Get and save attachments from a message."""
    message = (
        service.users()
        .messages()
        .get(userId="me", id=message_id, format="full")
        .execute()
    )

    subject = ""
    for header in message["payload"]["headers"]:
        if header["name"] == "Subject":
            subject = header["value"]
            break

    if "parts" not in message["payload"]:
        return []

    saved_attachments = []
    saved_filenames = []
    parts = message["payload"].get("parts", [])

    def process_parts(parts, prefix=""):
        nonlocal saved_attachments
        nonlocal saved_filenames
        for part in parts:
            if "filename" in part and part["filename"]:
                if "body" in part and "attachmentId" in part["body"]:
                    attachment = (
                        service.users()
                        .messages()
                        .attachments()
                        .get(
                            userId="me",
                            messageId=message_id,
                            id=part["body"]["attachmentId"],
                        )
                        .execute()
                    )

                    file_data = base64.urlsafe_b64decode(attachment["data"])

                    saved_attachments.append(file_data)
                    saved_filenames.append(
                        part["filename"].replace(" ", "_").replace("/", "-")
                    )

            # Recursively process nested parts
            if "parts" in part:
                process_parts(part["parts"], prefix + "--")

    process_parts(parts)
    return saved_attachments, saved_filenames


def export_wo_password(raw_pdf_data, output_pdf: Path, password: str) -> None:
    with TemporaryFile() as file_handle:
        file_handle.write(raw_pdf_data)
        reader = PdfReader(file_handle)

        if reader.is_encrypted:
            reader.decrypt(password)

        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

    print(f"Created unprotected PDF: {output_pdf}")


def upload_file(path: Path, service):
    media = MediaFileUpload(path)

    google_drive_file = (
        service.files()
        .create(body={"name": path.name}, media_body=media, fields="id")
        .execute()
    )

    return google_drive_file.get("id")


@click.command()
@click.password_option("--password")
@click.option("--subject", default="Lohnabrechnung")
def main(password, subject):
    export_path = Path(git_root(subject))
    export_path.mkdir(exist_ok=True)

    gmail_service = _gmail_service()

    query = "subject:" + subject
    messages = search_messages(gmail_service, query)

    print(f"Found {len(messages)} emails matching the query.")

    for message in messages:
        attachments, filenames = get_attachments(gmail_service, message["id"])
        if len(attachments) > 1:
            raise ValueError("Some emails contained more than one attachment.")

        export_wo_password(attachments[0], export_path / filenames[0], password)

    google_drive_service = _google_drive_service()

    for file_path in export_path.iterdir():
        upload_file(file_path, google_drive_service)


if __name__ == "__main__":
    main()
