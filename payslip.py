import base64
import os
import pickle
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryFile

from dataclasses import dataclass

import click
from git_root import git_root
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.http import MediaFileUpload
from pypdf import PdfReader, PdfWriter

# If modifying these SCOPES, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.file",
]


def _credentials() -> Credentials:
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


def _gmail_service(credentials: Credentials, version: str = "v1") -> Resource:
    return build("gmail", version, credentials=credentials)


def _google_drive_service(credentials: Credentials, version: str = "v3") -> Resource:
    return build("drive", version, credentials=credentials)


@dataclass
class Message:
    id: str
    threadId: str


def search_messages(service: Resource, query: str) -> list[Message]:
    """Search for messages that match the query."""
    result = service.users().messages().list(userId="me", q=query).execute()
    messages = []

    if "messages" in result:
        messages.extend([Message(**message) for message in result["messages"]])

    while "nextPageToken" in result:
        page_token = result["nextPageToken"]
        result = (
            service.users()
            .messages()
            .list(userId="me", q=query, pageToken=page_token)
            .execute()
        )
        if "messages" in result:
            messages.extend([Message(**message) for message in result["messages"]])

    return messages


def get_attachments(
    service: Resource, message_id: str
) -> tuple[list[bytes], list[str]]:
    """Get and save attachments from a message."""
    message_content = (
        service.users()
        .messages()
        .get(userId="me", id=message_id, format="full")
        .execute()
    )

    subject = ""
    for header in message_content["payload"]["headers"]:
        if header["name"] == "Subject":
            subject = header["value"]

    if "parts" not in message_content["payload"]:
        return [], []

    saved_attachments = []
    saved_filenames = []
    parts = message_content["payload"].get("parts", [])

    # TODO: Annotate.
    def process_parts(parts, prefix: str = "") -> None:
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

                    # TODO: First validate if attachment is a pdf file.
                    file_data = base64.urlsafe_b64decode(attachment["data"])

                    saved_attachments.append(file_data)
                    saved_filenames.append(
                        part["filename"].replace(" ", "_").replace("/", "-")
                    )

            # Recursively process nested parts
            if "parts" in part:
                process_parts(part["parts"], prefix + "--")

    process_parts(parts)
    breakpoint()
    return saved_attachments, saved_filenames


def export_pdf_wo_password(
    raw_pdf_data: bytes, output_pdf: Path, password: str
) -> None:
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

    print(f"Wrote unencrypted PDF to: {output_pdf}")


def upload_file(path: Path, service: Resource):
    media = MediaFileUpload(path)

    google_drive_file = (
        service.files()
        .create(body={"name": path.name}, media_body=media, fields="id")
        .execute()
    )

    return google_drive_file.get("id")


@click.command()
@click.password_option("--password")
@click.option("--subject", default="Lohnabrechnung", type=str)
def main(password: str, subject: str) -> None:
    export_path = Path(git_root(subject))
    export_path.mkdir(exist_ok=True)

    credentials = _credentials()

    gmail_service = _gmail_service(credentials)

    query = "subject:" + subject
    messages = search_messages(gmail_service, query)

    print(f"Found {len(messages)} emails matching the query.")

    for message in messages:
        attachments, filenames = get_attachments(gmail_service, message.id)
        if len(attachments) > 1:
            raise ValueError("Some emails contained more than one attachment.")

        export_pdf_wo_password(attachments[0], export_path / filenames[0], password)

    google_drive_service = _google_drive_service(credentials)

    # TODO: Check if file already present on Google Drive.

    for file_path in export_path.iterdir():
        upload_file(file_path, google_drive_service)


if __name__ == "__main__":
    main()
