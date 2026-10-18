import base64
import pickle
import shutil
import webbrowser
from datetime import datetime
from pathlib import Path
from io import BytesIO

from dataclasses import dataclass

import click
from git_root import git_root
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.http import MediaFileUpload
from pypdf import PasswordType, PdfReader, PdfWriter

# If modifying these SCOPES, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.file",
]

# Secrets live next to this script, independent of the working directory.
_SECRETS_DIR = Path(__file__).resolve().parent
_TOKEN_PATH = _SECRETS_DIR / "token.json"
_CLIENT_SECRETS_PATH = _SECRETS_DIR / "credentials.json"
_CLIENT_SECRETS_CONSOLE_URL = "https://console.cloud.google.com/apis/credentials"


def _ensure_client_secrets() -> Path:
    """Make sure the OAuth2 client secrets are present locally.

    If they aren't, open the Google Cloud console in the browser so that a
    'Desktop app' OAuth client can be created and downloaded, then copy the
    downloaded file into place.
    """
    if _CLIENT_SECRETS_PATH.exists():
        return _CLIENT_SECRETS_PATH

    print(
        f"No OAuth2 client secrets found at {_CLIENT_SECRETS_PATH}.\n"
        "Opening the Google Cloud console. Create an OAuth client ID of type "
        "'Desktop app' (with the Gmail and Google Drive APIs enabled) and "
        "download its JSON file."
    )
    webbrowser.open(_CLIENT_SECRETS_CONSOLE_URL)

    downloads = sorted(
        (Path.home() / "Downloads").glob("client_secret_*.json"),
        key=lambda path: path.stat().st_mtime,
    )
    downloaded = click.prompt(
        "Path to the downloaded client secrets file",
        type=click.Path(exists=True, dir_okay=False, path_type=Path),
        default=downloads[-1] if downloads else None,
    )
    shutil.copyfile(downloaded, _CLIENT_SECRETS_PATH)
    return _CLIENT_SECRETS_PATH


def _login() -> Credentials:
    """Run the OAuth2 flow in the browser."""
    flow = InstalledAppFlow.from_client_secrets_file(_ensure_client_secrets(), SCOPES)
    return flow.run_local_server(port=0)


def _credentials() -> Credentials:
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if _TOKEN_PATH.exists():
        try:
            creds = Credentials.from_authorized_user_file(_TOKEN_PATH, SCOPES)
        except ValueError:
            # Malformed token file, e.g. missing refresh token.
            creds = None

    if creds and creds.valid and creds.has_scopes(SCOPES):
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            # Refresh token was revoked or has expired.
            creds = None
    else:
        creds = None

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        creds = _login()

    # Save the credentials for the next run
    _TOKEN_PATH.write_text(creds.to_json())
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

    relevant_attachments = []
    relevant_filenames = []
    parts = message_content["payload"].get("parts", [])

    # TODO: Annotate.
    def process_parts(parts, prefix: str = "") -> None:
        nonlocal relevant_attachments
        nonlocal relevant_filenames
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

                    relevant_attachments.append(file_data)
                    relevant_filenames.append(
                        part["filename"].replace(" ", "_").replace("/", "-")
                    )

            # Recursively process nested parts
            if "parts" in part:
                process_parts(part["parts"], prefix + "--")

    process_parts(parts)
    return relevant_attachments, relevant_filenames


def export_pdf_wo_password(
    raw_pdf_data: bytes, output_pdf: Path, password: str
) -> bool:
    """Write the decrypted PDF and return whether decryption succeeded."""
    reader = PdfReader(BytesIO(raw_pdf_data))

    # decrypt doesn't raise on a wrong password, it returns NOT_DECRYPTED.
    if reader.is_encrypted and reader.decrypt(password) == PasswordType.NOT_DECRYPTED:
        print(f"Could not decrypt {output_pdf.name}: wrong password.")
        return False

    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

    print(f"Wrote unencrypted PDF to: {output_pdf}")
    return True


def list_uploaded_files(service: Resource) -> list[dict]:
    """List non-trashed files on Google Drive created by this app."""
    files: list[dict] = []
    page_token = None
    while True:
        result = (
            service.files()
            .list(
                q="trashed=false",
                fields="nextPageToken, files(id, name, createdTime)",
                pageSize=1000,
                pageToken=page_token,
            )
            .execute()
        )
        files.extend(result.get("files", []))
        page_token = result.get("nextPageToken")
        if not page_token:
            return files


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

    failed_filenames = []
    for message in messages:
        attachments, filenames = get_attachments(gmail_service, message.id)
        if len(attachments) > 1:
            raise ValueError("Some emails contained more than one attachment.")
        if not attachments:
            print(f"Skipping email {message.id}: no attachment.")
            continue

        if not export_pdf_wo_password(
            attachments[0], export_path / filenames[0], password
        ):
            failed_filenames.append(filenames[0])

    if failed_filenames:
        print(
            f"Could not decrypt {len(failed_filenames)} of {len(messages)} PDFs: "
            + ", ".join(failed_filenames)
        )

    google_drive_service = _google_drive_service(credentials)

    uploaded_names = {
        file["name"] for file in list_uploaded_files(google_drive_service)
    }

    for file_path in sorted(export_path.iterdir()):
        if file_path.name in uploaded_names:
            print(f"Already on Google Drive, skipping: {file_path.name}")
            continue
        upload_file(file_path, google_drive_service)
        print(f"Uploaded to Google Drive: {file_path.name}")


if __name__ == "__main__":
    main()
