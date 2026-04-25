import base64
import json
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.settings.basic",
]


def get_credentials_path() -> Path:
    return Path(os.getenv("GOOGLE_CLIENT_SECRETS_FILE", "credentials.json"))


def get_token_path() -> Path:
    return Path(os.getenv("GOOGLE_TOKEN_FILE", "token.json"))


def get_client_config(
    provided_config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if provided_config:
        return provided_config

    raw_json = os.getenv("GOOGLE_CLIENT_CONFIG_JSON")
    if raw_json:
        return json.loads(raw_json)

    credentials_path = get_credentials_path()
    if not credentials_path.exists():
        raise FileNotFoundError(
            f"Missing OAuth client secrets file: {credentials_path}. "
            "Provide a local credentials.json file or configure OAuth client settings in Streamlit secrets."
        )

    return json.loads(credentials_path.read_text(encoding="utf-8"))


def get_client_config_type(client_config: dict[str, Any]) -> str:
    if "installed" in client_config:
        return "installed"
    if "web" in client_config:
        return "web"
    return "unknown"


def inspect_local_oauth_setup() -> tuple[bool, str]:
    credentials_path = get_credentials_path()
    if not credentials_path.exists():
        return (
            False,
            f"Missing `{credentials_path}`. Create a Google OAuth client of type "
            "`Desktop app`, download the JSON file, and save it with that name.",
        )

    try:
        client_config = get_client_config()
    except Exception as exc:
        return False, f"Could not read `{credentials_path}`: {exc}"

    config_type = get_client_config_type(client_config)
    if config_type != "installed":
        return (
            False,
            "Your `credentials.json` is not a Desktop app OAuth client. "
            "For `localhost`, use a Google OAuth client of type `Desktop app`. "
            "Use a `Web application` client only for the hosted Streamlit Cloud app.",
        )

    return True, f"Found a valid Desktop app OAuth client at `{credentials_path}`."


def load_saved_credentials() -> Credentials | None:
    token_path = get_token_path()
    if not token_path.exists():
        return None

    try:
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            save_credentials(creds)
        return creds if creds and creds.valid else None
    except Exception:
        return None


def save_credentials(creds: Credentials) -> None:
    get_token_path().write_text(creds.to_json(), encoding="utf-8")


def authenticate_interactive() -> Credentials:
    client_config = get_client_config()
    config_type = get_client_config_type(client_config)
    if config_type != "installed":
        raise ValueError(
            "Local Gmail sign-in requires a Google OAuth client of type `Desktop app`. "
            "Replace `credentials.json` with a Desktop app client file and try again."
        )

    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
    creds = flow.run_local_server(port=0, open_browser=True)
    save_credentials(creds)
    return creds


def create_web_flow(
    redirect_uri: str,
    client_config: dict[str, Any] | None = None,
    state: str | None = None,
) -> Flow:
    flow = Flow.from_client_config(get_client_config(client_config), scopes=SCOPES, state=state)
    flow.redirect_uri = redirect_uri
    return flow


def exchange_code_for_credentials(
    code: str,
    redirect_uri: str,
    client_config: dict[str, Any] | None = None,
    state: str | None = None,
) -> Credentials:
    flow = create_web_flow(
        redirect_uri=redirect_uri,
        client_config=client_config,
        state=state,
    )
    flow.fetch_token(code=code)
    return flow.credentials


def get_gmail_service(creds: Credentials):
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def get_profile(service) -> dict[str, Any]:
    return service.users().getProfile(userId="me").execute()


def get_gmail_signature(service) -> str:
    try:
        response = service.users().settings().sendAs().list(userId="me").execute()
    except HttpError:
        return ""

    send_as_entries = response.get("sendAs", [])
    if not send_as_entries:
        return ""

    default_entry = next(
        (entry for entry in send_as_entries if entry.get("isDefault")),
        send_as_entries[0],
    )
    return default_entry.get("signature", "") or ""


def build_message(
    to_email: str,
    subject: str,
    html_body: str,
    from_email: str | None = None,
) -> dict[str, str]:
    message = MIMEMultipart("alternative")
    message["To"] = to_email
    message["Subject"] = subject
    if from_email:
        message["From"] = from_email

    message.attach(MIMEText(html_body, "html", "utf-8"))
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    return {"raw": raw_message}


def send_email(
    service,
    to_email: str,
    subject: str,
    html_body: str,
    from_email: str | None = None,
) -> dict[str, Any]:
    body = build_message(
        to_email=to_email,
        subject=subject,
        html_body=html_body,
        from_email=from_email,
    )
    return service.users().messages().send(userId="me", body=body).execute()
