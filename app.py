import os
import time
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from gmail_service import (
    authenticate_interactive,
    create_web_flow,
    exchange_code_for_credentials,
    get_gmail_service,
    get_gmail_signature,
    get_profile,
    load_saved_credentials,
    send_email,
)
from template_utils import (
    append_signature,
    auto_map_variables,
    detect_template_variables,
    render_template,
)
from validation import annotate_recipients, get_duplicate_emails, validate_required_columns


APP_TITLE = "Gmail Mail Merge Sender"
LOG_FILE = Path("send_log.csv")
DEFAULT_TEMPLATE = """<p>Hi {{Name}},</p>
<p>Welcome to the {{Department}} team. Your joining date is <strong>{{Joining Date}}</strong>.</p>
<p>Best regards,<br>HR Team</p>"""


load_dotenv()
st.set_page_config(page_title=APP_TITLE, page_icon="📧", layout="wide")


def load_send_log() -> pd.DataFrame:
    if LOG_FILE.exists():
        return pd.read_csv(LOG_FILE)

    return pd.DataFrame(
        columns=["recipient_email", "subject", "status", "timestamp", "error_message"]
    )


def append_log_entry(recipient_email: str, subject: str, status: str, error_message: str = "") -> None:
    entry = pd.DataFrame(
        [
            {
                "recipient_email": recipient_email,
                "subject": subject,
                "status": status,
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "error_message": error_message,
            }
        ]
    )
    header = not LOG_FILE.exists()
    entry.to_csv(LOG_FILE, mode="a", header=header, index=False)


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(uploaded_file)
    if suffix == ".xlsx":
        return pd.read_excel(uploaded_file)
    raise ValueError("Unsupported file type. Please upload a CSV or XLSX file.")


def standardize_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    renamed = dataframe.copy()
    rename_map: dict[str, str] = {}
    for column in renamed.columns:
        column_name = str(column).strip()
        normalized = column_name.casefold()
        if normalized == "email":
            rename_map[column] = "Email"
        elif normalized == "subject":
            rename_map[column] = "Subject"
        else:
            rename_map[column] = column_name
    return renamed.rename(columns=rename_map)


def get_today_sent_count(log_df: pd.DataFrame) -> int:
    if log_df.empty:
        return 0
    timestamps = pd.to_datetime(log_df["timestamp"], errors="coerce")
    today = datetime.now().date()
    return int(((timestamps.dt.date == today) & (log_df["status"] == "sent")).sum())


def get_subject_for_row(row: dict, subject_template: str) -> str:
    row_subject = row.get("Subject")
    if row_subject is None or pd.isna(row_subject) or str(row_subject).strip() == "":
        return subject_template
    return str(row_subject)


def render_email_content(
    row: dict,
    subject_template: str,
    body_template: str,
    variable_mapping: dict[str, str | None],
    include_signature: bool,
    signature_html: str,
) -> tuple[str, str]:
    subject_source = get_subject_for_row(row, subject_template)
    rendered_subject = render_template(subject_source or "", row, variable_mapping)
    rendered_body = render_template(body_template, row, variable_mapping)
    final_body = append_signature(rendered_body, signature_html, include_signature)
    return rendered_subject, final_body


def build_sample_recipients_csv() -> bytes:
    sample_df = pd.DataFrame(
        [
            {
                "Email": "alex@example.com",
                "Subject": "Welcome to the team",
                "Name": "Alex",
                "Joining Date": "2026-05-01",
                "Department": "Product",
            },
            {
                "Email": "jamie@example.com",
                "Subject": "Your onboarding details",
                "Name": "Jamie",
                "Joining Date": "2026-05-05",
                "Department": "Engineering",
            },
        ]
    )
    return sample_df.to_csv(index=False).encode("utf-8")


def initialize_session_state() -> None:
    defaults = {
        "creds": None,
        "gmail_profile": None,
        "gmail_signature": "",
        "oauth_state": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def get_google_client_config_from_secrets() -> dict | None:
    if "google_oauth" not in st.secrets:
        return None

    oauth_section = dict(st.secrets["google_oauth"])
    return {
        "web": {
            "client_id": oauth_section["client_id"],
            "project_id": oauth_section.get("project_id", ""),
            "auth_uri": oauth_section.get("auth_uri", "https://accounts.google.com/o/oauth2/auth"),
            "token_uri": oauth_section.get("token_uri", "https://oauth2.googleapis.com/token"),
            "auth_provider_x509_cert_url": oauth_section.get(
                "auth_provider_x509_cert_url",
                "https://www.googleapis.com/oauth2/v1/certs",
            ),
            "client_secret": oauth_section["client_secret"],
            "redirect_uris": [oauth_section["redirect_uri"]],
        }
    }


def get_streamlit_redirect_uri() -> str | None:
    if "google_oauth" not in st.secrets:
        return None
    return str(st.secrets["google_oauth"]["redirect_uri"]).rstrip("/")


def get_allowed_sender_emails() -> set[str]:
    if "app_security" not in st.secrets:
        return set()

    raw_value = st.secrets["app_security"].get("allowed_emails", [])
    if isinstance(raw_value, str):
        values = [raw_value]
    else:
        values = list(raw_value)

    return {str(value).strip().casefold() for value in values if str(value).strip()}


def is_allowed_sender(email_address: str | None) -> bool:
    allowed_emails = get_allowed_sender_emails()
    if not allowed_emails:
        return True
    if not email_address:
        return False
    return email_address.strip().casefold() in allowed_emails


def is_cloud_oauth_configured() -> bool:
    return get_google_client_config_from_secrets() is not None and get_streamlit_redirect_uri() is not None


def finish_cloud_oauth_if_needed() -> None:
    query_params = st.query_params
    auth_code = query_params.get("code")
    returned_state = query_params.get("state")
    oauth_error = query_params.get("error")

    if oauth_error:
        st.error(f"Google OAuth returned an error: {oauth_error}")
        for key in ["code", "state", "error", "scope", "authuser", "prompt"]:
            if key in query_params:
                del query_params[key]
        return

    if not auth_code:
        return

    if st.session_state["creds"] is not None:
        for key in ["code", "state", "scope", "authuser", "prompt"]:
            if key in query_params:
                del query_params[key]
        return

    expected_state = st.session_state.get("oauth_state")
    if expected_state and returned_state and returned_state != expected_state:
        st.error("OAuth state mismatch. Please try connecting Gmail again.")
        return

    try:
        st.session_state["creds"] = exchange_code_for_credentials(
            code=auth_code,
            redirect_uri=get_streamlit_redirect_uri() or "",
            client_config=get_google_client_config_from_secrets(),
            state=expected_state,
        )
        for key in ["code", "state", "scope", "authuser", "prompt"]:
            if key in query_params:
                del query_params[key]
        st.session_state["oauth_state"] = None
        st.success("Google authentication completed.")
    except Exception as exc:
        st.error(f"Cloud OAuth failed: {exc}")


def render_cloud_auth_button() -> None:
    redirect_uri = get_streamlit_redirect_uri()
    client_config = get_google_client_config_from_secrets()
    if not redirect_uri or not client_config:
        st.warning("Add Google OAuth settings to Streamlit secrets to enable hosted sign-in.")
        return

    flow = create_web_flow(redirect_uri=redirect_uri, client_config=deepcopy(client_config))
    authorization_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    st.session_state["oauth_state"] = state
    st.link_button("Connect Gmail", authorization_url, use_container_width=True)


initialize_session_state()
finish_cloud_oauth_if_needed()

st.title(APP_TITLE)
st.caption("Safely send personalized HTML emails through the Gmail API.")


with st.container(border=True):
    st.subheader("Step 1: Connect Gmail")
    credentials_path = Path.cwd() / "credentials.json"
    local_oauth_ready = False
    if is_cloud_oauth_configured():
        st.write(
            "This app is configured for hosted Google OAuth. Use the button below to sign in, "
            "and make sure your Google OAuth web client redirect URI matches this app URL."
        )
    else:
        st.write(
            "Place your Google OAuth desktop client file at "
            f"`{credentials_path}` before connecting."
        )
        local_oauth_ready = credentials_path.exists()
        if local_oauth_ready:
            local_oauth_message = (
                f"Found `{credentials_path}`. If local Gmail sign-in fails, make sure it is a "
                "`Desktop app` OAuth client from Google Cloud."
            )
        else:
            local_oauth_message = (
                f"Missing `{credentials_path}`. Create a Google OAuth client of type "
                "`Desktop app`, download the JSON file, and save it with that name."
            )
        if local_oauth_ready:
            st.success(local_oauth_message)
        else:
            st.warning(local_oauth_message)

    if st.session_state["creds"] is None:
        st.session_state["creds"] = load_saved_credentials() if not is_cloud_oauth_configured() else None

    col1, col2 = st.columns([1, 2])
    with col1:
        if is_cloud_oauth_configured():
            render_cloud_auth_button()
        elif st.button(
            "Connect Gmail",
            use_container_width=True,
            disabled=not local_oauth_ready,
        ):
            try:
                st.session_state["creds"] = authenticate_interactive()
                st.success("Google authentication completed.")
            except Exception as exc:
                st.error(str(exc))

    if st.session_state["creds"] is not None:
        try:
            service = get_gmail_service(st.session_state["creds"])
            st.session_state["gmail_profile"] = get_profile(service)
            st.session_state["gmail_signature"] = get_gmail_signature(service)
            connected_email = st.session_state["gmail_profile"].get("emailAddress")
            if not is_allowed_sender(connected_email):
                st.session_state["creds"] = None
                st.session_state["gmail_profile"] = None
                st.session_state["gmail_signature"] = ""
                service = None
                st.error(
                    "This hosted app is restricted. The signed-in Gmail address is not allowed to use it."
                )
        except Exception as exc:
            st.error(f"Failed to load Gmail account details: {exc}")
            service = None
    else:
        service = None

    with col2:
        profile = st.session_state.get("gmail_profile")
        if profile:
            st.success(f"Connected as `{profile.get('emailAddress', 'Unknown account')}`")
            if st.session_state["gmail_signature"]:
                st.info("A Gmail signature was found and can be appended to outgoing emails.")
            else:
                st.info("No Gmail signature was returned for the default send-as address.")
        else:
            st.warning("Connect an account to continue.")


dataframe = None
annotated_df = None

with st.container(border=True):
    st.subheader("Step 2: Upload recipients")
    st.write("Download a sample CSV if you want a ready-made upload format.")
    st.download_button(
        "Download sample CSV template",
        data=build_sample_recipients_csv(),
        file_name="gmail_mail_merge_template.csv",
        mime="text/csv",
    )
    st.caption("Required column: `Email`. Optional examples: `Subject`, `Name`, `Joining Date`, `Department`.")

    uploaded_file = st.file_uploader(
        "Upload a CSV or XLSX file",
        type=["csv", "xlsx"],
    )

    if uploaded_file is not None:
        try:
            dataframe = standardize_columns(read_uploaded_file(uploaded_file))
            errors = validate_required_columns(dataframe)
            if errors:
                for error in errors:
                    st.error(error)
            else:
                annotated_df = annotate_recipients(dataframe)
                duplicate_emails = get_duplicate_emails(annotated_df)

                metric_cols = st.columns(4)
                metric_cols[0].metric("Rows", len(annotated_df))
                metric_cols[1].metric(
                    "Valid Emails", int(annotated_df["_is_valid_email"].sum())
                )
                metric_cols[2].metric(
                    "Invalid Emails", int((~annotated_df["_is_valid_email"]).sum())
                )
                metric_cols[3].metric("Duplicate Emails", len(duplicate_emails))

                st.dataframe(annotated_df.drop(columns=["_normalized_email"]), use_container_width=True)
                if duplicate_emails:
                    st.warning(
                        "Duplicate recipient emails detected: "
                        + ", ".join(duplicate_emails[:10])
                        + (" ..." if len(duplicate_emails) > 10 else "")
                    )
        except Exception as exc:
            st.error(f"Could not read the file: {exc}")


subject_template = ""
body_template = ""
variable_mapping: dict[str, str | None] = {}
include_signature = False

with st.container(border=True):
    st.subheader("Step 3: Write subject and template")
    subject_template = st.text_input(
        "Default subject",
        value="Welcome to the team",
        help="Used when the uploaded file does not include a Subject column or when a row subject is blank.",
    )
    body_template = st.text_area(
        "HTML email template",
        value=DEFAULT_TEMPLATE,
        height=260,
        help="Use variables like {{Name}} or {{Joining Date}}.",
    )

    include_signature = st.checkbox(
        "Include Gmail signature",
        value=bool(st.session_state["gmail_signature"]),
        disabled=not bool(st.session_state["gmail_signature"]),
        help="The signature is fetched from your Gmail send-as settings and appended manually.",
    )

    if annotated_df is not None:
        variables = detect_template_variables(subject_template, body_template)
        if variables:
            st.write("Detected template variables")
            auto_mapping = auto_map_variables(variables, list(annotated_df.columns))
            mapping_columns = st.columns(2)
            for index, variable in enumerate(variables):
                options = [""] + [column for column in annotated_df.columns if not column.startswith("_")]
                default_value = auto_mapping.get(variable) or ""
                selected = mapping_columns[index % 2].selectbox(
                    f"Map `{variable}`",
                    options=options,
                    index=options.index(default_value) if default_value in options else 0,
                    key=f"map_{variable}",
                    help="Pick the uploaded column that should fill this template variable.",
                )
                variable_mapping[variable] = selected or None
            unmapped_variables = [name for name, column in variable_mapping.items() if column is None]
            if unmapped_variables:
                st.warning(
                    "These variables are currently unmapped and will render as empty text: "
                    + ", ".join(f"`{name}`" for name in unmapped_variables)
                )
        else:
            st.info("No template variables detected yet.")


preview_row = None

with st.container(border=True):
    st.subheader("Step 4: Preview")
    if annotated_df is None:
        st.info("Upload recipients to preview personalized emails.")
    else:
        valid_rows = annotated_df[annotated_df["_is_valid_email"]].reset_index(drop=True)
        if valid_rows.empty:
            st.warning("No valid recipient rows are available for preview.")
        else:
            preview_index = st.number_input(
                "Preview row number",
                min_value=1,
                max_value=len(valid_rows),
                value=1,
                step=1,
            )
            preview_row = valid_rows.iloc[int(preview_index) - 1].to_dict()
            try:
                rendered_subject, rendered_body = render_email_content(
                    row=preview_row,
                    subject_template=subject_template,
                    body_template=body_template,
                    variable_mapping=variable_mapping,
                    include_signature=include_signature,
                    signature_html=st.session_state["gmail_signature"],
                )
                st.text_input("Preview subject", value=rendered_subject, disabled=True)
                st.write("Preview body")
                st.markdown(rendered_body, unsafe_allow_html=True)
            except Exception as exc:
                st.error(f"Preview failed: {exc}")


with st.container(border=True):
    st.subheader("Step 5: Send test")
    test_email = st.text_input(
        "Test recipient email",
        help="Sends one personalized email using the preview row content.",
    )
    if st.button("Send test email", use_container_width=True, disabled=service is None or preview_row is None):
        try:
            test_df = annotate_recipients(pd.DataFrame([{"Email": test_email}]))
            if not bool(test_df.iloc[0]["_is_valid_email"]):
                raise ValueError("Enter a valid test email address.")

            rendered_subject, rendered_body = render_email_content(
                row=preview_row,
                subject_template=subject_template,
                body_template=body_template,
                variable_mapping=variable_mapping,
                include_signature=include_signature,
                signature_html=st.session_state["gmail_signature"],
            )
            send_email_service = get_gmail_service(st.session_state["creds"])
            from_email = st.session_state.get("gmail_profile", {}).get("emailAddress")

            send_email(
                service=send_email_service,
                to_email=test_email,
                subject=rendered_subject,
                html_body=rendered_body,
                from_email=from_email,
            )
            append_log_entry(test_email, rendered_subject, "sent")
            st.success("Test email sent successfully.")
        except Exception as exc:
            append_log_entry(test_email, "test-email", "failed", str(exc))
            st.error(f"Test email failed: {exc}")


with st.container(border=True):
    st.subheader("Step 6: Send batch")
    allow_duplicates = st.checkbox(
        "Allow duplicate email addresses",
        value=False,
        help="By default, only the first occurrence of each email address is sent.",
    )
    daily_send_limit = st.number_input(
        "Daily send safety limit",
        min_value=1,
        value=50,
        step=1,
        help="This app will not send more than this many emails in one day based on the local send log.",
    )
    delay_seconds = st.number_input(
        "Delay between sends (seconds)",
        min_value=0.0,
        value=2.0,
        step=0.5,
        help="Adds a pause between each send to reduce the chance of account issues.",
    )

    ready_to_send = service is not None and annotated_df is not None and body_template.strip()
    if st.button("Send batch", type="primary", use_container_width=True, disabled=not ready_to_send):
        log_df = load_send_log()
        already_sent_today = get_today_sent_count(log_df)
        remaining_today = max(int(daily_send_limit) - already_sent_today, 0)

        if remaining_today <= 0:
            st.error("Daily send safety limit reached based on today's local send log.")
        else:
            run_service = get_gmail_service(st.session_state["creds"])
            from_email = st.session_state.get("gmail_profile", {}).get("emailAddress")

            send_queue = annotated_df.copy()
            progress_bar = st.progress(0)
            status_placeholder = st.empty()

            attempted = 0
            processed_count = 0
            sent_count = 0

            for _, row in send_queue.iterrows():
                row_dict = row.to_dict()
                recipient_email = row_dict.get("Email", "")
                subject_for_log = row_dict.get("Subject") or subject_template

                if attempted >= remaining_today:
                    append_log_entry(
                        recipient_email,
                        subject_for_log,
                        "skipped",
                        "Skipped because the daily send safety limit was reached.",
                    )
                elif not row_dict.get("_is_valid_email", False):
                    append_log_entry(
                        recipient_email,
                        subject_for_log,
                        "skipped",
                        "Skipped because the email format is invalid.",
                    )
                elif row_dict.get("_is_duplicate", False) and not allow_duplicates:
                    append_log_entry(
                        recipient_email,
                        subject_for_log,
                        "skipped",
                        "Skipped because this recipient email is a duplicate.",
                    )
                else:
                    try:
                        rendered_subject, rendered_body = render_email_content(
                            row=row_dict,
                            subject_template=subject_template,
                            body_template=body_template,
                            variable_mapping=variable_mapping,
                            include_signature=include_signature,
                            signature_html=st.session_state["gmail_signature"],
                        )
                        send_email(
                            service=run_service,
                            to_email=recipient_email,
                            subject=rendered_subject,
                            html_body=rendered_body,
                            from_email=from_email,
                        )
                        append_log_entry(recipient_email, rendered_subject, "sent")
                        sent_count += 1
                    except Exception as exc:
                        append_log_entry(recipient_email, subject_for_log, "failed", str(exc))

                    attempted += 1

                processed_count += 1
                progress_bar.progress(min(processed_count / max(len(send_queue), 1), 1.0))
                status_placeholder.write(
                    f"Processed {processed_count} of {len(send_queue)} rows. Sent {sent_count} emails."
                )
                if delay_seconds > 0 and processed_count < len(send_queue):
                    time.sleep(float(delay_seconds))

            st.success(f"Batch complete. Sent {sent_count} emails.")


with st.container(border=True):
    st.subheader("Send log")
    log_df = load_send_log()
    if log_df.empty:
        st.info("No send attempts have been logged yet.")
    else:
        st.dataframe(log_df, use_container_width=True)
        st.download_button(
            "Download send log CSV",
            data=log_df.to_csv(index=False).encode("utf-8"),
            file_name="send_log.csv",
            mime="text/csv",
        )

st.caption(
    "Use this tool responsibly. Gmail enforces sending limits and anti-spam protections, "
    "and this app is designed to respect them rather than bypass them."
)
