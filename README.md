# Gmail Mail Merge Sender

A Streamlit app for safely sending personalized HTML email batches through the Gmail API.

## Features

- OAuth 2.0 authentication with Gmail
- CSV and XLSX recipient uploads
- Required `Email` column validation
- Optional per-row `Subject` column override
- HTML templates with variables like `{{Name}}` and `{{Joining Date}}`
- Variable detection and column mapping
- Personalized preview before sending
- Gmail signature fetch from Gmail send-as settings
- Test email sending
- Batch sending through the Gmail API
- Duplicate detection with optional override
- Daily send safety limit
- Delay between sends
- Send logging to `send_log.csv`

## Project Files

- `app.py`: Streamlit user interface and workflow
- `gmail_service.py`: Gmail OAuth, profile, signature, and send helpers
- `template_utils.py`: Variable detection and Jinja2 rendering helpers
- `validation.py`: File validation, email validation, and duplicate checks

## Setup

1. Create and activate a Python virtual environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Copy the environment example if you want custom paths:

```bash
cp .env.example .env
```

4. Create a Google Cloud project and enable the Gmail API.
5. Configure the OAuth consent screen.
6. For local use, create an OAuth 2.0 client ID for a Desktop app.
7. Download the OAuth client JSON file and save it as `credentials.json` in the project root, or point `GOOGLE_CLIENT_SECRETS_FILE` to it.

Do not commit `credentials.json`, `token.json`, `.env`, or `send_log.csv`. The included `.gitignore` excludes them by default.

## OAuth scopes used

- `https://www.googleapis.com/auth/gmail.send`
- `https://www.googleapis.com/auth/gmail.readonly`
- `https://www.googleapis.com/auth/gmail.settings.basic`

These scopes allow the app to send messages, read the authenticated Gmail profile, and fetch the default Gmail signature from send-as settings.

## Run the App

```bash
streamlit run app.py
```

When you click **Connect Gmail**, the app opens Google's OAuth flow in your browser. After approval, the token is saved locally in `token.json` by default.

## Deploy To Streamlit Community Cloud

This app is prepared for Streamlit Community Cloud, but the hosted deployment uses a web OAuth flow instead of the local desktop OAuth flow.

### 1. Push the repo to GitHub

Make sure these files are in your repository:

- `app.py`
- `gmail_service.py`
- `template_utils.py`
- `validation.py`
- `requirements.txt`
- `.streamlit/secrets.toml.example`

### 2. Create a Google OAuth client for the deployed app

In Google Cloud:

1. Open your existing project.
2. Go to **APIs & Services** > **Credentials**.
3. Create an OAuth client ID of type **Web application**.
4. Add your Streamlit app URL as an authorized redirect URI.

Example:

```text
https://your-app-name.streamlit.app
```

Use the exact deployed URL with no extra path unless you intentionally deploy behind a custom path.

### 3. Add Streamlit secrets

Copy `.streamlit/secrets.toml.example` into the Streamlit Cloud app's **Advanced settings** > **Secrets** field and replace the placeholder values:

```toml
[google_oauth]
client_id = "your-google-web-client-id.apps.googleusercontent.com"
client_secret = "your-google-client-secret"
project_id = "your-google-cloud-project-id"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
redirect_uri = "https://your-app-name.streamlit.app"
```

### 4. Deploy on Streamlit Community Cloud

From Streamlit Community Cloud:

1. Click **Create app**.
2. Choose your GitHub repository.
3. Set the branch to `main`.
4. Set the entrypoint to `app.py`.
5. Paste the secrets above into **Advanced settings**.
6. Deploy the app.

### 5. Test sign-in

After deployment:

1. Open the app URL.
2. Click **Connect Gmail**.
3. Complete the Google consent flow.
4. Confirm the app returns to Streamlit and shows your Gmail address.

If you get a `redirect_uri_mismatch` error, your Google OAuth web client redirect URI does not exactly match the deployed Streamlit app URL.

## Expected Recipient File Format

The uploaded CSV or XLSX file must include:

- `Email`

It may also include:

- `Subject`
- `Name`
- `Joining Date`
- `Department`
- Any other fields you want to reference in the template

Example:

| Email | Subject | Name | Joining Date | Department |
| --- | --- | --- | --- | --- |
| alex@example.com | Welcome aboard | Alex | 2026-05-01 | Product |

## Workflow

1. Connect Gmail
2. Upload recipients
3. Write a default subject and HTML template
4. Review detected variables and map them to uploaded columns
5. Preview a personalized email
6. Send a test email
7. Send the batch with a daily safety limit and delay between sends

## Notes

- The Gmail signature is fetched from the authenticated account's send-as settings and appended manually when enabled.
- Gmail does not automatically append the signature for API-sent messages in this app.
- The app validates email format before sending.
- Duplicate email addresses are skipped by default after the first occurrence.
- The app logs every send attempt with recipient, subject, status, timestamp, and error details.
- This app does not attempt to bypass Gmail limits or anti-spam protections.

## Troubleshooting

- If authentication fails, verify that `credentials.json` is a Desktop OAuth client file.
- If hosted sign-in fails on Streamlit Cloud, verify that your Google OAuth client type is **Web application** and that the authorized redirect URI exactly matches the `redirect_uri` value in Streamlit secrets.
- If you change OAuth scopes, delete `token.json` and authenticate again.
- If signature retrieval returns blank, confirm the Gmail account has a default signature configured.
- If XLSX upload fails, make sure `openpyxl` is installed.
