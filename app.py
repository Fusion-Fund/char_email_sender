"""
Streamlit Bulk Gmail Sender
Reads a CSV (name, email), personalizes each email, and sends via Gmail SMTP.
"""

import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
import os

load_dotenv()

st.set_page_config(page_title="Bulk Gmail Sender", page_icon="✉️", layout="wide")
st.title("Bulk Gmail Sender")
st.markdown("Send personalized emails to a list of recipients from a CSV file.")

# Sidebar: credentials and sender name
with st.sidebar:
    st.header("Sender Details")
    gmail_address = st.text_input(
        "Gmail address",
        value=os.getenv("GMAIL_ADDRESS", ""),
        placeholder="you@gmail.com",
        help="Your Gmail account used to send emails.",
    )
    gmail_app_password = st.text_input(
        "App password",
        value=os.getenv("GMAIL_APP_PASSWORD", ""),
        type="password",
        placeholder="16-character app password",
        help="Generate at: Google Account → Security → App Passwords",
    )
    sender_name = st.text_input(
        "Sender name",
        value="",
        placeholder="John Doe",
        help="Your name as it appears at the end of each email.",
    )


def validate_csv(df: pd.DataFrame) -> tuple[bool, list[str]]:
    """Validate CSV has name and email columns. Returns (is_valid, messages)."""
    messages = []
    cols_lower = {c.strip().lower(): c for c in df.columns}
    if "name" not in cols_lower:
        messages.append("CSV must have a 'name' column.")
    if "email" not in cols_lower:
        messages.append("CSV must have an 'email' column.")
    if messages:
        return False, messages
    return True, messages


def build_email_body(name: str, body: str, sender: str) -> str:
    """Build personalized email body."""
    return f"""Dear {name},

{body}

Best regards,
{sender}"""


def send_email(
    gmail_address: str,
    app_password: str,
    to_email: str,
    subject: str,
    body: str,
) -> str | None:
    """Send a single email. Returns None on success, error message on failure."""
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = gmail_address
        msg["To"] = to_email
        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(gmail_address, app_password.replace(" ", ""))
            server.sendmail(gmail_address, to_email, msg.as_string())
        return None
    except Exception as e:
        return str(e)


# Main area
st.header("1. Upload recipients (CSV)")
uploaded_file = st.file_uploader(
    "Upload CSV file",
    type=["csv"],
    help="CSV must have 'name' and 'email' columns.",
)

df = None
if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file)
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        is_valid, msgs = validate_csv(df)
        if not is_valid:
            for m in msgs:
                st.error(m)
        else:
            st.success(f"Loaded {len(df)} recipients.")
            with st.expander("Preview data"):
                st.dataframe(df, use_container_width=True)
            email_col = [c for c in df.columns if c.strip().lower() == "email"][0]
            invalid = df[df[email_col].isna() | (df[email_col].astype(str).str.strip() == "")]
            if len(invalid) > 0:
                st.warning(f"{len(invalid)} row(s) have empty email addresses.")
    except Exception as e:
        st.error(f"Could not parse CSV: {e}")

st.header("2. Compose email")
subject = st.text_input("Subject", placeholder="Your email subject")
body = st.text_area("Email body", placeholder="The main content of your email.", height=150)

if df is not None and subject and body and sender_name:
    name_col = [c for c in df.columns if c.strip().lower() == "name"][0]
    email_col = [c for c in df.columns if c.strip().lower() == "email"][0]
    first_row = df.iloc[0]
    sample_body = build_email_body(
        str(first_row[name_col]),
        body,
        sender_name,
    )
    with st.expander("Preview sample email (first recipient)"):
        st.text(f"To: {first_row[email_col]}\nSubject: {subject}\n\n{sample_body}")

st.header("3. Send emails")
if st.button("Send All Emails", type="primary"):
    if not gmail_address or not gmail_app_password:
        st.error("Please enter Gmail address and app password in the sidebar.")
    elif not sender_name:
        st.error("Please enter your sender name in the sidebar.")
    elif df is None:
        st.error("Please upload a valid CSV file first.")
    elif not subject:
        st.error("Please enter a subject.")
    elif not body:
        st.error("Please enter an email body.")
    else:
        name_col = [c for c in df.columns if c.strip().lower() == "name"][0]
        email_col = [c for c in df.columns if c.strip().lower() == "email"][0]
        rows = df[df[email_col].notna() & (df[email_col].astype(str).str.strip() != "")].to_dict("records")
        total = len(rows)

        if total == 0:
            st.error("No valid recipients with email addresses.")
        else:
            progress_bar = st.progress(0)
            status_container = st.status("Sending emails...", expanded=True)
            results = []

            with status_container:
                for i, row in enumerate(rows):
                    to_email = str(row[email_col]).strip()
                    name = str(row.get(name_col, "")).strip() or "there"
                    body_text = build_email_body(name, body, sender_name)

                    err = send_email(
                        gmail_address,
                        gmail_app_password,
                        to_email,
                        subject,
                        body_text,
                    )
                    if err:
                        results.append({"email": to_email, "status": "failed", "error": err})
                        st.write(f"❌ {to_email}: {err}")
                    else:
                        results.append({"email": to_email, "status": "success"})
                        st.write(f"✓ {to_email}")

                    progress_bar.progress((i + 1) / total)
                    time.sleep(0.5)

            success_count = sum(1 for r in results if r["status"] == "success")
            fail_count = len(results) - success_count
            st.success(f"Done: {success_count} sent, {fail_count} failed.")

            with st.expander("Detailed results"):
                for r in results:
                    if r["status"] == "success":
                        st.write(f"✓ {r['email']}")
                    else:
                        st.write(f"❌ {r['email']}: {r.get('error', 'Unknown error')}")
