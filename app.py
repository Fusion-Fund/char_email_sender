"""
Streamlit Bulk Gmail Sender
Reads a CSV (name, email), personalizes each email, and sends via Gmail SMTP.
"""

import base64
import html
import re
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

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


def personalize_text(text: str, receiver: str, sender: str) -> str:
    """Replace [RECEIVER], [NAME], and [SENDER] placeholders."""
    return text.replace("[RECEIVER]", receiver).replace("[NAME]", receiver).replace("[SENDER]", sender)


def _text_to_html(text: str) -> str:
    """Convert body text with markdown-like bullets and links to HTML."""
    # Escape HTML first so user content is safe
    text = html.escape(text)

    # Convert [text](url) to links (url and text already escaped)
    text = re.sub(
        r"\[([^\]]+)\]\(([^)]+)\)",
        lambda m: f'<a href="{m.group(2)}" style="color:#0066cc;text-decoration:underline;">{m.group(1)}</a>',
        text,
    )

    # Convert bullet points: lines starting with - or * (must have space after)
    lines = text.split("\n")
    result = []
    in_list = False
    for line in lines:
        stripped = line.lstrip()
        if stripped.startswith("- ") or stripped.startswith("* "):
            bullet_content = stripped[2:]
            if not in_list:
                result.append('<ul style="list-style-type:disc;margin-left:1.5em;padding-left:1em;">')
                in_list = True
            result.append(f"<li>{bullet_content}</li>")
        else:
            if in_list:
                result.append("</ul>")
                in_list = False
            result.append(line + "<br>" if line else "<br>")
    if in_list:
        result.append("</ul>")

    return "".join(result)


def build_email_body(name: str, body: str, sender: str) -> str:
    """Build personalized email body."""
    personalized_body = personalize_text(body, name, sender)
    return f"""Dear {name},

{personalized_body}

Best regards,
{sender}"""


def _image_subtype(fname: str) -> str:
    """Return MIME subtype for image filename."""
    ext = fname.lower().split(".")[-1] if "." in fname else ""
    return "jpeg" if ext in ("jpg", "jpeg") else (ext if ext in ("png", "gif") else "png")


def _build_html_email(name: str, html_body: str, sender: str) -> str:
    """Wrap HTML body with greeting and sign-off."""
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head><body>
<p>Dear {html.escape(name)},</p>
<div>{html_body}</div>
<p>Best regards,<br>{html.escape(sender)}</p>
</body></html>"""


def build_email_body_html(name: str, body: str, sender: str, image_list: list[tuple[str, bytes]], body_is_html: bool = False) -> str:
    """Build HTML email body. If body_is_html, body is WYSIWYG HTML; else convert from markdown-like text."""
    personalized_body = personalize_text(body, name, sender)
    parts = personalized_body.split("[IMAGE]")
    if not image_list:
        html_body = parts[0] if parts else ""
        if not body_is_html:
            html_body = _text_to_html(html_body)
    else:
        html_parts = []
        for i, part in enumerate(parts):
            html_parts.append(part if body_is_html else _text_to_html(part))
            if i < len(parts) - 1 and i < len(image_list):
                html_parts.append(f'<img src="cid:img{i}" alt="image" style="max-width:100%;"/>')
        html_body = "".join(html_parts)
    return _build_html_email(name, html_body, sender)


def build_preview_html(name: str, body: str, sender: str, image_list: list[tuple[str, str]], body_is_html: bool = False) -> str:
    """Build HTML for preview with data URLs. image_list is list of (filename, base64_data_url)."""
    personalized_body = personalize_text(body, name, sender)
    parts = personalized_body.split("[IMAGE]")
    if not image_list:
        html_body = parts[0] if parts else ""
        if not body_is_html:
            html_body = _text_to_html(html_body)
    else:
        html_parts = []
        for i, part in enumerate(parts):
            html_parts.append(part if body_is_html else _text_to_html(part))
            if i < len(parts) - 1 and i < len(image_list):
                _, data_url = image_list[i]
                html_parts.append(f'<img src="{data_url}" alt="image" style="max-width:100%;"/>')
        html_body = "".join(html_parts)
    return f"""<div style="font-family:sans-serif;">
<p>Dear {html.escape(name)},</p>
<div>{html_body}</div>
<p>Best regards,<br>{html.escape(sender)}</p>
</div>"""


def send_email(
    gmail_address: str,
    app_password: str,
    to_email: str,
    subject: str,
    body: str,
    cc_email: str = "",
    attachments: list[tuple[str, bytes]] | None = None,
    inline_images: list[tuple[str, bytes]] | None = None,
) -> str | None:
    """Send a single email. Returns None on success, error message on failure."""
    try:
        if inline_images:
            msg = MIMEMultipart("related")
            msg["Subject"] = subject
            msg["From"] = gmail_address
            msg["To"] = to_email
            if cc_email:
                msg["Cc"] = cc_email
            msg.attach(MIMEText(body, "html", "utf-8"))

            for i, (fname, data) in enumerate(inline_images):
                subtype = _image_subtype(fname)
                img = MIMEImage(data, _subtype=subtype)
                img.add_header("Content-ID", f"<img{i}>")
                img.add_header("Content-Disposition", "inline", filename=fname)
                msg.attach(img)

            if attachments:
                for fname, data in attachments:
                    img = MIMEImage(data, _subtype=_image_subtype(fname))
                    img.add_header("Content-Disposition", "attachment", filename=fname)
                    msg.attach(img)
        else:
            msg = MIMEMultipart("mixed")
            msg["Subject"] = subject
            msg["From"] = gmail_address
            msg["To"] = to_email
            if cc_email:
                msg["Cc"] = cc_email
            msg.attach(MIMEText(body, "html", "utf-8"))

            if attachments:
                for fname, data in attachments:
                    img = MIMEImage(data, _subtype=_image_subtype(fname))
                    img.add_header("Content-Disposition", "attachment", filename=fname)
                    msg.attach(img)

        recipients = [to_email] + ([cc_email] if cc_email else [])

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(gmail_address, app_password.replace(" ", ""))
            server.sendmail(gmail_address, recipients, msg.as_string())
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
subject = st.text_input("Subject", placeholder="Use [RECEIVER] and [SENDER] for personalization.")
cc_email = st.text_input("CC", placeholder="cc@example.com", help="This address will be CC'd on every email sent.")
with st.expander("📝 Formatting (bullets, links)"):
    st.markdown("**Bullets:** Start a line with `- ` or `* ` (space required). **Links:** `[text](https://url.com)`")
col_edit, col_preview = st.columns(2)
with col_edit:
    body = st.text_area(
        "Email body",
        placeholder="Use [RECEIVER], [SENDER], [IMAGE]. Bullets: - item. Links: [text](url)",
        height=200,
        key="email_body",
    )
with col_preview:
    st.caption("Live preview")
    if body and sender_name:
        preview_name = "Recipient"
        if df is not None:
            name_col = [c for c in df.columns if c.strip().lower() == "name"][0]
            preview_name = str(df.iloc[0][name_col])
        preview_html = build_preview_html(preview_name, body, sender_name, [], body_is_html=False)
        st.components.v1.html(preview_html, height=200, scrolling=True)
    else:
        st.info("Enter body + sender name to see preview.")
image_files = st.file_uploader("Drag and drop images to embed in body", type=["png", "jpg", "jpeg", "gif"], accept_multiple_files=True, help="Use [IMAGE] in the body where you want each image. Images are placed in order.")

if df is not None and subject and body and sender_name:
    name_col = [c for c in df.columns if c.strip().lower() == "name"][0]
    email_col = [c for c in df.columns if c.strip().lower() == "email"][0]
    first_row = df.iloc[0]
    first_name = str(first_row[name_col])
    sample_subject = personalize_text(subject, first_name, sender_name)
    with st.expander("Preview sample email (first recipient)"):
        st.write(f"**To:** {first_row[email_col]}")
        st.write(f"**Subject:** {sample_subject}")
        if cc_email and cc_email.strip():
            st.write(f"**CC:** {cc_email.strip()}")
        st.write("---")
        if image_files and "[IMAGE]" in personalize_text(body, first_name, sender_name):
            preview_image_list = []
            for f in image_files:
                data = f.read()
                f.seek(0)
                ext = f.name.lower().split(".")[-1] if "." in f.name else "png"
                subtype = "jpeg" if ext in ("jpg", "jpeg") else ext
                b64 = base64.b64encode(data).decode()
                preview_image_list.append((f.name, f"data:image/{subtype};base64,{b64}"))
            preview_html = build_preview_html(first_name, body, sender_name, preview_image_list, body_is_html=False)
        else:
            preview_html = build_preview_html(first_name, body, sender_name, [], body_is_html=False)
        st.components.v1.html(preview_html, height=250, scrolling=True)

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
            image_data_list = [(f.name, f.read()) for f in image_files] if image_files else []
            num_images_needed = body.count("[IMAGE]")
            inline_list = image_data_list[:num_images_needed] if num_images_needed else []
            cc = cc_email.strip() if cc_email and cc_email.strip() else ""
            use_inline = bool(inline_list) and "[IMAGE]" in body

            progress_bar = st.progress(0)
            status_container = st.status("Sending emails...", expanded=True)
            results = []

            with status_container:
                for i, row in enumerate(rows):
                    to_email = str(row[email_col]).strip()
                    name = str(row.get(name_col, "")).strip() or "there"
                    subject_text = personalize_text(subject, name, sender_name)
                    if use_inline:
                        body_text = build_email_body_html(name, body, sender_name, inline_list, body_is_html=False)
                        err = send_email(
                            gmail_address,
                            gmail_app_password,
                            to_email,
                            subject_text,
                            body_text,
                            cc_email=cc,
                            inline_images=inline_list,
                        )
                    else:
                        body_text = build_email_body_html(name, body, sender_name, [], body_is_html=False)
                        err = send_email(
                            gmail_address,
                            gmail_app_password,
                            to_email,
                            subject_text,
                            body_text,
                            cc_email=cc,
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
