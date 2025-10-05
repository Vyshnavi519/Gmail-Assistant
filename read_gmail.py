"""
read_and_save_emails.py

What this script does :
- Reads the latest N emails from your Gmail (using Gmail API)
- For each email:
  - creates a folder: emails/email_<messageId>/
  - saves metadata.json and body.txt
  - saves attachments (if any)
  - extracts text from attachments (PDF/DOCX/PNG/JPG/TXT) and saves extracted_<name>.txt
  - saves extracted_full.txt (all extracted text for that email)



(You already have credentials.json from Google Cloud; keep it out of git)
"""

from __future__ import print_function
import os
import json
import re
import io
import base64
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from PyPDF2 import PdfReader
from docx import Document
from PIL import Image
import pytesseract



# ---- Scope: we only read emails for now. When you want to create drafts, add compose scope. ----
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# If Tesseract is installed in a custom location, uncomment and set the path:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def sanitize_filename(name: str) -> str:
    """Remove problematic characters from filenames."""
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def get_email_body_from_payload(payload):
    """
    Try to retrieve a sensible plain-text body from the payload.
    Gmail messages can be nested; this walks parts recursively looking for text/plain.
    """
    # If payload itself has data (small messages)
    if payload.get("body", {}).get("data"):
        data = payload["body"]["data"]
        text = base64.urlsafe_b64decode(data.encode("UTF-8")).decode("utf-8", errors="replace")
        return text

    # If there are parts, search them
    parts = payload.get("parts", [])
    for part in parts:
        mime = part.get("mimeType", "")
        if mime == "text/plain" and part.get("body", {}).get("data"):
            data = part["body"]["data"]
            text = base64.urlsafe_b64decode(data.encode("UTF-8")).decode("utf-8", errors="replace")
            return text
        # recursive dive
        inner = get_email_body_from_payload(part)
        if inner:
            return inner
    return ""


def collect_all_parts(payload, out_list):
    """Recursively collect all parts (useful to find attachments no matter how deep)."""
    if not payload:
        return
    if payload.get("parts"):
        for p in payload["parts"]:
            collect_all_parts(p, out_list)
    else:
        # single part (could be an attachment or inline)
        out_list.append(payload)


def save_attachment(service, msg_id, part, save_dir):
    """
    Saves an attachment to disk. Handles parts where body has 'attachmentId' or 'data'.
    Returns tuple (filename, bytes_data)
    """
    filename = part.get("filename") or "unknown"
    filename = sanitize_filename(filename)
    body = part.get("body", {})

    data_bytes = None

    if "attachmentId" in body:
        # Official attachment reference — fetch via attachments.get
        att_id = body["attachmentId"]
        att = service.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
        raw = att.get("data")
        if raw:
            data_bytes = base64.urlsafe_b64decode(raw.encode("UTF-8"))
    elif body.get("data"):
        # Inline small attachment
        raw = body.get("data")
        data_bytes = base64.urlsafe_b64decode(raw.encode("UTF-8"))

    if data_bytes is None:
        return None, None

    file_path = os.path.join(save_dir, filename)
    with open(file_path, "wb") as f:
        f.write(data_bytes)

    return filename, data_bytes


def extract_text_from_bytes(data_bytes: bytes, filename: str) -> str:
    """Extract text from bytes based on the file extension. Returns extracted text or empty string."""
    lower = filename.lower()
    try:
        if lower.endswith(".pdf"):
            # PDF extraction with PyPDF2
            reader = PdfReader(io.BytesIO(data_bytes))
            pages_text = []
            for p in reader.pages:
                t = p.extract_text() or ""
                pages_text.append(t)
            return "\n".join(pages_text).strip()
        elif lower.endswith(".docx"):
            # Word .docx extraction
            doc = Document(io.BytesIO(data_bytes))
            return "\n".join([p.text for p in doc.paragraphs]).strip()
        elif lower.endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
            # Image -> OCR
            img = Image.open(io.BytesIO(data_bytes))
            # optional: convert to RGB to avoid issues
            if img.mode != "RGB":
                img = img.convert("RGB")
            text = pytesseract.image_to_string(img)
            return text.strip()
        elif lower.endswith(".txt"):
            return data_bytes.decode("utf-8", errors="replace")
        else:
            # Not supported natively — try to decode as text as a fallback
            try:
                return data_bytes.decode("utf-8", errors="replace")
            except Exception:
                return ""
    except Exception as e:
        return f"[Error extracting text: {e}]"


def save_email_folder(service, message):
    """
    Given a Gmail message resource (as returned by messages.get with format='full'),
    create a folder and save metadata, body, attachments and extracted text.
    """
    msg_id = message.get("id")
    thread_id = message.get("threadId", "")
    internal_date = message.get("internalDate")  # milliseconds-since-epoch as string
    date_readable = ""
    if internal_date:
        try:
            ts = int(internal_date) / 1000.0
            date_readable = datetime.utcfromtimestamp(ts).strftime("%Y%m%d_%H%M%S")
        except:
            date_readable = internal_date

    # Build a friendly folder name: emails/email_<msgid>_<date> (safe)
    folder_name = f"emails/email_{sanitize_filename(msg_id)}"
    if date_readable:
        folder_name += f"_{date_readable}"
    ensure_dir(folder_name)

    # Extract headers metadata
    headers = message.get("payload", {}).get("headers", [])
    meta = {}
    for h in headers:
        name = h.get("name", "").lower()
        if name in ["from", "to", "subject", "date", "message-id"]:
            meta[name] = h.get("value")
    meta["id"] = msg_id
    meta["threadId"] = thread_id
    meta["snippet"] = message.get("snippet", "")

    # Save metadata.json
    with open(os.path.join(folder_name, "metadata.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2, ensure_ascii=False)

    # Save body text (plain text if possible)
    body_text = get_email_body_from_payload(message.get("payload", {})) or ""
    with open(os.path.join(folder_name, "body.txt"), "w", encoding="utf-8") as f:
        f.write(body_text)

    # Now gather all parts and save attachments & extracted text
    parts_list = []
    collect_all_parts(message.get("payload", {}), parts_list)

    extracted_texts = []

    for part in parts_list:
        # decide if this part is an attachment (has filename) or has attachmentId
        filename = part.get("filename")
        body = part.get("body", {})
        if (filename and (body.get("attachmentId") or body.get("data"))):
            fname, bytes_data = save_attachment(service, msg_id, part, folder_name)
            if fname and bytes_data:
                # Extract text
                extracted = extract_text_from_bytes(bytes_data, fname)
                extracted_texts.append({
                    "attachment": fname,
                    "text": extracted
                })
                # Save extracted text file
                txt_name = f"extracted_{os.path.splitext(fname)[0]}.txt"
                with open(os.path.join(folder_name, txt_name), "w", encoding="utf-8") as tx:
                    tx.write(extracted)
    # Save combined extracted text file (concatenated)
    combined = "\n\n".join([et["text"] for et in extracted_texts if et["text"]])
    with open(os.path.join(folder_name, "extracted_full.txt"), "w", encoding="utf-8") as comb:
        comb.write(combined)

    print(f"Saved email -> {folder_name} (body + {len(extracted_texts)} extracted attachments)")
    return folder_name


def main():
    """Authenticate and process N latest emails, saving each to its own folder."""
    creds = None
    # token.json is created by the OAuth flow; it should not be in your git repo
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # This requires credentials.json in the same folder
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w", encoding="utf-8") as token_file:
            token_file.write(creds.to_json())

    service = build("gmail", "v1", credentials=creds)

    # Configure how many recent emails you want to process
    MAX_EMAILS = 10
    resp = service.users().messages().list(userId="me", maxResults=MAX_EMAILS).execute()
    messages = resp.get("messages", [])

    if not messages:
        print("No messages found.")
        return

    # For each message: fetch full message and save
    for m in messages:
        msg_id = m["id"]
        # fetch full message (format='full' is default) so we can access parts/attachments
        full = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
        save_email_folder(service, full)


if __name__ == "__main__":
    main()
