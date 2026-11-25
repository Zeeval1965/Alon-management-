"""Read-only Gmail access for incoming protocol emails.

Uses the Gmail API with a refresh token (scope gmail.readonly only). Nothing here writes to
Gmail or to the database. The caller decides what to do with the results.

Environment variables (set in Render, never in code):
  GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET, GMAIL_REFRESH_TOKEN
"""
import os

import requests

TOKEN_URL = "https://oauth2.googleapis.com/token"
API = "https://gmail.googleapis.com/gmail/v1/users/me"

# Each site is recognized by a phrase in the subject. A phrase is matched against the subject only.
# Adding a site = adding a line here (and a sender to SENDERS_ALLOWED if one is needed later).
SITE_SUBJECT_PHRASES = {
    "בית העמודים": "בית העמודים",
    "הלל הזקן 14": "הלל הזקן",
}


def route_subject(subject, phrases=None):
    """Return the site name for a subject, or None. A subject that matches two sites is not routed."""
    phrases = SITE_SUBJECT_PHRASES if phrases is None else phrases
    subject = (subject or "").strip()
    hits = [site for site, phrase in phrases.items() if phrase in subject]
    return hits[0] if len(hits) == 1 else None


def access_token():
    """Exchange the stored refresh token for a short-lived access token."""
    resp = requests.post(TOKEN_URL, data={
        "client_id": os.environ["GMAIL_CLIENT_ID"],
        "client_secret": os.environ["GMAIL_CLIENT_SECRET"],
        "refresh_token": os.environ["GMAIL_REFRESH_TOKEN"],
        "grant_type": "refresh_token",
    }, timeout=20)
    resp.raise_for_status()
    return resp.json()["access_token"]


def _headers(token):
    return {"Authorization": f"Bearer {token}"}


def list_recent_messages(token, max_results=50, query="newer_than:14d"):
    """Message ids matching a Gmail search query (last 14 days by default)."""
    resp = requests.get(f"{API}/messages", headers=_headers(token),
                        params={"q": query, "maxResults": max_results}, timeout=20)
    resp.raise_for_status()
    return [m["id"] for m in resp.json().get("messages", [])]


def message_summary(token, message_id):
    """Subject, sender, date and attachment names of one message. Body is not read here."""
    resp = requests.get(f"{API}/messages/{message_id}", headers=_headers(token),
                        params={"format": "full"}, timeout=20)
    resp.raise_for_status()
    msg = resp.json()
    headers = {h["name"].lower(): h["value"] for h in msg.get("payload", {}).get("headers", [])}
    attachments = []

    def walk(part):
        if part.get("filename"):
            attachments.append({"name": part["filename"], "attachmentId": part.get("body", {}).get("attachmentId")})
        for sub in part.get("parts", []) or []:
            walk(sub)

    walk(msg.get("payload", {}))
    return {
        "id": message_id,
        "subject": headers.get("subject", ""),
        "from": headers.get("from", ""),
        "to": headers.get("to", ""),
        "cc": headers.get("cc", ""),
        "date": headers.get("date", ""),
        "attachments": attachments,
    }


def recipients_of(info):
    """Recipients as stored on a meeting: 'אל: ...; עותק: ...' from the To and Cc headers. Empty if none."""
    parts = []
    if info.get("to"):
        parts.append("אל: " + info["to"])
    if info.get("cc"):
        parts.append("עותק: " + info["cc"])
    return "; ".join(parts)


def scan_for_sites(max_results=50):
    """Dry run: list recent messages, route each subject to a site. Writes nothing."""
    token = access_token()
    rows = []
    for mid in list_recent_messages(token, max_results=max_results):
        info = message_summary(token, mid)
        info["site"] = route_subject(info["subject"])
        rows.append(info)
    return rows
