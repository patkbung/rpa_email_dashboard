"""
test_imap.py — Quick IMAP diagnostic script.

Run with:
    python3 test_imap.py

What it checks:
  1. Loads .env credentials.
  2. Connects and authenticates with Gmail.
  3. Lists ALL mailboxes → shows exact spam folder name.
  4. Tries to open INBOX and [Gmail]/Spam.
  5. Counts emails matching EMAIL_SEARCH_KEYWORD in each.
"""

import imaplib
import os
import re
# pyrefly: ignore [missing-import]
from dotenv import load_dotenv

load_dotenv()

EMAIL_ADDRESS        = os.getenv("EMAIL_ADDRESS", "").strip()
EMAIL_APP_PASSWORD   = os.getenv("EMAIL_APP_PASSWORD", "").strip()
EMAIL_SEARCH_KEYWORD = os.getenv("EMAIL_SEARCH_KEYWORD", "").strip()

IMAP_HOST = "imap.gmail.com"
IMAP_PORT = 993

SEP = "─" * 55

def main():
    print(SEP)
    print("  Gmail IMAP Diagnostic")
    print(SEP)

    # ── Step 1: Validate .env ────────────────────────────────────────
    print("\n[1] Checking .env …")
    if not EMAIL_ADDRESS or EMAIL_ADDRESS == "your_gmail@gmail.com":
        print("  [✘] EMAIL_ADDRESS is not set in .env")
        return
    if not EMAIL_APP_PASSWORD or EMAIL_APP_PASSWORD == "your_16_char_app_password":
        print("  [✘] EMAIL_APP_PASSWORD is not set in .env")
        print("       → Go to: https://myaccount.google.com/apppasswords")
        return
    print(f"  [✔] EMAIL_ADDRESS        = {EMAIL_ADDRESS}")
    print(f"  [✔] EMAIL_APP_PASSWORD   = {'*' * len(EMAIL_APP_PASSWORD)} (hidden)")
    print(f"  [✔] EMAIL_SEARCH_KEYWORD = {EMAIL_SEARCH_KEYWORD or '(empty)'}")

    # ── Step 2: Connect ──────────────────────────────────────────────
    print(f"\n[2] Connecting to {IMAP_HOST}:{IMAP_PORT} …")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    except OSError as exc:
        print(f"  [✘] Network error: {exc}")
        return

    # ── Step 3: Authenticate ─────────────────────────────────────────
    print("[3] Authenticating …")
    try:
        mail.login(EMAIL_ADDRESS, EMAIL_APP_PASSWORD)
        print("  [✔] Login successful!")
    except imaplib.IMAP4.error as exc:
        print(f"  [✘] Login FAILED: {exc}")
        print("\n  Checklist:")
        print("  • Is 2-Step Verification ON for this Google account?")
        print("  • Did you generate an App Password (not your normal password)?")
        print("  • Is IMAP enabled? Gmail Settings → See all settings → Forwarding and POP/IMAP")
        mail.logout()
        return

    # ── Step 4: List all mailboxes ───────────────────────────────────
    print("\n[4] Listing all mailboxes …")
    status, mailbox_list = mail.list()
    if status == "OK":
        print(f"  Found {len(mailbox_list)} mailbox(es):\n")
        for item in mailbox_list:
            if item:
                decoded = item.decode("utf-8", errors="replace") if isinstance(item, bytes) else item
                # Highlight Gmail system folders
                marker = "  ← SPAM?" if any(k in decoded.lower() for k in ["spam", "junk", "สแปม"]) else ""
                print(f"    {decoded}{marker}")
    else:
        print(f"  [✘] LIST failed: {status}")

    # ── Step 5: Try opening specific folders ─────────────────────────
    print("\n[5] Testing mailbox access …")
    test_folders = ["INBOX", "[Gmail]/Spam", "[Gmail]/All Mail"]

    for folder in test_folders:
        quoted = f'"{folder}"' if any(c in folder for c in " []()") else folder
        try:
            status, data = mail.select(quoted, readonly=True)
            if status == "OK":
                count = data[0].decode() if data and data[0] else "?"
                print(f"  [✔] '{folder}' → {count} message(s)")
            else:
                msg = data[0].decode() if data and data[0] else "unknown"
                print(f"  [✘] '{folder}' → {status}: {msg}")
        except imaplib.IMAP4.error as exc:
            print(f"  [✘] '{folder}' → error: {exc}")

    # ── Step 6: Search keyword in INBOX and Spam ─────────────────────
    if EMAIL_SEARCH_KEYWORD:
        print(f"\n[6] Searching for subject '{EMAIL_SEARCH_KEYWORD}' …")
        search_folders = ["INBOX", "[Gmail]/Spam"]
        criterion      = f'(SUBJECT "{EMAIL_SEARCH_KEYWORD}")'

        for folder in search_folders:
            quoted = f'"{folder}"' if any(c in folder for c in " []()") else folder
            try:
                status, _ = mail.select(quoted, readonly=True)
                if status != "OK":
                    print(f"  [⚠] Cannot open '{folder}'")
                    continue
                status, data = mail.search(None, criterion)
                ids = data[0].split() if status == "OK" and data and data[0] else []
                print(f"  {'[✔]' if ids else '[ℹ]'} '{folder}' → {len(ids)} matching email(s)")
            except imaplib.IMAP4.error as exc:
                print(f"  [✘] '{folder}' → search error: {exc}")
    else:
        print("\n[6] Skipping search — EMAIL_SEARCH_KEYWORD is empty in .env")

    # ── Done ─────────────────────────────────────────────────────────
    mail.logout()
    print(f"\n{SEP}")
    print("  Diagnostic complete.")
    print(SEP)


if __name__ == "__main__":
    main()
