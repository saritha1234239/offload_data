"""Check the SMTP settings in .env step by step, without sending anything by default.

    python smtp_check.py                          # connect, encrypt, sign in
    python smtp_check.py --host smtp.office365.com   # try another server without editing .env
    python smtp_check.py --send                   # also send a short test email to EMAIL_FROM
    python smtp_check.py --send --to someone@example.com

The password is read from .env and never printed.
"""
import argparse
import os
import smtplib
import socket
import ssl
import sys
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

from dotenv import load_dotenv

ENV_FILE = Path(__file__).with_name(".env")

# Common Microsoft 365 / Exchange Online replies and what to do about them.
KNOWN_REPLIES = {
    "5.7.139": "Sign-in blocked by Microsoft 365: SMTP AUTH is disabled for this mailbox or the whole "
               "organisation, Security Defaults / Conditional Access block it, or (on smtp-hve) this is "
               "not an HVE account. Ask IT to enable Authenticated SMTP for the mailbox.",
    "5.7.3": "Wrong username or password. Use the full mailbox address; with MFA use an app password.",
    "5.7.8": "Wrong username or password.",
    "5.2.1": "The mailbox cannot be used: shared mailbox, no Exchange licence, or sign-in blocked.",
    "5.7.57": "The server requires sign-in before sending (set SMTP_USERNAME and SMTP_PASSWORD).",
    "5.7.60": "The account may not send as EMAIL_FROM. Use the same address as SMTP_USERNAME, "
              "or ask IT for 'Send As' permission.",
    "5.2.0": "The account may not send as EMAIL_FROM (SendAsDenied). Use the SMTP_USERNAME address.",
}


def explain(reply):
    text = reply.decode("utf-8", "replace") if isinstance(reply, bytes) else str(reply)
    hint = next((h for code, h in KNOWN_REPLIES.items() if code in text), "")
    return text.strip(), hint


def ok(message):
    print(f"  [OK]   {message}")


def fail(message, hint=""):
    print(f"  [FAIL] {message}")
    if hint:
        print(f"         -> {hint}")


def main():
    parser = argparse.ArgumentParser(description="Check the SMTP settings in .env.")
    parser.add_argument("--host", help="Server to test instead of SMTP_HOST or SMTP_SERVER.")
    parser.add_argument("--port", type=int, help="Port to test instead of SMTP_PORT.")
    send_options = parser.add_mutually_exclusive_group()
    send_options.add_argument(
        "--send",
        action="store_true",
        help="Send a test email after checking SMTP (the default).",
    )
    send_options.add_argument(
        "--check-only",
        action="store_true",
        help="Check SMTP settings without sending a test email.",
    )
    parser.add_argument(
        "--to",
        help="Recipient(s) for the test email (default: EMAIL_RECIPIENTS, then EMAIL_FROM / SMTP_USERNAME).",
    )
    args = parser.parse_args()

    load_dotenv(ENV_FILE, override=False)
    host = (args.host or "").strip()
    if not host:
        host = os.environ.get("SMTP_HOST", "").strip()
    if not host:
        host = os.environ.get("SMTP_SERVER", "").strip()
    try:
        port = args.port or int(os.environ.get("SMTP_PORT", "587") or 587)
    except ValueError:
        fail("SMTP_PORT must be a number.")
        return 1
    starttls = os.environ.get("SMTP_STARTTLS", "true").strip().lower() in {"1", "true", "yes"}
    username = os.environ.get("SMTP_USERNAME", "").strip()
    password = os.environ.get("SMTP_PASSWORD", "")
    sender = os.environ.get("EMAIL_FROM", "").strip() or username
    recipients = os.environ.get("EMAIL_RECIPIENTS", "")

    print("Settings (from .env):")
    print(f"  server      {host or '(missing SMTP_HOST and SMTP_SERVER)'}:{port}")
    print(f"  STARTTLS    {'yes' if starttls else 'no'}{'  (port 465 uses direct TLS)' if port == 465 else ''}")
    print(f"  username    {'(configured)' if username else '(none)'}")
    print(f"  password    {'(set)' if password else '(missing)'}")
    print(f"  from        {'(configured)' if sender else '(missing)'}")
    print(f"  recipients  {len([r for r in recipients.split(',') if r.strip()])} in EMAIL_RECIPIENTS")
    if "smtp-hve" in host:
        print("  note        smtp-hve is Microsoft High Volume Email: only IT-created HVE accounts can sign in.")
    if not host:
        fail("Set SMTP_HOST or SMTP_SERVER in .env.")
        return 1
    if not 1 <= port <= 65535:
        fail("SMTP_PORT must be between 1 and 65535.")
        return 1
    if bool(username) != bool(password):
        fail("Set both SMTP_USERNAME and SMTP_PASSWORD, or leave both blank.")
        return 1
    if username and port != 465 and not starttls:
        fail(
            "Refusing password sign-in without TLS.",
            "Enable STARTTLS or use port 465 for implicit TLS.",
        )
        return 1
    print()

    print("1. Reach the server")
    try:
        addresses = sorted({info[4][0] for info in socket.getaddrinfo(host, port, proto=socket.IPPROTO_TCP)})
        ok(f"{host} resolves to {', '.join(addresses[:3])}")
    except socket.gaierror as exc:
        fail(f"Cannot resolve {host}: {exc}", "Check the server name and the internet connection.")
        return 1

    context = ssl.create_default_context()
    try:
        if port == 465:
            smtp = smtplib.SMTP_SSL(host, port, timeout=30, context=context)
        else:
            smtp = smtplib.SMTP(host, port, timeout=30)
    except (OSError, smtplib.SMTPException) as exc:
        fail(f"Cannot connect to {host}:{port}: {exc}",
             "Port blocked by the network/firewall or wrong port. Microsoft 365 uses 587.")
        return 1

    with smtp:
        code, banner = smtp.ehlo()
        ok(f"Connected (SMTP response {code})")

        print("2. Encrypt the connection")
        if port == 465:
            ok("Direct TLS (port 465)")
        elif starttls:
            if not smtp.has_extn("starttls"):
                fail("Server does not offer STARTTLS.")
                return 1
            try:
                smtp.starttls(context=context)
                smtp.ehlo()
                ok("STARTTLS enabled")
            except (ssl.SSLError, smtplib.SMTPException) as exc:
                fail(f"STARTTLS failed: {exc}")
                return 1
        else:
            print("  [SKIP] SMTP_STARTTLS is false; Microsoft 365 only allows sign-in after STARTTLS.")

        print("3. Password sign-in offered?")
        auth = smtp.esmtp_features.get("auth", "").strip()
        if auth:
            ok(f"Yes: {auth}")
        else:
            fail("The server does not offer sign-in on this connection.",
                 "Wrong server (a relay/MX endpoint) or STARTTLS not enabled. Use smtp.office365.com:587.")
            if username:
                return 1

        print("4. Sign in")
        if not username:
            print("  [SKIP] No SMTP_USERNAME: sending without sign-in (relay).")
        else:
            try:
                smtp.login(username, password)
                ok("Password sign-in succeeded")
            except smtplib.SMTPAuthenticationError as exc:
                text, hint = explain(exc.smtp_error)
                fail(f"Sign-in rejected ({exc.smtp_code}): {text}", hint)
                return 1
            except smtplib.SMTPException as exc:
                fail(f"Sign-in failed: {exc}")
                return 1

        print("5. Send a test email")
        if args.check_only:
            print("  [SKIP] --check-only specified; no test email sent.")
            print("\nResult: settings work up to sign-in.")
            return 0
        to = (args.to or recipients or sender).strip()
        if not to:
            fail("No recipient: set EMAIL_RECIPIENTS, EMAIL_FROM, or pass --to.")
            return 1
        message = EmailMessage()
        message["Subject"] = f"SMTP test from offload dashboard ({datetime.now():%d %b %Y %H:%M})"
        message["From"] = sender
        message["To"] = to
        message.set_content("This is a test email from smtp_check.py. The SMTP settings in .env work.")
        try:
            refused = smtp.send_message(message)
        except smtplib.SMTPResponseException as exc:
            text, hint = explain(exc.smtp_error)
            fail(f"Send rejected ({exc.smtp_code}): {text}", hint)
            return 1
        except smtplib.SMTPException as exc:
            fail(f"Send failed: {exc}")
            return 1
        if refused:
            fail(f"Some recipients were refused: {', '.join(refused)}")
            return 1
        ok("Test email accepted by the server. Check the recipient's inbox (and junk folder).")
    print("\nResult: everything works.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
