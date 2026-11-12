import json
import logging
import os
import re
import smtplib
import ssl
import sys
import urllib.error
import urllib.request
import uuid
from datetime import datetime, timedelta, timezone
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


LOGGER = logging.getLogger("dashboard_delivery")
TIMEZONE = timezone(timedelta(hours=5, minutes=30), name="IST")
DEFAULT_DASHBOARD_URL = "https://offload-dashboard.onrender.com/"
DEFAULT_EMAIL_FROM = "saritha.m@criticalog.com"
DEFAULT_WHATSAPP_API_VERSION = "v22.0"
WHATSAPP_BASE_URL = "https://graph.facebook.com"


class DeliveryError(Exception):
    pass


def _required_env(name):
    value = os.environ.get(name, "").strip()
    if not value:
        raise DeliveryError(f"Missing required setting: {name}.")
    return value


def _parse_recipients(value, setting_name):
    recipients = [address.strip() for address in value.split(",") if address.strip()]
    if not recipients:
        raise DeliveryError(f"Configure at least one recipient in {setting_name}.")
    return recipients


def render_dashboard(url=None):
    from playwright.sync_api import Error as PlaywrightError
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
    from playwright.sync_api import sync_playwright

    dashboard_url = url or os.environ.get("DASHBOARD_URL", DEFAULT_DASHBOARD_URL)
    chromium_path = os.environ.get("CHROMIUM_PATH")
    try:
        with sync_playwright() as playwright:
            launch_options = {"headless": True, "args": ["--no-sandbox"]}
            if chromium_path:
                launch_options["executable_path"] = chromium_path
            browser = playwright.chromium.launch(**launch_options)
            try:
                page = browser.new_page(
                    viewport={"width": 1440, "height": 1000},
                    device_scale_factor=1,
                )
                response = page.goto(
                    dashboard_url,
                    wait_until="networkidle",
                    timeout=180_000,
                )
                if response is None or not response.ok:
                    status = response.status if response is not None else "no response"
                    raise DeliveryError(f"Dashboard returned HTTP {status}.")
                page.wait_for_function(
                    "() => Array.isArray(window.offloadWorkbookRecords) "
                    "&& window.offloadWorkbookRecords.length > 0",
                    timeout=30_000,
                )
                return page.screenshot(full_page=True, type="png")
            finally:
                browser.close()
    except DeliveryError:
        raise
    except PlaywrightTimeoutError as exc:
        raise DeliveryError("Timed out while loading the dashboard for its snapshot.") from exc
    except PlaywrightError as exc:
        raise DeliveryError("Could not render the dashboard snapshot in Chromium.") from exc


def build_email(snapshot, recipients, sender, now=None):
    now = now or datetime.now(TIMEZONE)
    date_label = now.strftime("%d %b %Y")
    message = MIMEMultipart("related")
    message["Subject"] = f"Offload dashboard snapshot - {date_label}"
    message["From"] = sender
    message["To"] = "Undisclosed recipients"
    message["Date"] = now.strftime("%a, %d %b %Y %H:%M:%S %z")

    content = MIMEMultipart("alternative")
    content.attach(MIMEText(
        f"Daily offload dashboard snapshot for {date_label} is attached.",
        "plain",
        "utf-8",
    ))
    content.attach(MIMEText(
        "<p>Daily offload dashboard snapshot for "
        f"{date_label}.</p><img src=\"cid:dashboard-snapshot\" "
        "alt=\"Offload dashboard snapshot\" style=\"width:100%;height:auto\">",
        "html",
        "utf-8",
    ))
    message.attach(content)

    image = MIMEImage(snapshot, _subtype="png")
    image.add_header("Content-ID", "<dashboard-snapshot>")
    image.add_header(
        "Content-Disposition",
        "inline",
        filename=f"offload-dashboard-{now:%Y-%m-%d}.png",
    )
    message.attach(image)
    return message, recipients


def send_email(snapshot):
    host = _required_env("SMTP_HOST")
    recipients = _parse_recipients(
        _required_env("EMAIL_RECIPIENTS"),
        "EMAIL_RECIPIENTS",
    )
    sender = os.environ.get("EMAIL_FROM", DEFAULT_EMAIL_FROM).strip() or DEFAULT_EMAIL_FROM
    port = int(os.environ.get("SMTP_PORT", "587"))
    username = os.environ.get("SMTP_USERNAME", "")
    password = os.environ.get("SMTP_PASSWORD", "")
    use_starttls = os.environ.get("SMTP_STARTTLS", "true").strip().lower() in {
        "1", "true", "yes",
    }
    if bool(username) != bool(password):
        raise DeliveryError("Configure both SMTP_USERNAME and SMTP_PASSWORD, or neither.")

    message, envelope_recipients = build_email(snapshot, recipients, sender)
    try:
        if port == 465:
            with smtplib.SMTP_SSL(
                host,
                port,
                timeout=30,
                context=ssl.create_default_context(),
            ) as smtp:
                if username:
                    smtp.login(username, password)
                smtp.sendmail(sender, envelope_recipients, message.as_string())
        else:
            with smtplib.SMTP(host, port, timeout=30) as smtp:
                if use_starttls:
                    smtp.starttls(context=ssl.create_default_context())
                if username:
                    smtp.login(username, password)
                smtp.sendmail(sender, envelope_recipients, message.as_string())
    except (OSError, smtplib.SMTPException) as exc:
        raise DeliveryError(
            f"SMTP delivery failed ({type(exc).__name__}); check relay settings and access."
        ) from exc
    LOGGER.info("Email snapshot sent to %d recipient(s).", len(recipients))


def _whatsapp_settings():
    names = (
        "WHATSAPP_ACCESS_TOKEN",
        "WHATSAPP_PHONE_NUMBER_ID",
        "WHATSAPP_RECIPIENTS",
        "WHATSAPP_TEMPLATE_NAME",
    )
    values = {name: os.environ.get(name, "").strip() for name in names}
    configured = [bool(value) for value in values.values()]
    if not any(configured):
        LOGGER.warning("WhatsApp is not configured; sending the email only.")
        return None
    if not all(configured):
        missing = ", ".join(name for name, value in values.items() if not value)
        raise DeliveryError(f"WhatsApp configuration is incomplete: {missing}.")
    if os.environ.get("WHATSAPP_RECIPIENTS_OPTED_IN", "").strip().lower() != "true":
        raise DeliveryError(
            "Set WHATSAPP_RECIPIENTS_OPTED_IN=true only after all recipients have opted in."
        )

    api_version = os.environ.get(
        "WHATSAPP_API_VERSION",
        DEFAULT_WHATSAPP_API_VERSION,
    ).strip()
    if not re.fullmatch(r"v\d+\.\d+", api_version):
        raise DeliveryError("WHATSAPP_API_VERSION must look like v22.0.")

    recipients = _parse_recipients(values["WHATSAPP_RECIPIENTS"], "WHATSAPP_RECIPIENTS")
    normalized_recipients = []
    for recipient in recipients:
        if not re.fullmatch(r"\+?[1-9]\d{6,14}", recipient):
            raise DeliveryError(
                "WhatsApp recipients must be phone numbers in international format."
            )
        normalized_recipients.append(recipient.lstrip("+"))

    if not re.fullmatch(r"\d+", values["WHATSAPP_PHONE_NUMBER_ID"]):
        raise DeliveryError("WHATSAPP_PHONE_NUMBER_ID must contain only digits.")

    return {
        **values,
        "api_version": api_version,
        "recipients": normalized_recipients,
        "language": os.environ.get("WHATSAPP_TEMPLATE_LANGUAGE", "en_US").strip(),
    }


def _read_whatsapp_json(request):
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise DeliveryError(
            f"WhatsApp Cloud API returned HTTP {exc.code}."
        ) from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise DeliveryError(
            f"WhatsApp Cloud API request failed ({type(exc).__name__})."
        ) from exc


def _upload_whatsapp_image(snapshot, settings):
    boundary = uuid.uuid4().hex
    fields = [
        ("messaging_product", "whatsapp"),
        ("type", "image/png"),
    ]
    chunks = []
    for name, value in fields:
        chunks.extend((
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode(),
            value.encode(),
            b"\r\n",
        ))
    chunks.extend((
        f"--{boundary}\r\n".encode(),
        b'Content-Disposition: form-data; name="file"; filename="dashboard.png"\r\n',
        b"Content-Type: image/png\r\n\r\n",
        snapshot,
        b"\r\n",
        f"--{boundary}--\r\n".encode(),
    ))
    url = (
        f"{WHATSAPP_BASE_URL}/{settings['api_version']}/"
        f"{settings['WHATSAPP_PHONE_NUMBER_ID']}/media"
    )
    request = urllib.request.Request(
        url,
        data=b"".join(chunks),
        headers={
            "Authorization": f"Bearer {settings['WHATSAPP_ACCESS_TOKEN']}",
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        },
        method="POST",
    )
    response = _read_whatsapp_json(request)
    media_id = response.get("id")
    if not isinstance(media_id, str) or not media_id:
        raise DeliveryError("WhatsApp Cloud API did not return an uploaded image ID.")
    return media_id


def send_whatsapp(snapshot):
    settings = _whatsapp_settings()
    if settings is None:
        return

    media_id = _upload_whatsapp_image(snapshot, settings)
    endpoint = (
        f"{WHATSAPP_BASE_URL}/{settings['api_version']}/"
        f"{settings['WHATSAPP_PHONE_NUMBER_ID']}/messages"
    )
    failures = 0
    for recipient in settings["recipients"]:
        payload = {
            "messaging_product": "whatsapp",
            "to": recipient,
            "type": "template",
            "template": {
                "name": settings["WHATSAPP_TEMPLATE_NAME"],
                "language": {"code": settings["language"]},
                "components": [{
                    "type": "header",
                    "parameters": [{
                        "type": "image",
                        "image": {"id": media_id},
                    }],
                }],
            },
        }
        request = urllib.request.Request(
            endpoint,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {settings['WHATSAPP_ACCESS_TOKEN']}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            _read_whatsapp_json(request)
        except DeliveryError:
            failures += 1
    if failures:
        raise DeliveryError(
            f"WhatsApp delivery failed for {failures} of "
            f"{len(settings['recipients'])} recipient(s)."
        )
    LOGGER.info("WhatsApp snapshot sent to %d recipient(s).", len(settings["recipients"]))


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        stream=sys.stdout,
    )
    try:
        snapshot = render_dashboard()
        send_email(snapshot)
        send_whatsapp(snapshot)
    except (DeliveryError, ValueError) as exc:
        LOGGER.error("%s", exc)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
