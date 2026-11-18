import json
import logging
import os
import sys
import types
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

import dashboard_delivery
from dashboard_delivery import DeliveryError


PNG_DATA = b"\x89PNG\r\n\x1a\npng-data"


class DashboardDeliveryTests(unittest.TestCase):
    def test_render_dashboard_waits_for_data_and_captures_full_page(self):
        calls = {}

        class FakePage:
            def goto(self, url, **kwargs):
                calls["url"] = url
                calls["goto"] = kwargs
                return SimpleNamespace(ok=True, status=200)

            def wait_for_function(self, expression, **kwargs):
                calls["wait"] = (expression, kwargs)

            def screenshot(self, **kwargs):
                calls["screenshot"] = kwargs
                return PNG_DATA

        class FakeBrowser:
            def new_page(self, **kwargs):
                calls["viewport"] = kwargs
                return FakePage()

            def close(self):
                calls["closed"] = True

        class FakeChromium:
            def launch(self, **kwargs):
                calls["launch"] = kwargs
                return FakeBrowser()

        class FakePlaywright:
            chromium = FakeChromium()

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        sync_api = types.ModuleType("playwright.sync_api")
        sync_api.Error = type("PlaywrightError", (Exception,), {})
        sync_api.TimeoutError = type("PlaywrightTimeoutError", (Exception,), {})
        sync_api.sync_playwright = FakePlaywright
        playwright = types.ModuleType("playwright")
        playwright.__path__ = []
        with patch.dict(sys.modules, {
            "playwright": playwright,
            "playwright.sync_api": sync_api,
        }):
            image = dashboard_delivery.render_dashboard("https://dashboard.example.com/")

        self.assertEqual(image, PNG_DATA)
        self.assertEqual(calls["url"], "https://dashboard.example.com/")
        self.assertEqual(calls["goto"]["wait_until"], "networkidle")
        self.assertEqual(calls["screenshot"], {"full_page": True, "type": "png"})
        self.assertTrue(calls["closed"])

    def test_build_email_hides_recipients_and_embeds_png(self):
        ist = timezone(timedelta(hours=5, minutes=30), name="IST")
        now = datetime(2026, 6, 30, 14, 30, tzinfo=ist)
        message, recipients = dashboard_delivery.build_email(
            PNG_DATA,
            ["one@example.com", "two@example.com"],
            "sender@example.com",
            now,
        )

        self.assertEqual(recipients, ["one@example.com", "two@example.com"])
        self.assertEqual(message["To"], "Undisclosed recipients")
        self.assertNotIn("one@example.com", message.as_string())
        self.assertNotIn("two@example.com", message.as_string())
        image = next(part for part in message.walk() if part.get_content_type() == "image/png")
        self.assertEqual(image.get_payload(decode=True), PNG_DATA)
        self.assertEqual(image["Content-ID"], "<dashboard-snapshot>")
        self.assertIn("30 Jun 2026", message["Subject"])

    def test_send_email_uses_starttls_and_bcc_envelope(self):
        environment = {
            "SMTP_HOST": "relay.example.com",
            "SMTP_PORT": "587",
            "SMTP_STARTTLS": "true",
            "EMAIL_RECIPIENTS": "one@example.com,two@example.com",
            "EMAIL_FROM": "sender@example.com",
        }
        with patch.dict(os.environ, environment, clear=True), \
                patch("dashboard_delivery.smtplib.SMTP") as smtp_factory:
            smtp = smtp_factory.return_value.__enter__.return_value
            dashboard_delivery.send_email(PNG_DATA)

        smtp_factory.assert_called_once_with("relay.example.com", 587, timeout=30)
        smtp.starttls.assert_called_once()
        sent = smtp.sendmail.call_args.args
        self.assertEqual(sent[0], "sender@example.com")
        self.assertEqual(sent[1], ["one@example.com", "two@example.com"])
        self.assertNotIn("one@example.com", sent[2])

    def test_sender_defaults_to_smtp_username_and_is_required(self):
        environment = {"SMTP_HOST": "smtp.example.com", "EMAIL_RECIPIENTS": "one@example.com",
                       "SMTP_USERNAME": "me@example.com", "SMTP_PASSWORD": "pw", "SMTP_STARTTLS": "false"}
        with patch.dict(os.environ, environment, clear=True), \
                patch("dashboard_delivery.smtplib.SMTP") as smtp_factory:
            smtp = smtp_factory.return_value.__enter__.return_value
            smtp.has_extn.return_value = True
            dashboard_delivery.send_email(PNG_DATA)
        self.assertEqual(smtp.sendmail.call_args[0][0], "me@example.com")
        with patch.dict(os.environ, {"SMTP_HOST": "smtp.example.com", "EMAIL_RECIPIENTS": "one@example.com"},
                        clear=True):
            with self.assertRaisesRegex(DeliveryError, "EMAIL_FROM"):
                dashboard_delivery.send_email(PNG_DATA)

    def test_send_email_rejects_unpaired_smtp_credentials(self):
        with patch.dict(os.environ, {
            "SMTP_HOST": "relay.example.com",
            "EMAIL_RECIPIENTS": "one@example.com",
            "SMTP_USERNAME": "user",
        }, clear=True):
            with self.assertRaisesRegex(DeliveryError, "Configure both SMTP_USERNAME"):
                dashboard_delivery.send_email(PNG_DATA)

    def test_main_loads_dotenv_without_overriding_process_environment(self):
        with TemporaryDirectory() as temp_dir:
            env_file = Path(temp_dir) / ".env"
            env_file.write_text(
                "SMTP_HOST=from-file.example.com\nSMTP_PASSWORD=file-password\n",
                encoding="utf-8",
            )
            with patch.object(dashboard_delivery, "ENV_FILE", env_file), \
                    patch.object(dashboard_delivery, "LOG_DIRECTORY", Path(temp_dir)), \
                    patch.object(dashboard_delivery, "LOG_PATH", Path(temp_dir) / "email.log"), \
                    patch.dict(os.environ, {"SMTP_HOST": "from-process.example.com"}, clear=True), \
                    patch.object(dashboard_delivery, "render_dashboard", return_value=PNG_DATA), \
                    patch.object(dashboard_delivery, "send_email"), \
                    patch.object(dashboard_delivery, "send_whatsapp"):
                self.assertEqual(dashboard_delivery.main([]), 0)
                root = logging.getLogger()  # detach the temp log file before the folder is removed
                for handler in list(root.handlers):
                    root.removeHandler(handler)
                    handler.close()
                self.assertEqual(os.environ["SMTP_HOST"], "from-process.example.com")
                self.assertEqual(os.environ["SMTP_PASSWORD"], "file-password")

    def test_whatsapp_unconfigured_is_optional(self):
        with patch.dict(os.environ, {
            "WHATSAPP_RECIPIENTS_OPTED_IN": "false",
        }, clear=True), \
                patch.object(dashboard_delivery.LOGGER, "warning") as warning:
            self.assertIsNone(dashboard_delivery._whatsapp_settings())
        warning.assert_called_once()

    def test_whatsapp_rejects_partial_configuration(self):
        with patch.dict(os.environ, {
            "WHATSAPP_ACCESS_TOKEN": "private-token",
        }, clear=True):
            with self.assertRaisesRegex(DeliveryError, "configuration is incomplete"):
                dashboard_delivery._whatsapp_settings()

    def test_whatsapp_requires_recipient_opt_in(self):
        environment = {
            "WHATSAPP_ACCESS_TOKEN": "private-token",
            "WHATSAPP_PHONE_NUMBER_ID": "123456789",
            "WHATSAPP_RECIPIENTS": "+15551234567",
            "WHATSAPP_TEMPLATE_NAME": "daily_dashboard",
            "WHATSAPP_RECIPIENTS_OPTED_IN": "false",
        }
        with patch.dict(os.environ, environment, clear=True):
            with self.assertRaisesRegex(DeliveryError, "only after all recipients have opted in"):
                dashboard_delivery._whatsapp_settings()

    def test_whatsapp_template_sends_uploaded_image_to_each_recipient(self):
        settings = {
            "WHATSAPP_ACCESS_TOKEN": "private-token",
            "WHATSAPP_PHONE_NUMBER_ID": "123456789",
            "WHATSAPP_TEMPLATE_NAME": "daily_dashboard",
            "WHATSAPP_RECIPIENTS_OPTED_IN": "true",
            "api_version": "v22.0",
            "language": "en_US",
            "recipients": ["15551234567", "15557654321"],
        }
        requests = []

        def capture_request(request):
            requests.append(request)
            return {"messages": [{"id": "message-id"}]}

        with patch.object(dashboard_delivery, "_whatsapp_settings", return_value=settings), \
                patch.object(dashboard_delivery, "_upload_whatsapp_image", return_value="media-id"), \
                patch.object(dashboard_delivery, "_read_whatsapp_json", side_effect=capture_request):
            dashboard_delivery.send_whatsapp(PNG_DATA)

        self.assertEqual(len(requests), 2)
        payload = json.loads(requests[0].data.decode("utf-8"))
        self.assertEqual(payload["to"], "15551234567")
        self.assertEqual(payload["template"]["name"], "daily_dashboard")
        self.assertEqual(
            payload["template"]["components"][0]["parameters"][0]["image"]["id"],
            "media-id",
        )
        self.assertEqual(requests[0].get_header("Authorization"), "Bearer private-token")

    def test_whatsapp_recipients_must_be_in_international_format(self):
        environment = {
            "WHATSAPP_ACCESS_TOKEN": "private-token",
            "WHATSAPP_PHONE_NUMBER_ID": "123456789",
            "WHATSAPP_RECIPIENTS": "not-a-number",
            "WHATSAPP_TEMPLATE_NAME": "daily_dashboard",
            "WHATSAPP_RECIPIENTS_OPTED_IN": "true",
        }
        with patch.dict(os.environ, environment, clear=True):
            with self.assertRaisesRegex(DeliveryError, "international format"):
                dashboard_delivery._whatsapp_settings()


if __name__ == "__main__":
    unittest.main()
