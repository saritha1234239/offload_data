import os
import sys
import unittest
from contextlib import redirect_stdout
from io import StringIO
from unittest.mock import patch

import smtp_check


class SmtpCheckTests(unittest.TestCase):
    def test_refuses_password_auth_without_encryption(self):
        environment = {
            "SMTP_HOST": "smtp.example.com",
            "SMTP_PORT": "587",
            "SMTP_STARTTLS": "false",
            "SMTP_USERNAME": "user@example.com",
            "SMTP_PASSWORD": "private-password",
        }
        output = StringIO()
        with patch.dict(os.environ, environment, clear=True), \
                patch.object(smtp_check, "load_dotenv"), \
                patch.object(sys, "argv", ["smtp_check.py"]), \
                patch.object(smtp_check.socket, "getaddrinfo") as resolve, \
                patch.object(smtp_check.smtplib, "SMTP") as smtp_factory, \
                redirect_stdout(output):
            result = smtp_check.main()

        self.assertEqual(result, 1)
        resolve.assert_not_called()
        smtp_factory.assert_not_called()
        self.assertIn("Refusing password sign-in without TLS", output.getvalue())
        self.assertNotIn("private-password", output.getvalue())
        self.assertNotIn("user@example.com", output.getvalue())

    def test_rejects_unpaired_credentials(self):
        with patch.dict(os.environ, {
            "SMTP_HOST": "smtp.example.com",
            "SMTP_USERNAME": "user@example.com",
        }, clear=True), \
                patch.object(smtp_check, "load_dotenv"), \
                patch.object(sys, "argv", ["smtp_check.py"]), \
                patch.object(smtp_check.socket, "getaddrinfo") as resolve, \
                redirect_stdout(StringIO()):
            result = smtp_check.main()

        self.assertEqual(result, 1)
        resolve.assert_not_called()

    def test_uses_smtp_server_when_smtp_host_is_not_set(self):
        environment = {
            "SMTP_SERVER": "relay.example.com",
            "SMTP_STARTTLS": "false",
        }
        with patch.dict(os.environ, environment, clear=True), \
                patch.object(smtp_check, "load_dotenv"), \
                patch.object(sys, "argv", ["smtp_check.py", "--check-only"]), \
                patch.object(smtp_check.socket, "getaddrinfo", return_value=[
                    (None, None, None, None, ("127.0.0.1", 587)),
                ]), \
                patch.object(smtp_check.smtplib, "SMTP") as smtp_factory, \
                redirect_stdout(StringIO()):
            smtp = smtp_factory.return_value
            smtp.ehlo.return_value = (250, b"ready")
            result = smtp_check.main()

        self.assertEqual(result, 0)
        smtp_factory.assert_called_once_with("relay.example.com", 587, timeout=30)
        smtp_factory.return_value.send_message.assert_not_called()

    def test_test_email_defaults_to_email_recipients(self):
        environment = {
            "SMTP_HOST": "smtp.example.com",
            "SMTP_STARTTLS": "false",
            "EMAIL_FROM": "sender@example.com",
            "EMAIL_RECIPIENTS": "recipient@example.com",
        }
        with patch.dict(os.environ, environment, clear=True), \
                patch.object(smtp_check, "load_dotenv"), \
                patch.object(sys, "argv", ["smtp_check.py"]), \
                patch.object(smtp_check.socket, "getaddrinfo", return_value=[
                    (None, None, None, None, ("127.0.0.1", 587)),
                ]), \
                patch.object(smtp_check.smtplib, "SMTP") as smtp_factory, \
                redirect_stdout(StringIO()):
            smtp = smtp_factory.return_value
            smtp.ehlo.return_value = (250, b"ready")
            smtp.esmtp_features = {}
            smtp.send_message.return_value = {}
            result = smtp_check.main()

        self.assertEqual(result, 0)
        message = smtp.send_message.call_args.args[0]
        self.assertEqual(message["To"], "recipient@example.com")


if __name__ == "__main__":
    unittest.main()
