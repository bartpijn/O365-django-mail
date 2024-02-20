import os
import unittest

from django.core import mail
from django.test import TestCase
from django.test.utils import override_settings

O365_MAIL_TEST_CLIENT_ID = os.getenv('O365_MAIL_TEST_CLIENT_ID')
O365_MAIL_TEST_CLIENT_SECRET = os.getenv('O365_MAIL_TEST_CLIENT_SECRET')
O365_MAIL_TEST_TENANT_ID = os.getenv('O365_MAIL_TEST_TENANT_ID')


@unittest.skipUnless(O365_MAIL_TEST_CLIENT_ID,
                     "Set O365_MAIL_TEST_CLIENT_ID environment variable to run integration tests")
@unittest.skipUnless(O365_MAIL_TEST_CLIENT_SECRET,
                     "Set O365_MAIL_TEST_CLIENT_SECRET environment variable to run integration tests")
@unittest.skipUnless(O365_MAIL_TEST_TENANT_ID,
                     "Set O365_MAIL_TEST_TENANT_ID environment variable to run integration tests")
@override_settings(MAILJET_API_KEY=O365_MAIL_TEST_CLIENT_ID,
                   O365_MAIL_CLIENT_SECRET=O365_MAIL_TEST_CLIENT_SECRET,
                   O365_MAIL_TENANT_ID=O365_MAIL_TEST_TENANT_ID,
                   EMAIL_BACKEND="o365_django_mail.backends.O365EmailBackend")
class TestO365EmailBackend(TestCase):
    """Office 365 API integration tests."""

    def setUp(self):
        self.message = mail.EmailMultiAlternatives(
            'Subject', 'Text content', 'from@example.com', ['to@example.com'])
        self.message.attach_alternative('<p>HTML content</p>', "text/html")

    def test_send_mail(self):
        sent_count = self.message.send(fail_silently=False)

        self.assertEqual(sent_count, 1)  # noqa: F821
