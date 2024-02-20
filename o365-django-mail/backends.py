from typing import Optional

import msal
from django.conf import settings
from django.core.exceptions import ImproperlyConfigured
from django.core.mail import EmailMessage
from django.core.mail.backends.base import BaseEmailBackend
from office365.graph_client import GraphClient


class O365EmailBackend(BaseEmailBackend):
    def __init__(self, fail_silently=False, **kwargs):
        super(O365EmailBackend, self).__init__(fail_silently=fail_silently, **kwargs)

        try:
            self._client_id = settings.O365_MAIL_CLIENT_ID
            self._client_secret = settings.O365_MAIL_CLIENT_SECRET
            self._tenant_id = settings.O365_MAIL_TENANT_ID
        except:
            if not self.fail_silently:
                raise ImproperlyConfigured(
                    "Please set O365_MAIL_CLIENT_ID, O365_MAIL_CLIENT_SECRET and O365_MAIL_TENANT_ID in Django settings to use O365 mail.")

        self.client = GraphClient(self._acquire_token_msal)

    def send_messages(self, email_messages: Optional[list[EmailMessage]]) -> Optional[int]:
        if not email_messages:
            return 0

        num_sent = 0
        for message in email_messages:
            is_sent = self._send_message(message)
            if is_sent:
                num_sent += 1
            else:
                if not self.fail_silently:
                    raise Warning('Email message was not sent using O365.')

        return num_sent

    def _send_message(self, message: EmailMessage) -> bool:
        """Send email using Graph API."""
        msg = self.client.me.send_mail(
            subject=message.subject,
            body=message.body,
            to_recipientss=message.recipients(),
            bcc_recipients=message.bcc,
            cc_recipients=message.cc,
        )
        msg.execute_query()
        if msg.sent_datetime:
            return True
        return False

    @staticmethod
    def _acquire_token_msal() -> dict:
        """
        Acquire token via MSAL
        """
        authority_url = 'https://login.microsoftonline.com/{tenant_id_or_name}'
        app = msal.ConfidentialClientApplication(
            authority=authority_url,
            client_id='{client_id}',
            client_credential='{client_secret}'
        )
        token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        return token
