# ============================================================
# mailer.py
# Outlook mail bridge for Liquimech ERP Desktop App
# ============================================================

from __future__ import annotations

from pathlib import Path
from typing import Optional

try:
    import win32com.client as win32
except ImportError:
    win32 = None


class MailerService:
    """
    Simple Outlook-based mail helper.

    Notes:
    - Requires pywin32:
        pip install pywin32
    - Requires Microsoft Outlook installed and configured
    """

    def __init__(self):
        self.available = win32 is not None

    def is_available(self) -> bool:
        return self.available

    def _require_outlook(self) -> None:
        if win32 is None:
            raise RuntimeError(
                "Outlook integration is not available.\n\n"
                "Install pywin32 first:\n"
                "pip install pywin32"
            )

    def send_email_with_attachment(
        self,
        to_email: str,
        subject: str,
        body: str,
        attachment_path: str,
        cc: str = "",
        bcc: str = "",
        display_only: bool = False,
    ) -> None:
        """
        Send or display an Outlook email with one attachment.

        display_only=True will open draft in Outlook instead of sending immediately.
        """
        self._require_outlook()

        attachment = Path(attachment_path)
        if not attachment.exists():
            raise FileNotFoundError(f"Attachment not found:\n{attachment_path}")

        try:
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            mail.To = to_email
            mail.Subject = subject
            mail.Body = body

            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc

            mail.Attachments.Add(str(attachment))

            if display_only:
                mail.Display()
            else:
                mail.Send()

        except Exception as exc:
            raise RuntimeError(
                "Could not send email.\n\n"
                "Make sure Outlook is installed and configured.\n\n"
                f"Details:\n{exc}"
            ) from exc

    def create_draft_with_attachment(
        self,
        to_email: str,
        subject: str,
        body: str,
        attachment_path: str,
        cc: str = "",
        bcc: str = "",
    ) -> None:
        """
        Open an Outlook draft with the attachment instead of sending immediately.
        """
        self.send_email_with_attachment(
            to_email=to_email,
            subject=subject,
            body=body,
            attachment_path=attachment_path,
            cc=cc,
            bcc=bcc,
            display_only=True,
        )