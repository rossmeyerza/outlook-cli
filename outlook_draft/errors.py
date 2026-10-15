from __future__ import annotations


class OutlookDraftError(Exception):
    """Base exception."""


class TokenExpiredError(OutlookDraftError):
    """Token has expired and needs re-authentication."""


class TokenNotFoundError(OutlookDraftError):
    """No token file found."""


class OutlookAPIError(OutlookDraftError):
    """Outlook REST API returned an error."""

    def __init__(self, status: int, message: str = ""):
        self.status = status
        super().__init__(f"Outlook API {status}: {message}")
