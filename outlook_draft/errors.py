from __future__ import annotations


class OutlookDraftError(Exception):
    """Base exception."""


class CliError(OutlookDraftError):
    """A CLI failure with a message and process exit code."""

    def __init__(self, message: str, exit_code: int = 1):
        self.exit_code = exit_code
        super().__init__(message)


class TokenExpiredError(OutlookDraftError):
    """Token has expired and needs re-authentication."""


class TokenNotFoundError(OutlookDraftError):
    """No token file found."""


class OutlookAPIError(OutlookDraftError):
    """Outlook REST API returned an error."""

    def __init__(self, status: int, message: str = ""):
        self.status = status
        super().__init__(f"Outlook API {status}: {message}")
