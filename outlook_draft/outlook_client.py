from __future__ import annotations

import logging
import time
from typing import Any

import httpx

from . import config
from .calendar_time import outlook_timezone_prefer_header
from .errors import OutlookAPIError
from .token_manager import TokenManager

log = logging.getLogger(__name__)


class OutlookClient:
    """Sync HTTP client for Outlook and Microsoft Graph operations."""

    def __init__(
        self,
        token_manager: TokenManager,
        *,
        base_url: str = config.OUTLOOK_BASE_URL,
        default_headers: dict[str, str] | None = None,
    ):
        self._tm = token_manager
        self._base_url = base_url.rstrip("/")
        self._default_headers = default_headers or {}
        self._client: httpx.Client | None = None

    def _ensure_client(self) -> httpx.Client:
        if self._client is None or self._client.is_closed:
            self._client = httpx.Client(
                timeout=httpx.Timeout(30.0, connect=10.0),
                follow_redirects=True,
            )
        return self._client

    def close(self) -> None:
        if self._client and not self._client.is_closed:
            self._client.close()

    def _headers(self, extra_headers: dict[str, str] | None = None) -> dict[str, str]:
        headers = {
            "Authorization": f"Bearer {self._tm.token}",
            "Content-Type": "application/json",
        }
        headers.update(self._default_headers)
        if extra_headers:
            headers.update(extra_headers)
        return headers

    def _request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, str] | None = None,
        json_body: dict | None = None,
        extra_headers: dict[str, str] | None = None,
        max_retries: int = 2,
    ) -> httpx.Response:
        """Make a request to the API with retry on 401/429/5xx."""
        client = self._ensure_client()
        url = f"{self._base_url}{path}"
        headers = self._headers(extra_headers)

        for attempt in range(max_retries + 1):
            try:
                resp = client.request(
                    method, url, params=params, json=json_body, headers=headers,
                )

                if resp.status_code in (200, 201, 202, 204):
                    return resp

                if resp.status_code == 401:
                    log.warning("401 Unauthorized, reloading token")
                    self._tm.force_reload()
                    headers = self._headers(extra_headers)
                    continue

                if resp.status_code == 429 or resp.status_code >= 500:
                    retry_after = float(resp.headers.get("Retry-After", str(2 ** attempt)))
                    log.warning("HTTP %d, retrying in %.1fs", resp.status_code, retry_after)
                    time.sleep(retry_after)
                    continue

                raise OutlookAPIError(resp.status_code, resp.text[:500])

            except httpx.HTTPError as e:
                if attempt < max_retries:
                    wait = 2 ** attempt
                    log.warning("Network error: %s, retrying in %ds", e, wait)
                    time.sleep(wait)
                    continue
                raise OutlookAPIError(0, str(e)) from e

        raise OutlookAPIError(0, "Max retries exceeded")

    def list_unread(self, top: int = 20) -> list[dict[str, Any]]:
        """List unread emails, most recent first."""
        resp = self._request(
            "GET",
            "/me/mailFolders/inbox/messages",
            params={
                "$filter": "IsRead eq false",
                "$top": str(top),
                "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview",
                "$orderby": "receivedDateTime desc",
            },
        )
        return resp.json().get("value", [])

    # ── Draft operations ──────────────────────────────────────────────

    def create_draft(
        self,
        *,
        subject: str,
        body: str,
        to: list[str],
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        content_type: str = "Text",
        importance: str = "Normal",
    ) -> dict[str, Any]:
        """Create a new email draft in the Drafts folder.

        Args:
            subject: Email subject line.
            body: Email body content.
            to: List of recipient email addresses.
            cc: Optional list of CC addresses.
            bcc: Optional list of BCC addresses.
            content_type: 'Text' or 'HTML'.
            importance: 'Low', 'Normal', or 'High'.

        Returns:
            The created message object from Outlook.
        """
        def _addr(email: str) -> dict:
            return {"EmailAddress": {"Address": email.strip()}}

        payload: dict[str, Any] = {
            "Subject": subject,
            "Body": {
                "ContentType": content_type,
                "Content": body,
            },
            "ToRecipients": [_addr(a) for a in to],
            "Importance": importance,
        }
        if cc:
            payload["CcRecipients"] = [_addr(a) for a in cc]
        if bcc:
            payload["BccRecipients"] = [_addr(a) for a in bcc]

        resp = self._request("POST", "/me/messages", json_body=payload)
        return resp.json()

    def create_reply_draft(
        self,
        message_id: str,
        *,
        reply_all: bool = False,
    ) -> dict[str, Any]:
        """Create a reply draft tied to an existing message."""
        action = "createReplyAll" if reply_all else "createReply"
        resp = self._request("POST", f"/me/messages/{message_id}/{action}")
        return resp.json()

    def list_drafts(self, top: int = 20) -> list[dict[str, Any]]:
        """List messages in the Drafts folder."""
        resp = self._request(
            "GET",
            "/me/mailfolders/drafts/messages",
            params={
                "$top": str(top),
                "$select": "id,subject,toRecipients,createdDateTime,importance,bodyPreview",
                "$orderby": "createdDateTime desc",
            },
        )
        return resp.json().get("value", [])

    def get_draft(self, message_id: str) -> dict[str, Any]:
        """Get a single draft message by ID."""
        resp = self._request("GET", f"/me/messages/{message_id}")
        return resp.json()

    def delete_draft(self, message_id: str) -> None:
        """Delete a draft message."""
        self._request("DELETE", f"/me/messages/{message_id}")

    # ── Calendar ───────────────────────────────────────────────────────

    def _calendar_headers(self) -> dict[str, str]:
        return {"Prefer": outlook_timezone_prefer_header()}

    def get_agenda(self, days: int = 7, top: int = 20) -> list[dict[str, Any]]:
        """Get upcoming calendar events."""
        from datetime import datetime, timezone, timedelta
        now = datetime.now(timezone.utc)
        start = now.isoformat()
        end = (now + timedelta(days=days)).isoformat()

        resp = self._request(
            "GET",
            "/me/calendarview",
            params={
                "startDateTime": start,
                "endDateTime": end,
                "$top": str(top),
                "$select": "id,subject,start,end,location,organizer,isAllDay,isCancelled,type,seriesMasterId,recurrence",
                "$orderby": "start/dateTime",
            },
            extra_headers=self._calendar_headers(),
        )
        return resp.json().get("value", [])

    def get_event(self, event_id: str) -> dict[str, Any]:
        """Get a single calendar event by ID, including body."""
        resp = self._request(
            "GET",
            f"/me/events/{event_id}",
            params={
                "$select": "id,subject,start,end,location,organizer,isAllDay,isCancelled,body,attendees,type,seriesMasterId,recurrence",
            },
            extra_headers=self._calendar_headers(),
        )
        return resp.json()

    def create_event(
        self,
        subject: str,
        start_dt: str,
        end_dt: str,
        location: str | None = None,
        body: str | None = None,
        attendees: list[str] | None = None,
    ) -> dict[str, Any]:
        """Create a new calendar event using the configured Outlook timezone."""
        payload: dict[str, Any] = {
            "Subject": subject,
            "Start": {"DateTime": start_dt, "TimeZone": config.OUTLOOK_TIMEZONE},
            "End": {"DateTime": end_dt, "TimeZone": config.OUTLOOK_TIMEZONE},
        }
        if location:
            payload["Location"] = {"DisplayName": location}
        if body:
            payload["Body"] = {"ContentType": "Text", "Content": body}
        if attendees:
            payload["Attendees"] = [
                {"EmailAddress": {"Address": a}, "Type": "Required"}
                for a in attendees
            ]
        resp = self._request("POST", "/me/events", json_body=payload)
        return resp.json()

    def update_event(
        self,
        event_id: str,
        *,
        subject: str | None = None,
        start_dt: str | None = None,
        end_dt: str | None = None,
        location: str | None = None,
        body: str | None = None,
    ) -> dict[str, Any]:
        """Update a calendar event."""
        payload: dict[str, Any] = {}
        if subject is not None:
            payload["Subject"] = subject
        if start_dt is not None:
            payload["Start"] = {"DateTime": start_dt, "TimeZone": config.OUTLOOK_TIMEZONE}
        if end_dt is not None:
            payload["End"] = {"DateTime": end_dt, "TimeZone": config.OUTLOOK_TIMEZONE}
        if location is not None:
            payload["Location"] = {"DisplayName": location}
        if body is not None:
            payload["Body"] = {"ContentType": "Text", "Content": body}
        resp = self._request("PATCH", f"/me/events/{event_id}", json_body=payload)
        return resp.json() if resp.content else {}

    def delete_event(self, event_id: str) -> None:
        """Delete a calendar event from the user's calendar."""
        self._request("DELETE", f"/me/events/{event_id}")

    def cancel_event(self, event_id: str, *, comment: str = "") -> None:
        """Cancel a calendar event as organizer and notify attendees."""
        self._request("POST", f"/me/events/{event_id}/cancel", json_body={"Comment": comment})

    def respond_to_event(
        self,
        event_id: str,
        response: str,
        *,
        comment: str = "",
        send_response: bool = True,
    ) -> None:
        """Accept, decline, or tentatively accept a calendar event."""
        if response not in {"accept", "decline", "tentativelyAccept"}:
            raise ValueError(f"Unsupported event response: {response}")
        payload = {"Comment": comment, "SendResponse": send_response}
        self._request("POST", f"/me/events/{event_id}/{response}", json_body=payload)

    # ── Tasks ──────────────────────────────────────────────────────────

    def list_tasks(self, top: int = 20) -> list[dict[str, Any]]:
        """List incomplete tasks."""
        resp = self._request(
            "GET",
            "/me/tasks",
            params={
                "$filter": "status ne 'Completed'",
                "$top": str(top),
                "$select": "id,subject,status,dueDateTime,importance",
                "$orderby": "createdDateTime desc",
            },
        )
        return resp.json().get("value", [])

    def create_task(self, subject: str) -> dict[str, Any]:
        """Create a new task."""
        resp = self._request(
            "POST",
            "/me/tasks",
            json_body={"Subject": subject},
        )
        return resp.json()

    def complete_task(self, task_id: str) -> None:
        """Mark a task as complete."""
        self._request(
            "PATCH",
            f"/me/tasks/{task_id}",
            json_body={"Status": "Completed"},
        )

    def delete_task(self, task_id: str) -> None:
        """Delete a task."""
        self._request("DELETE", f"/me/tasks/{task_id}")

    # ── Email search / read ────────────────────────────────────────────

    def search_messages(
        self, query: str, top: int = 20
    ) -> list[dict[str, Any]]:
        """Search emails by keyword. Returns messages sorted by relevance."""
        resp = self._request(
            "GET",
            "/me/messages",
            params={
                "$search": f'"{query}"',
                "$top": str(top),
                "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview",
            },
        )
        return resp.json().get("value", [])

    def get_message(self, message_id: str) -> dict[str, Any]:
        """Get a single email message by ID, including full body."""
        resp = self._request(
            "GET",
            f"/me/messages/{message_id}",
            params={
                "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,importance,isRead",
            },
        )
        return resp.json()

    # ── Teams chats / messages ────────────────────────────────────────

    def list_teams_chats(self, top: int = 20) -> list[dict[str, Any]]:
        """List Teams chats for the current user."""
        resp = self._request(
            "GET",
            "/me/chats",
            params={
                "$top": str(top),
                "$select": "id,topic,chatType,createdDateTime,lastUpdatedDateTime,webUrl",
            },
        )
        return resp.json().get("value", [])

    def get_teams_chat(self, chat_id: str) -> dict[str, Any]:
        """Get a Teams chat by ID."""
        resp = self._request(
            "GET",
            f"/chats/{chat_id}",
            params={
                "$select": "id,topic,chatType,createdDateTime,lastUpdatedDateTime,webUrl,tenantId",
            },
        )
        return resp.json()

    def list_teams_chat_members(self, chat_id: str, top: int = 50) -> list[dict[str, Any]]:
        """List members in a Teams chat."""
        resp = self._request(
            "GET",
            f"/chats/{chat_id}/members",
        )
        members = resp.json().get("value", [])
        return members[:top]

    def list_teams_messages(self, chat_id: str, top: int = 20) -> list[dict[str, Any]]:
        """List messages for a Teams chat."""
        resp = self._request(
            "GET",
            f"/chats/{chat_id}/messages",
            params={"$top": str(top)},
            extra_headers={"Prefer": "include-unknown-enum-members"},
        )
        return resp.json().get("value", [])

    # ── People / contacts ─────────────────────────────────────────────

    def search_people(self, query: str, top: int = 10) -> list[dict[str, Any]]:
        """Search contacts and directory by name or email.

        Returns people sorted by relevance, including org directory
        users and implicit contacts from past emails.
        """
        resp = self._request(
            "GET",
            "/me/people",
            params={"$search": f'"{query}"', "$top": str(top)},
        )
        return resp.json().get("value", [])

    def update_draft(
        self,
        message_id: str,
        *,
        subject: str | None = None,
        body: str | None = None,
        content_type: str | None = None,
        to: list[str] | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        importance: str | None = None,
    ) -> dict[str, Any]:
        """Update fields on an existing draft."""
        def _addr(email: str) -> dict:
            return {"EmailAddress": {"Address": email.strip()}}

        payload: dict[str, Any] = {}
        if subject is not None:
            payload["Subject"] = subject
        if body is not None:
            ct = content_type or "Text"
            payload["Body"] = {"ContentType": ct, "Content": body}
        if to is not None:
            payload["ToRecipients"] = [_addr(a) for a in to]
        if cc is not None:
            payload["CcRecipients"] = [_addr(a) for a in cc]
        if bcc is not None:
            payload["BccRecipients"] = [_addr(a) for a in bcc]
        if importance is not None:
            payload["Importance"] = importance

        resp = self._request("PATCH", f"/me/messages/{message_id}", json_body=payload)
        return resp.json()
