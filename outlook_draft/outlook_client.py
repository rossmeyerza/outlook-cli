from __future__ import annotations

import logging
import time
from typing import Any

import httpx

from . import config
from .calendar_time import outlook_timezone_prefer_header
from .errors import OutlookAPIError, SendingDisabledError
from .links import encode_share_id
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
        url = path if path.startswith("https://") else f"{self._base_url}{path}"
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
                "$select": (
                    "id,subject,start,end,location,organizer,isAllDay,isCancelled,"
                    "body,attendees,type,seriesMasterId,recurrence,"
                    "isOnlineMeeting,onlineMeetingProvider,onlineMeeting"
                ),
            },
            extra_headers=self._calendar_headers(),
        )
        return resp.json()

    def find_rooms(self, room_list: str | None = None) -> list[dict[str, Any]]:
        """Find available rooms, optionally within a room list."""
        params = {"roomList": room_list} if room_list else None
        resp = self._request("GET", "/me/findRooms", params=params)
        return resp.json().get("value", [])

    def get_schedule(
        self,
        *,
        schedules: list[str],
        start_dt: str,
        end_dt: str,
        interval_minutes: int = 30,
    ) -> list[dict[str, Any]]:
        """Get free/busy availability for users or rooms."""
        payload = {
            "Schedules": schedules,
            "StartTime": {"DateTime": start_dt, "TimeZone": config.OUTLOOK_TIMEZONE},
            "EndTime": {"DateTime": end_dt, "TimeZone": config.OUTLOOK_TIMEZONE},
            "AvailabilityViewInterval": interval_minutes,
        }
        resp = self._request("POST", "/me/calendar/getSchedule", json_body=payload)
        return resp.json().get("value", [])

    def find_meeting_times(
        self,
        *,
        attendees: list[str],
        start_dt: str,
        end_dt: str,
        duration_minutes: int = 30,
        max_candidates: int = 10,
    ) -> list[dict[str, Any]]:
        """Ask Outlook to suggest meeting times for attendees."""
        payload = {
            "Attendees": [
                {"EmailAddress": {"Address": email}, "Type": "Required"}
                for email in attendees
            ],
            "TimeConstraint": {
                "Timeslots": [
                    {
                        "Start": {"DateTime": start_dt, "TimeZone": config.OUTLOOK_TIMEZONE},
                        "End": {"DateTime": end_dt, "TimeZone": config.OUTLOOK_TIMEZONE},
                    }
                ]
            },
            "MeetingDuration": f"PT{duration_minutes}M",
            "MaxCandidates": max_candidates,
        }
        resp = self._request("POST", "/me/findMeetingTimes", json_body=payload)
        return resp.json().get("MeetingTimeSuggestions", [])

    def create_event(
        self,
        subject: str,
        start_dt: str,
        end_dt: str,
        location: str | None = None,
        body: str | None = None,
        attendees: list[str] | None = None,
        is_online_meeting: bool = False,
        online_meeting_provider: str = "teamsForBusiness",
    ) -> dict[str, Any]:
        """Create a calendar event using the configured Outlook timezone."""
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
        if is_online_meeting:
            payload["IsOnlineMeeting"] = True
            payload["OnlineMeetingProvider"] = online_meeting_provider
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
        is_online_meeting: bool | None = None,
        online_meeting_provider: str = "teamsForBusiness",
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
        if is_online_meeting is not None:
            payload["IsOnlineMeeting"] = bool(is_online_meeting)
            if is_online_meeting:
                payload["OnlineMeetingProvider"] = online_meeting_provider
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

    def update_message_read_state(self, message_id: str, *, is_read: bool) -> None:
        """Mark a message read or unread."""
        self._request("PATCH", f"/me/messages/{message_id}", json_body={"IsRead": is_read})

    def send_message(self, message_id: str) -> None:
        """Send an existing draft message.

        Intentionally disabled for agent safety. Keep the endpoint knowledge here
        so it can be restored deliberately if the operator ever wants it.
        Original endpoint: POST /me/messages/{message_id}/send
        """
        raise SendingDisabledError("Email sending is intentionally disabled")

    def send_mail(
        self,
        *,
        subject: str,
        body: str,
        to: list[str],
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        content_type: str = "Text",
        save_to_sent_items: bool = True,
        importance: str = "Normal",
    ) -> None:
        """Send a new email immediately."""
        def _addr(email: str) -> dict:
            return {"EmailAddress": {"Address": email.strip()}}

        # Intentionally disabled for agent safety. Keep the payload shape here
        # so it can be restored deliberately if the operator ever wants it.
        # Original endpoint: POST /me/sendmail
        _ = (subject, body, to, cc, bcc, content_type, save_to_sent_items, importance)
        raise SendingDisabledError("Email sending is intentionally disabled")

    def list_mail_folders(self, top: int = 100) -> list[dict[str, Any]]:
        """List top-level mail folders."""
        resp = self._request(
            "GET",
            "/me/mailfolders",
            params={
                "$top": str(top),
                "$select": "id,displayName,parentFolderId,totalItemCount,unreadItemCount",
            },
        )
        return resp.json().get("value", [])

    def move_message(self, message_id: str, destination_id: str) -> dict[str, Any]:
        """Move a message to another mail folder."""
        resp = self._request(
            "POST",
            f"/me/messages/{message_id}/move",
            json_body={"DestinationId": destination_id},
        )
        return resp.json()

    def archive_message(self, message_id: str) -> dict[str, Any]:
        """Move a message to the archive folder."""
        return self.move_message(message_id, "archive")

    def get_share_drive_item(self, share_url: str) -> dict[str, Any]:
        """Resolve a SharePoint/OneDrive sharing URL to a Graph driveItem."""
        share_id = encode_share_id(share_url)
        resp = self._request("GET", f"/shares/{share_id}/driveItem")
        return resp.json()

    def download_share_url(self, share_url: str) -> bytes:
        """Download the bytes behind a SharePoint/OneDrive sharing URL.

        Uses Graph `/shares/{id}/driveItem/content`, which redirects to a
        short-lived signed download URL that does not require an SPO token.
        """
        share_id = encode_share_id(share_url)
        resp = self._request("GET", f"/shares/{share_id}/driveItem/content")
        return resp.content

    def list_message_attachments(self, message_id: str) -> list[dict[str, Any]]:
        """List message attachments."""
        resp = self._request(
            "GET",
            f"/me/messages/{message_id}/attachments",
            params={"$select": "id,name,contentType,size,isInline"},
        )
        return resp.json().get("value", [])

    def get_message_attachment(self, message_id: str, attachment_id: str) -> dict[str, Any]:
        """Get a message attachment, including contentBytes for file attachments."""
        resp = self._request("GET", f"/me/messages/{message_id}/attachments/{attachment_id}")
        return resp.json()

    # ── Teams chats / messages ────────────────────────────────────────

    def get_current_user(self) -> dict[str, Any]:
        """Get the current Graph user profile."""
        resp = self._request(
            "GET",
            "/me",
            params={"$select": "id,displayName,mail,userPrincipalName"},
        )
        return resp.json()

    def list_teams_chats(self, top: int = 20) -> list[dict[str, Any]]:
        """List Teams chats for the current user, following Graph paging."""
        chats: list[dict[str, Any]] = []
        page_size = min(max(top, 1), 50)
        path = "/me/chats"
        params: dict[str, str] | None = {
            "$top": str(page_size),
            "$select": "id,topic,chatType,createdDateTime,lastUpdatedDateTime,lastMessagePreview,webUrl",
        }

        while len(chats) < top and path:
            resp = self._request("GET", path, params=params)
            data = resp.json()
            chats.extend(data.get("value", []))
            next_link = data.get("@odata.nextLink")
            if not next_link:
                break
            path = next_link
            params = None

        return chats[:top]

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

    def list_teams_message_metadata(self, chat_id: str, top: int = 10) -> list[dict[str, Any]]:
        """List Teams message metadata for sorting.

        Graph does not allow `$select` on this endpoint in this tenant, so this
        fetches the normal message objects and callers only inspect metadata.
        """
        resp = self._request(
            "GET",
            f"/chats/{chat_id}/messages",
            params={"$top": str(top)},
            extra_headers={"Prefer": "include-unknown-enum-members"},
        )
        return resp.json().get("value", [])

    def send_teams_message(self, chat_id: str, content: str, *, content_type: str = "text") -> dict[str, Any]:
        """Send a message to a Teams chat.

        Intentionally disabled for agent safety. Teams reading remains enabled.
        Original endpoint: POST /chats/{chat_id}/messages
        """
        _ = (chat_id, content, content_type)
        raise SendingDisabledError("Teams message sending is intentionally disabled")

    def _send_teams_message_internal(self, chat_id: str, content: str, *, content_type: str = "text") -> dict[str, Any]:
        """Send a Teams message — internal use by gateway only.

        Not exposed via the CLI parser. Only the gateway command calls this.
        """
        resp = self._request(
            "POST",
            f"/chats/{chat_id}/messages",
            json_body={"body": {"contentType": content_type, "content": content}},
        )
        return resp.json()

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

    def create_contact(
        self,
        *,
        display_name: str,
        email: str,
        given_name: str | None = None,
        surname: str | None = None,
        company: str | None = None,
        mobile_phone: str | None = None,
    ) -> dict[str, Any]:
        """Create a personal Outlook contact."""
        payload: dict[str, Any] = {
            "DisplayName": display_name,
            "EmailAddresses": [{"Address": email, "Name": display_name}],
        }
        if given_name is not None:
            payload["GivenName"] = given_name
        if surname is not None:
            payload["Surname"] = surname
        if company is not None:
            payload["CompanyName"] = company
        if mobile_phone is not None:
            payload["MobilePhone1"] = mobile_phone
        resp = self._request("POST", "/me/contacts", json_body=payload)
        return resp.json()

    def update_contact(
        self,
        contact_id: str,
        *,
        display_name: str | None = None,
        email: str | None = None,
        given_name: str | None = None,
        surname: str | None = None,
        company: str | None = None,
        mobile_phone: str | None = None,
    ) -> dict[str, Any]:
        """Update a personal Outlook contact."""
        payload: dict[str, Any] = {}
        if display_name is not None:
            payload["DisplayName"] = display_name
        if email is not None:
            payload["EmailAddresses"] = [{"Address": email, "Name": display_name or email}]
        if given_name is not None:
            payload["GivenName"] = given_name
        if surname is not None:
            payload["Surname"] = surname
        if company is not None:
            payload["CompanyName"] = company
        if mobile_phone is not None:
            payload["MobilePhone1"] = mobile_phone
        resp = self._request("PATCH", f"/me/contacts/{contact_id}", json_body=payload)
        return resp.json() if resp.content else {}

    # ── Mailbox settings ──────────────────────────────────────────────

    def get_mailbox_settings(self) -> dict[str, Any]:
        """Get mailbox settings."""
        resp = self._request("GET", "/me/mailboxSettings")
        return resp.json()

    def update_mailbox_settings(
        self,
        *,
        time_zone: str | None = None,
        automatic_replies: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Update mailbox settings."""
        payload: dict[str, Any] = {}
        if time_zone is not None:
            payload["TimeZone"] = time_zone
        if automatic_replies is not None:
            payload["AutomaticRepliesSetting"] = automatic_replies
        resp = self._request("PATCH", "/me/mailboxSettings", json_body=payload)
        return resp.json() if resp.content else {}

    def update_task(
        self,
        task_id: str,
        *,
        subject: str | None = None,
        due_dt: str | None = None,
        importance: str | None = None,
        status: str | None = None,
    ) -> dict[str, Any]:
        """Update a task."""
        payload: dict[str, Any] = {}
        if subject is not None:
            payload["Subject"] = subject
        if due_dt is not None:
            payload["DueDateTime"] = {"DateTime": due_dt, "TimeZone": config.OUTLOOK_TIMEZONE}
        if importance is not None:
            payload["Importance"] = importance
        if status is not None:
            payload["Status"] = status
        resp = self._request("PATCH", f"/me/tasks/{task_id}", json_body=payload)
        return resp.json() if resp.content else {}

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
