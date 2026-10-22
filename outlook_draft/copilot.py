"""Microsoft 365 Copilot chat client.

Copilot has no public API that the captured tokens can use. This client uses
the same StreamHub WebSocket as the Copilot web app (copilot.cloud.microsoft):
SignalR JSON frames, a `chat_request` item per prompt, and streamed
`message_part` updates. Microsoft does not document this protocol, and it can
change without notice.
"""

from __future__ import annotations

import base64
import contextlib
import html
import json
import logging
import re
import time
import uuid
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any
from urllib.parse import quote
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from websockets.exceptions import ConnectionClosed, InvalidStatus
from websockets.sync.client import connect

from . import config
from .errors import CopilotError, TokenExpiredError, TokenNotFoundError
from .token_manager import TokenManager, _decode_claims

log = logging.getLogger(__name__)

STREAM_HUB_URL = "wss://substrate.office.com/m365Copilot/StreamHub"
ORIGIN = "https://copilot.cloud.microsoft"
USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36"
)
RECORD_SEPARATOR = "\x1e"
PING_INTERVAL_SECONDS = 15
REPLAY_QUIET_SECONDS = 4
DEFAULT_TIMEOUT_SECONDS = 240

# Query parameters that the web app sends, without per-session values.
STREAM_PARAMS = {
    "source": '"officeweb"',
    "product": "Office",
    "agentHost": "Bizchat.FullScreen",
    "licenseType": "Premium",
    "isEdu": "false",
    "agent": "work",
    "scenario": "officeweb",
}

CLIENT_INFO = {
    "clientPlatform": "ccmcopilot-web",
    "clientAppName": "Office",
    "clientEntrypoint": "ccmcopilot-officeweb",
    "ProductCategory": "Chat",
    "clientAppType": "Web",
    "productEntryPoint": "ChatPanel",
    "deviceOS": "Linux",
    "deviceType": "Desktop",
    "clientPlatformVersion": "Unknown",
}

# Copied from the web app's chat_request (October 2026).
OPTIONS_SETS = [
    "at_mention_plugins_enable", "enable_confirmation_interstitial", "enable_plugin_auth_interstitial",
    "enable_request_response_interstitials", "enable_response_action_processing", "enterprise_flux_image",
    "enterprise_flux_web", "enterprise_flux_work", "enterprise_toolbox_with_skdsstore",
    "enterprise_pagination_support", "search_result_progress_messages_with_search_queries",
    "enable_search_result_progress_messages", "flux_v3_gptv_enable_upload_multi_image_in_turn_wo_ch",
    "rich_responses", "gptvnorm2048", "enterprise_flux_work_code_interpreter", "cwc_code_interpreter_citation_fix",
    "code_interpreter_interactive_charts", "enterprise_code_interpreter_citation_fix",
    "cwc_code_interpreter_interactive_charts_inline_image", "code_interpreter_matplotlib_patching",
    "enable_batch_token_processing", "disable_cea_message_listener", "update_textdoc_response_after_streaming",
    "deepleo_networking_timeout_10minutes_canmore", "enable_selective_url_redaction", "update_memory_plugin",
    "add_custom_instructions", "agent_recommendations", "enable_gg_gpt", "enable_inferred_memory_read",
    "flux_v3_references", "flux_v3_references_entities", "flux_v3_references_ci", "add_filestore_filetype",
    "cwc_code_interpreter_citation_sourceannotations", "cwc_flux_v3", "cwc_code_interpreter",
    "cdxcwc_code_interpreter_hallucinated_url_filter", "flux_v3_image_gen_enable_dimensions",
    "flux_v3_image_gen_enable_non_watermarked_storage", "flux_v3_image_gen_enable_icon_dimensions",
    "flux_v3_image_gen_enable_system_text_with_params",
    "flux_v3_image_gen_enable_designer_dimensions_meta_prompting_in_system_prompts",
    "flux_v3_image_gen_enable_story", "bizchat_enable_federated_connectors",
    "enterprise_toolbox_with_skdsstore_search_message_extensions", "enable_ME_auth_interstitial",
    "teams_action_message_extensions_enable", "teams_search_me_annotation_card_json_enable",
    "enterprise_flux_work_gptv", "enable_image_logging_gptv_compliant", "imglogcwcgptv",
    "evaluation_data_extension",
]

ALLOWED_MESSAGE_TYPES = [
    "Chat", "Suggestion", "InternalSearchQuery", "InternalSearchResult", "Disengaged", "InternalLoaderMessage",
    "Progress", "GeneratedCode", "RenderCardRequest", "AdsQuery", "SemanticSerp", "GenerateContentQuery",
    "GenerateGraphicArt", "SearchQuery", "ConfirmationCard", "AuthError", "DeveloperLogs", "TriggerPlugin",
    "HintInvocation", "MemoryUpdate", "EndOfRequest", "Internal", "TriggerConfirmation", "ResumeInvokeAction",
    "ResumeUserInputRequest", "TriggerUserInputRequest", "EscapeHatch", "TriggerPluginAuth", "ResumePluginAuth",
    "VoiceSystemSignal", "MeteringInfo", "MeteringOutOfCredits", "SideBySide", "ReferencesListComplete",
    "CompleteExtension", "TriggerExtension", "SwitchRespondingEndpoint", "BotTurnUnlocked", "NextTurnSuggestions",
    "PromotedCapabilities", "ChatRequestReceived", "ChatName", "AgentHint",
]

# Copilot marks citations and images inline as kindref.
_INLINE_MARKER = re.compile("[^]*")
_OPEN_MARKER_TAIL = re.compile("[^]*$")
# Copilot wraps recognised entities in tags such as <Person>Adobe</Person>.
_ENTITY_TAG = re.compile(r"</?[A-Z][A-Za-z]{2,30}>")
_PARTIAL_TAG_TAIL = re.compile(r"</?[A-Za-z]*$")
_TRAILING_SPACES = re.compile(r"[ \t]+\n")
_CITED_RESULT = re.compile(r"turn\d+search(\d+)")

# Copilot finds its tools with this call. It is not an action.
TOOL_DISCOVERY_CALLS = {"api_tool_list_resources"}
# Copilot sends these when it wants the user to confirm, give input, or sign in.
# The client never confirms: it stops and reports them.
CONFIRMATION_MESSAGE_TYPES = {
    "ConfirmationCard", "TriggerConfirmation", "TriggerUserInputRequest", "TriggerPluginAuth",
}


@dataclass(frozen=True)
class CopilotFile:
    """A OneDrive or SharePoint file to give Copilot as context."""

    name: str
    item_id: str
    drive_id: str
    web_url: str
    site_id: str
    web_id: str
    list_id: str

    @property
    def spo_id(self) -> str:
        """The Substrate file ID: SPO_<base64url(siteId,webId,listId)>_<drive item ID>."""
        triple = ",".join([self.site_id, self.web_id, self.list_id])
        encoded = base64.urlsafe_b64encode(triple.encode()).decode().rstrip("=")
        return f"SPO_{encoded}_{self.item_id}"

    @property
    def extension(self) -> str:
        return self.name.rsplit(".", 1)[-1].lower() if "." in self.name else ""


@dataclass
class CopilotAnswer:
    text: str
    conversation_id: str
    request_id: str
    chat_name: str | None = None
    sources: list[dict[str, str]] = field(default_factory=list)
    complete: bool = True
    tool_calls: list[dict[str, Any]] = field(default_factory=list)
    pending_confirmations: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "text": self.text,
            "conversationId": self.conversation_id,
            "requestId": self.request_id,
            "chatName": self.chat_name,
            "sources": self.sources,
            "complete": self.complete,
            "toolCalls": self.tool_calls,
            "pendingConfirmations": self.pending_confirmations,
        }


def copilot_file_from_drive_item(item: dict[str, Any]) -> CopilotFile:
    """Build a CopilotFile from a Graph driveItem that includes `sharepointIds`."""
    if "folder" in item:
        raise CopilotError(f"{item.get('name', 'Item')} is a folder. Give Copilot a file.")
    ids = item.get("sharepointIds") or {}
    missing = [key for key in ("siteId", "webId", "listId") if not ids.get(key)]
    if missing:
        raise CopilotError(f"Graph did not return {', '.join(missing)} for {item.get('name', 'the file')}.")
    return CopilotFile(
        name=item["name"],
        item_id=item["id"],
        drive_id=(item.get("parentReference") or {}).get("driveId", ""),
        web_url=item.get("webUrl", ""),
        site_id=ids["siteId"],
        web_id=ids["webId"],
        list_id=ids["listId"],
    )


def build_file_annotation(file: CopilotFile) -> dict[str, Any]:
    """Build the message annotation that attaches a file to a prompt."""
    return {
        "messageAnnotationType": "File",
        "id": file.spo_id,
        "text": file.name,
        "messageAnnotationSource": "UserAnnotated",
        "messageAnnotationMetadata": {
            "@type": "File",
            "annotationType": "File",
            "id": file.spo_id,
            "fileName": file.name,
            "fileType": file.extension,
            "fileInfo": {"fileUrl": file.web_url, "fileStoreType": "ODSP"},
        },
    }


def build_stream_url(token: str, *, session_id: str, client_session_id: str) -> str:
    """Build the StreamHub WebSocket URL for the token's user."""
    claims = _decode_claims(token)
    oid, tid = claims.get("oid"), claims.get("tid")
    if not oid or not tid:
        raise CopilotError("The Copilot token has no user or tenant ID.")
    params = {
        "chatsessionid": session_id,
        "XRoutingParameterSessionKey": oid,
        "clientrequestid": session_id,
        "X-SessionId": client_session_id,
        **STREAM_PARAMS,
        "access_token": token,
    }
    query = "&".join(f"{key}={quote(value, safe='')}" for key, value in params.items())
    return f"{STREAM_HUB_URL}/{oid}@{tid}?{query}"


def _location_info() -> dict[str, Any]:
    try:
        zone = ZoneInfo(config.LOCAL_TIMEZONE)
    except ZoneInfoNotFoundError:
        zone = ZoneInfo("UTC")
    offset = datetime.now(zone).utcoffset()
    hours = int(offset.total_seconds() // 3600) if offset else 0
    return {"timeZoneOffset": hours, "timeZone": str(zone)}


def build_chat_request(
    prompt: str,
    *,
    conversation_id: str,
    request_id: str,
    client_session_id: str,
    start_of_session: bool,
    annotations: Iterable[dict[str, Any]] = (),
) -> dict[str, Any]:
    """Build the chatStreamRequest body for one prompt."""
    client_info = {**CLIENT_INFO, "clientSessionId": client_session_id}
    message: dict[str, Any] = {
        "author": "user",
        "inputMethod": "Keyboard",
        "text": prompt,
        "requestId": request_id,
        "locationInfo": _location_info(),
        "locale": "en-us",
        "messageType": "Chat",
        "adaptiveCards": [],
        "clientPreferences": {"executionControls": {"web": {}, "work": {}}},
        "connectedFederatedConnections": ["dummyId"],
        "clientInfo": client_info,
    }
    annotations = list(annotations)
    if annotations:
        message["messageAnnotations"] = annotations
    return {
        "source": "officeweb",
        "optionsSets": OPTIONS_SETS,
        "streamingMode": "ConciseWithPadding",
        "options": {},
        "extraExtensionParameters": {},
        "allowedMessageTypes": ALLOWED_MESSAGE_TYPES,
        "sliceIds": [],
        "threadLevelGptId": {},
        "conversationId": conversation_id,
        "requestId": request_id,
        "traceId": request_id,
        "isStartOfSession": start_of_session,
        "clientInfo": client_info,
        "message": message,
        "gpts": [{
            "id": "bizchat-as-gpt-scenario",
            "source": "BuiltInAgents",
            "clientOverrides": {
                "capabilities": [{"name": "WebSearch"}, {"name": "WorkSearch"}],
                "deepResearchModels@odata.type": "Collection(String)",
            },
        }],
        "plugins": [{"Id": "BingWebSearch", "Source": "BuiltIn"}],
        "tone": "Magic",
        "disconnectBehavior": "continue",
    }


def encode_frame(obj: dict[str, Any]) -> str:
    return json.dumps(obj, separators=(",", ":")) + RECORD_SEPARATOR


def parse_frames(raw: str | bytes) -> list[dict[str, Any]]:
    """Split one WebSocket message into SignalR JSON frames."""
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", "replace")
    frames = []
    for part in raw.split(RECORD_SEPARATOR):
        if not part.strip():
            continue
        try:
            frames.append(json.loads(part))
        except json.JSONDecodeError:
            log.debug("Skipped a frame that is not JSON: %.80s", part)
    return frames


def clean_answer_text(text: str, *, final: bool = True) -> str:
    """Remove Copilot citation markers and entity tags from answer text.

    With final=False, the text is still streaming: an open marker or a partial
    tag at the end is also removed, so it does not show before it is complete.
    """
    text = _INLINE_MARKER.sub("", text)
    text = _ENTITY_TAG.sub("", text)
    if not final:
        text = _OPEN_MARKER_TAIL.sub("", text)
        text = _PARTIAL_TAG_TAIL.sub("", text)
        return text
    return _TRAILING_SPACES.sub("\n", text).strip()


def _tool_calls_from_invocation(invocation: Any) -> list[dict[str, Any]]:
    """Parse the `invocation` field of a message into tool calls.

    The field is a JSON list of OpenAI-style function calls. Hint messages use
    a plain-text form, which is not a tool call and is ignored.
    """
    if not isinstance(invocation, str):
        return []
    try:
        calls = json.loads(invocation)
    except json.JSONDecodeError:
        return []
    result = []
    for call in calls if isinstance(calls, list) else []:
        function = (call or {}).get("function") or {}
        name = function.get("name")
        if not name or name in TOOL_DISCOVERY_CALLS:
            continue
        arguments: Any = function.get("arguments")
        with contextlib.suppress(json.JSONDecodeError):
            arguments = json.loads(arguments) if isinstance(arguments, str) else arguments
        result.append({"id": call.get("id"), "tool": name, "arguments": arguments})
    return result


def _search_results_from_citation_message(text: str) -> dict[str, dict[str, str]]:
    """Map search result numbers to sources. Results without a name, such as thumbnails, are skipped."""
    try:
        data = json.loads(text)
    except (json.JSONDecodeError, TypeError):
        return {}
    results = (data or {}).get("searchResultsMap") or {}
    sources = {}
    for key, result in results.items():
        result = result or {}
        if result.get("sourceName") and result.get("sourceUrl"):
            sources[str(key)] = {"title": html.unescape(result["sourceName"]), "url": result["sourceUrl"]}
    return sources


def cited_result_numbers(text: str) -> list[str]:
    """Return the search result numbers that the answer cites, in order of first citation.

    A citation marker such as `turn1search3L2` points to entry "3" of the
    searchResultsMap of the same turn.
    """
    numbers: list[str] = []
    for marker in _INLINE_MARKER.findall(text):
        for number in _CITED_RESULT.findall(marker):
            if number not in numbers:
                numbers.append(number)
    return numbers


class AnswerCollector:
    """Collect one Copilot turn from StreamHub frames."""

    def __init__(self, request_id: str):
        self.request_id = request_id
        self.done = False
        self.chat_name: str | None = None
        self.progress: str | None = None
        self._search_results: dict[str, dict[str, str]] = {}
        self.errors: list[str] = []
        self.tool_calls: list[dict[str, Any]] = []
        self.pending_confirmations: list[dict[str, Any]] = []
        self._raw_text = ""
        self._emitted = ""
        self._seen_request = False

    @property
    def text(self) -> str:
        return clean_answer_text(self._raw_text)

    @property
    def sources(self) -> list[dict[str, str]]:
        """The search results that the answer cites. Copilot also returns results that it does not cite."""
        return [self._search_results[n] for n in cited_result_numbers(self._raw_text) if n in self._search_results]

    def feed(self, frame: dict[str, Any]) -> str:
        """Process one frame. Return the new answer text to show, if any."""
        if frame.get("type") == 7:
            self.errors.append(frame.get("error") or "Copilot closed the connection.")
            self.done = True
            return ""
        if frame.get("type") == 3 and frame.get("error"):
            self.errors.append(str(frame["error"]))
            return ""
        item = frame.get("item") or {}
        if item.get("type") != "message_part":
            return ""
        argument = item.get("argument") or {}
        for message in argument.get("messages") or []:
            self._feed_message(message)
        return self._delta()

    def _feed_message(self, message: dict[str, Any]) -> None:
        request_id = message.get("requestId")
        if request_id == self.request_id:
            self._seen_request = True
        elif request_id:
            return  # A message from an earlier turn of the conversation.
        if not self._seen_request:
            return
        message_type = message.get("messageType")
        text = message.get("text") or ""
        if message_type != "HintInvocation":
            self._record_tool_calls(message.get("invocation"))
        if message_type in CONFIRMATION_MESSAGE_TYPES:
            # Copilot waits for the user here. Stop, and do not confirm.
            self.pending_confirmations.append({"type": message_type, "text": text})
            self.done = True
        elif message_type == "EndOfRequest":
            self.done = True
        elif message_type == "ChatName":
            self.chat_name = text or self.chat_name
        elif message_type == "Progress":
            self.progress = text or self.progress
        elif message_type in ("AuthError", "Disengaged"):
            self.errors.append(text or message_type)
        elif message_type == "Internal" and "searchResultsMap" in text:
            self._search_results = _search_results_from_citation_message(text) or self._search_results
        elif message_type is None and message.get("author") == "bot" and text:
            self._raw_text = text

    def _record_tool_calls(self, invocation: Any) -> None:
        known = {call["id"] for call in self.tool_calls if call.get("id")}
        for call in _tool_calls_from_invocation(invocation):
            if call["id"] and call["id"] in known:
                continue
            self.tool_calls.append(call)

    def _delta(self) -> str:
        visible = clean_answer_text(self._raw_text, final=False)
        if not visible.startswith(self._emitted):
            return ""
        delta = visible[len(self._emitted):]
        self._emitted = visible
        return delta


class CopilotClient:
    """Send prompts to Microsoft 365 Copilot over the StreamHub WebSocket."""

    def __init__(self, token_manager: TokenManager, *, connector: Callable[..., Any] = connect):
        self._tm = token_manager
        self._connect = connector

    def ask(
        self,
        prompt: str,
        *,
        files: Iterable[CopilotFile] = (),
        conversation_id: str | None = None,
        timeout: float = DEFAULT_TIMEOUT_SECONDS,
        on_text: Callable[[str], None] | None = None,
        on_progress: Callable[[str], None] | None = None,
    ) -> CopilotAnswer:
        """Send one prompt and wait for the full answer.

        With conversation_id, the prompt continues that conversation. Copilot
        keeps the files of earlier turns, so a follow-up does not need them again.
        """
        files = list(files)
        try:
            return self._ask_once(prompt, files, conversation_id, timeout, on_text, on_progress)
        except InvalidStatus as exc:
            if exc.response.status_code not in (401, 403):
                raise CopilotError(f"Copilot rejected the connection: HTTP {exc.response.status_code}") from exc
            log.info("Copilot returned HTTP %s, refreshing the token", exc.response.status_code)
            if not self._tm.run_reauth(headless=True):
                raise CopilotError("Copilot authentication failed. Run 'outlook-cli auth' to retry.") from exc
        try:
            return self._ask_once(prompt, files, conversation_id, timeout, on_text, on_progress)
        except InvalidStatus as exc:
            raise CopilotError(f"Copilot rejected the connection: HTTP {exc.response.status_code}") from exc

    def _token(self) -> str:
        try:
            return self._tm.get_token(auto_reauth=True)
        except (TokenExpiredError, TokenNotFoundError) as exc:
            raise CopilotError(f"{exc}. Run 'outlook-cli auth' to retry.") from exc

    def _ask_once(
        self,
        prompt: str,
        files: list[CopilotFile],
        conversation_id: str | None,
        timeout: float,
        on_text: Callable[[str], None] | None,
        on_progress: Callable[[str], None] | None,
    ) -> CopilotAnswer:
        new_conversation = conversation_id is None
        conversation_id = conversation_id or str(uuid.uuid4())
        request_id = uuid.uuid4().hex
        client_session_id = str(uuid.uuid4())
        url = build_stream_url(self._token(), session_id=str(uuid.uuid4()), client_session_id=client_session_id)
        topic = f"/conversations/{conversation_id}"
        request = build_chat_request(
            prompt,
            conversation_id=conversation_id,
            request_id=request_id,
            client_session_id=client_session_id,
            start_of_session=new_conversation,
            annotations=[build_file_annotation(f) for f in files],
        )
        collector = AnswerCollector(request_id)

        with self._connect(url, origin=ORIGIN, user_agent_header=USER_AGENT, max_size=None, open_timeout=30) as ws:
            ws.send(encode_frame({"protocol": "json", "version": 1}))
            ws.recv(timeout=30)
            ws.send(encode_frame({"type": 6}))
            ws.send(encode_frame({"arguments": [], "invocationId": "1", "streamIds": ["0"], "target": "stream", "type": 4}))
            subscribe: dict[str, Any] = {"type": "subscribe", "topicId": topic}
            if new_conversation:
                subscribe["action"] = "create"
            ws.send(encode_frame({"invocationId": "0", "item": subscribe, "type": 2}))
            if not new_conversation:
                _drain_replay(ws)
            ws.send(encode_frame({
                "invocationId": "0",
                "item": {
                    "type": "chat_request",
                    "chatStreamRequest": request,
                    "idKey": str(uuid.uuid4()),
                    "noteTimestamp": int(time.time() * 1000),
                    "topicId": topic,
                },
                "type": 2,
            }))
            self._read_answer(ws, collector, timeout, on_text, on_progress)

        if collector.errors and not collector.text and not collector.pending_confirmations:
            raise CopilotError("; ".join(collector.errors))
        return CopilotAnswer(
            text=collector.text,
            conversation_id=conversation_id,
            request_id=request_id,
            chat_name=collector.chat_name,
            sources=collector.sources,
            complete=collector.done and not collector.errors and not collector.pending_confirmations,
            tool_calls=collector.tool_calls,
            pending_confirmations=collector.pending_confirmations,
        )

    @staticmethod
    def _read_answer(
        ws: Any,
        collector: AnswerCollector,
        timeout: float,
        on_text: Callable[[str], None] | None,
        on_progress: Callable[[str], None] | None,
    ) -> None:
        deadline = time.monotonic() + timeout
        last_ping = time.monotonic()
        last_progress = None
        while not collector.done and time.monotonic() < deadline:
            if time.monotonic() - last_ping > PING_INTERVAL_SECONDS:
                ws.send(encode_frame({"type": 6}))
                last_ping = time.monotonic()
            try:
                raw = ws.recv(timeout=5)
            except TimeoutError:
                continue
            except ConnectionClosed as exc:
                collector.errors.append(f"Copilot closed the connection ({exc}).")
                return
            for frame in parse_frames(raw):
                delta = collector.feed(frame)
                if delta and on_text:
                    on_text(delta)
            if on_progress and collector.progress and collector.progress != last_progress:
                last_progress = collector.progress
                on_progress(last_progress)


def _drain_replay(ws: Any) -> None:
    """Read the replay of the previous turn that follows a subscribe.

    A chat_request sent during the replay is lost, so wait until the replay
    ends or the connection is quiet.
    """
    quiet_until = time.monotonic() + REPLAY_QUIET_SECONDS
    hard_stop = time.monotonic() + 30
    while time.monotonic() < min(quiet_until, hard_stop):
        try:
            raw = ws.recv(timeout=1)
        except TimeoutError:
            continue
        if any(
            message.get("messageType") == "EndOfRequest"
            for frame in parse_frames(raw)
            for message in ((frame.get("item") or {}).get("argument") or {}).get("messages") or []
        ):
            return
        quiet_until = time.monotonic() + REPLAY_QUIET_SECONDS
