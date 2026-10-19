# SPDX-FileCopyrightText: 2025 INDUSTRIA DE DISEÑO TEXTIL, S.A. (INDITEX, S.A.)
# SPDX-FileCopyrightText: 2026 Jason Hastain (modifications)
# SPDX-License-Identifier: Apache-2.0
import json
import logging
import re
from urllib.parse import parse_qs, unquote, urlparse

from botbuilder.core import BotAdapter, TurnContext
from botbuilder.core.teams import TeamsInfo
from botbuilder.integration.aiohttp import CloudAdapter
from botbuilder.schema import (
    Activity,
    ActivityTypes,
    ChannelAccount,
    ConversationAccount,
    ConversationReference,
    Mention,
    TextFormatTypes,
)
from botbuilder.schema.teams import TeamsChannelAccount
from botframework.connector.aio.operations_async import ConversationsOperations
from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.models.aad_user_conversation_member import (
    AadUserConversationMember,
)
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.channel import Channel
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.chat_message_mention import ChatMessageMention
from msgraph.generated.models.chat_message_mentioned_identity_set import (
    ChatMessageMentionedIdentitySet,
)
from msgraph.generated.models.identity import Identity
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.teams.item.channels.item.messages.item.chat_message_item_request_builder import (
    ChatMessageItemRequestBuilder,
)
from msgraph.generated.teams.item.channels.item.messages.item.replies.replies_request_builder import (
    RepliesRequestBuilder,
)
from msgraph.generated.teams.item.channels.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)
from msgraph.graph_service_client import GraphServiceClient
from pydantic import BaseModel, Field

LOGGER = logging.getLogger(__name__)


class TeamsThread(BaseModel):
    thread_id: str = Field(
        description="Thread ID as a string in the format '1743086901347'"
    )
    title: str = Field(description="Message title")
    content: str = Field(description="Message content")


class ExtractedAttachment(BaseModel):
    type: str = Field(description="Attachment type, e.g. 'adaptive_card'")
    extracted_text: str = Field(description="Human-readable text extracted from the attachment")


class TeamsMessage(BaseModel):
    thread_id: str = Field(
        description="Thread ID as a string in the format '1743086901347'"
    )
    message_id: str = Field(description="Message ID")
    content: str = Field(description="Message content")
    attachments: list[ExtractedAttachment] = Field(
        default_factory=list,
        description="Extracted content from adaptive card attachments (e.g. bot alerts)",
    )


class TeamsMember(BaseModel):
    name: str = Field(
        description="Member name used in mentions and user information cards"
    )
    email: str = Field(description="Member email")


class PagedTeamsMessages(BaseModel):
    cursor: str | None = Field(
        description="Cursor to retrieve the next page of messages."
    )
    limit: int = Field(description="Page limit, maximum number of items to retrieve")
    total: int = Field(description="Total items available for retrieval")
    items: list[TeamsMessage] = Field(description="List of channel messages or threads")


class TeamsTeam(BaseModel):
    team_id: str = Field(description="The unique identifier of the team")
    display_name: str = Field(description="The display name of the team")
    description: str | None = Field(
        default=None, description="The description of the team"
    )


class TeamsChannel(BaseModel):
    channel_id: str = Field(description="The unique identifier of the channel")
    display_name: str = Field(description="The display name of the channel")
    description: str | None = Field(
        default=None, description="The description of the channel"
    )
    team_id: str = Field(description="The team ID this channel belongs to")


class TimeSlot(BaseModel):
    start: str = Field(description="Start time in ISO 8601 format")
    end: str = Field(description="End time in ISO 8601 format")


class AvailableTimesResult(BaseModel):
    date: str = Field(description="The date queried (YYYY-MM-DD)")
    slots: list[TimeSlot] = Field(description="Available time slots common to all attendees")
    attendees: list[str] = Field(description="Email addresses queried")
    duration_minutes: int = Field(description="Requested meeting duration in minutes")


class MeetingResult(BaseModel):
    event_id: str = Field(description="The created calendar event ID")
    subject: str = Field(description="Meeting subject")
    start: str = Field(description="Meeting start time")
    end: str = Field(description="Meeting end time")
    join_url: str | None = Field(default=None, description="Teams meeting join URL")
    attendees: list[str] = Field(description="Attendee email addresses")
    is_recurring: bool = Field(default=False, description="Whether this is a recurring meeting")


class ResolvedUser(BaseModel):
    display_name: str = Field(description="User's display name")
    email: str = Field(description="User's email address (mail or UPN)")
    user_id: str = Field(default="", description="Azure AD user ID if available")
    department: str = Field(default="", description="User's department")
    job_title: str = Field(default="", description="User's job title")
    relevance_score: float = Field(
        default=0.0,
        description="Relevance score from People API (higher = more relevant to caller)",
    )


class ResolvedUserList(BaseModel):
    users: list[ResolvedUser] = Field(description="List of matched users")
    query: str = Field(description="The original search query")
    total_matches: int = Field(description="Number of matches found")


class DirectMessageResult(BaseModel):
    chat_id: str = Field(description="The 1:1 chat ID")
    message_id: str = Field(default="", description="The sent message ID")
    recipient: str = Field(description="Recipient email or name")
    content_preview: str = Field(default="", description="Preview of sent content")


class CalendarEvent(BaseModel):
    event_id: str = Field(description="The unique event ID")
    subject: str = Field(description="Event subject/title")
    start: str = Field(description="Start time in ISO 8601 format")
    end: str = Field(description="End time in ISO 8601 format")
    timezone: str = Field(default="", description="Timezone of the event")
    organizer: str = Field(default="", description="Organizer email address")
    attendees: list[str] = Field(default_factory=list, description="Attendee email addresses")
    body_preview: str = Field(default="", description="Preview of the event body text")
    is_online_meeting: bool = Field(default=False, description="Whether this is an online meeting")
    join_url: str | None = Field(default=None, description="Teams meeting join URL if online")
    is_recurring: bool = Field(default=False, description="Whether this is a recurring event")
    series_master_id: str | None = Field(default=None, description="Series master event ID for recurring events")
    recurrence_pattern: str | None = Field(default=None, description="Human-readable recurrence description")
    web_link: str | None = Field(default=None, description="URL to open the event in Outlook Web")


class CalendarEventList(BaseModel):
    events: list[CalendarEvent] = Field(description="List of calendar events")
    total_returned: int = Field(description="Number of events returned")


class ChatInfo(BaseModel):
    chat_id: str = Field(description="The unique chat ID")
    chat_type: str = Field(description="Chat type: oneOnOne, group, or meeting")
    topic: str = Field(default="", description="Chat topic/name (group chats only)")
    last_updated: str = Field(default="", description="When the chat was last updated (ISO 8601)")
    last_message_preview: str = Field(default="", description="Preview of the last message")
    members: list[str] = Field(default_factory=list, description="Display names of chat members")


class ChatList(BaseModel):
    chats: list[ChatInfo] = Field(description="List of chats")
    total: int = Field(description="Number of chats returned")


class ChatMessageInfo(BaseModel):
    message_id: str = Field(description="The unique message ID")
    sender: str = Field(default="", description="Display name of the sender")
    content: str = Field(description="Message content (plain text)")
    created: str = Field(default="", description="When the message was sent (ISO 8601)")
    attachments: list[ExtractedAttachment] = Field(
        default_factory=list,
        description="Extracted content from adaptive card attachments",
    )


class ChatMessageList(BaseModel):
    chat_id: str = Field(description="The chat ID")
    messages: list[ChatMessageInfo] = Field(description="List of messages")
    total: int = Field(description="Number of messages returned")
    next_cursor: str | None = Field(default=None, description="Cursor for next page of messages")


class TranscriptInfo(BaseModel):
    transcript_id: str = Field(description="The unique transcript ID (base64 string)")
    meeting_id: str = Field(description="The online meeting ID this transcript belongs to")
    created: str = Field(default="", description="When the transcript was created (ISO 8601)")
    end: str = Field(default="", description="When the transcript ended (ISO 8601)")
    organizer: str = Field(default="", description="Display name of the meeting organizer")


class TranscriptList(BaseModel):
    transcripts: list[TranscriptInfo] = Field(description="List of transcripts for the meeting")
    meeting_id: str = Field(description="The online meeting ID queried")
    total: int = Field(description="Number of transcripts found")


class TranscriptContent(BaseModel):
    transcript_id: str = Field(description="The transcript ID")
    meeting_id: str = Field(description="The online meeting ID")
    content: str = Field(description="The transcript text in VTT format with speaker names and timestamps")


def _extract_adaptive_card_text(card_body: list) -> str:
    """Recursively extract human-readable text from adaptive card body elements.

    Handles TextBlock, FactSet, Column, ColumnSet, Container, RichTextBlock,
    ImageSet, ActionSet, and nested structures.
    """
    lines: list[str] = []

    for element in card_body:
        if not isinstance(element, dict):
            continue

        elem_type = element.get("type", "")

        if elem_type == "TextBlock":
            text = element.get("text", "")
            if text:
                lines.append(text)

        elif elem_type == "FactSet":
            for fact in element.get("facts", []):
                title = fact.get("title", "")
                value = fact.get("value", "")
                if title or value:
                    lines.append(f"{title}: {value}" if title else value)

        elif elem_type == "RichTextBlock":
            for inline in element.get("inlines", []):
                if isinstance(inline, dict):
                    text = inline.get("text", "")
                    if text:
                        lines.append(text)
                elif isinstance(inline, str):
                    lines.append(inline)

        elif elem_type == "Image":
            alt = element.get("altText", "")
            if alt:
                lines.append(f"[Image: {alt}]")

        elif elem_type == "ActionSet":
            for action in element.get("actions", []):
                title = action.get("title", "")
                url = action.get("url", "")
                if title and url:
                    lines.append(f"[{title}]({url})")
                elif title:
                    lines.append(f"[{title}]")

        elif elem_type in ("Container", "TableCell"):
            nested = element.get("items", [])
            if nested:
                lines.append(_extract_adaptive_card_text(nested))

        elif elem_type == "ColumnSet":
            for col in element.get("columns", []):
                nested = col.get("items", [])
                if nested:
                    lines.append(_extract_adaptive_card_text(nested))

        elif elem_type == "Table":
            for row in element.get("rows", []):
                cells = row.get("cells", [])
                for cell in cells:
                    nested = cell.get("items", [])
                    if nested:
                        lines.append(_extract_adaptive_card_text(nested))

        # Recurse into body if present (nested cards / fallback)
        if "body" in element and isinstance(element["body"], list):
            lines.append(_extract_adaptive_card_text(element["body"]))

    return "\n".join(line for line in lines if line)


def _extract_o365_connector_card_text(card: dict) -> str:
    """Extract human-readable text from an O365 Connector card.

    Handles summary, sections (activityTitle, activitySubtitle, text, facts),
    and potentialAction links.  All HTML is stripped from extracted fields.
    """
    lines: list[str] = []

    summary = card.get("summary", "")
    if summary:
        lines.append(_strip_html(summary))

    for section in card.get("sections") or []:
        if not isinstance(section, dict):
            continue

        for field in ("activityTitle", "activitySubtitle", "text"):
            value = section.get(field, "")
            if value:
                lines.append(_strip_html(value))

        for fact in section.get("facts") or []:
            if not isinstance(fact, dict):
                continue
            name = fact.get("name", "")
            value = _strip_html(fact.get("value", ""))
            if name or value:
                lines.append(f"{name}: {value}" if name else value)

    for action in card.get("potentialAction") or []:
        if not isinstance(action, dict):
            continue
        name = action.get("name", "")
        targets = action.get("targets") or []
        uri = ""
        for t in targets:
            if isinstance(t, dict):
                uri = t.get("uri", "")
                if uri:
                    break
        if name and uri:
            lines.append(f"[{name}]({uri})")
        elif name:
            lines.append(f"[{name}]")

    return "\n".join(line for line in lines if line)


def _extract_attachments(attachments) -> list[ExtractedAttachment]:
    """Process ChatMessage attachments and extract readable text from adaptive cards
    and O365 Connector cards."""
    if not attachments:
        return []

    results: list[ExtractedAttachment] = []
    for att in attachments:
        content_type = getattr(att, "content_type", None) or ""

        raw_content = getattr(att, "content", None) or ""
        if not raw_content:
            continue

        try:
            card = json.loads(raw_content)
        except (json.JSONDecodeError, TypeError):
            continue

        if content_type == "application/vnd.microsoft.card.adaptive":
            body = card.get("body", [])
            if not isinstance(body, list):
                continue
            text = _extract_adaptive_card_text(body)
            if text:
                results.append(ExtractedAttachment(
                    type="adaptive_card",
                    extracted_text=text,
                ))

        elif content_type == "application/vnd.microsoft.teams.card.o365connector":
            text = _extract_o365_connector_card_text(card)
            if text:
                results.append(ExtractedAttachment(
                    type="o365_connector_card",
                    extracted_text=text,
                ))

    return results


def _channel_membership_type(channel_type: str):
    """Map the string channel type callers pass to the Graph enum.

    Graph types this field as ChannelMembershipType, not str; passing the raw
    string works at runtime but is a type error and silently accepts typos.
    """
    from msgraph.generated.models.channel_membership_type import (
        ChannelMembershipType,
    )

    mapping = {
        "standard": ChannelMembershipType.Standard,
        "private": ChannelMembershipType.Private,
        "shared": ChannelMembershipType.Shared,
    }
    try:
        return mapping[channel_type.strip().lower()]
    except KeyError:
        raise ValueError(
            f"channel_type must be one of {sorted(mapping)}, got {channel_type!r}"
        ) from None


def _strip_html(html: str) -> str:
    """Strip HTML tags and decode common entities."""
    text = re.sub(r"<[^>]+>", "", html)
    text = (
        text.replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
    )
    return text.strip()


# Matches a single markdown separator cell, e.g. "---", ":--", "--:", ":-:".
_MD_SEPARATOR_CELL = re.compile(r"^:?-{1,}:?$")


def _is_md_table_separator(line: str) -> bool:
    """Return True if the line is a markdown table separator row.

    Accepts rows like ``| --- | :--: |``, ``---|---`` and single-column
    ``| --- |``. Requires at least one pipe so a lone ``---`` (a horizontal
    rule) is not treated as a table separator.
    """
    if "|" not in line:
        return False
    cells = _split_md_row(line)
    cells = [c for c in cells if c != ""]
    return bool(cells) and all(_MD_SEPARATOR_CELL.match(c) for c in cells)


def _split_md_row(line: str) -> list[str]:
    """Split a markdown table row into trimmed cell values.

    Handles optional leading/trailing pipes and escaped pipes (``\\|``).
    """
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    # Split on unescaped pipes, then restore escaped ones.
    cells = re.split(r"(?<!\\)\|", line)
    return [c.replace("\\|", "|").strip() for c in cells]


def _escape_cell(text: str) -> str:
    """HTML-escape a markdown table cell's text content."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def contains_markdown_table(content: str) -> bool:
    """Return True if the content contains a GitHub-style pipe table."""
    lines = content.splitlines()
    for i in range(1, len(lines)):
        if _is_md_table_separator(lines[i]) and "|" in lines[i - 1]:
            return True
    return False


def convert_markdown_tables_to_html(content: str) -> str:
    """Convert GitHub-style markdown pipe tables to HTML ``<table>`` markup.

    Teams' Graph HTML message body does not interpret markdown, so a pipe
    table is rendered as literal ``| --- |`` text. This finds contiguous
    pipe-table blocks (a header row, a separator row of dashes, and zero or
    more data rows) and rewrites them as bordered HTML tables that Teams
    renders as a real grid. Content outside table blocks is returned
    unchanged, so pre-existing HTML is left intact.
    """
    lines = content.splitlines()
    out: list[str] = []
    i = 0
    n = len(lines)
    while i < n:
        # A table is a header row immediately followed by a separator row.
        if (
            i + 1 < n
            and "|" in lines[i]
            and _is_md_table_separator(lines[i + 1])
        ):
            header = _split_md_row(lines[i])
            i += 2  # skip header + separator
            rows: list[list[str]] = []
            while i < n and "|" in lines[i] and lines[i].strip():
                rows.append(_split_md_row(lines[i]))
                i += 1
            out.append(_render_html_table(header, rows))
        else:
            out.append(lines[i])
            i += 1
    return "\n".join(out)


def _render_html_table(header: list[str], rows: list[list[str]]) -> str:
    """Render a header + rows into a bordered HTML table string."""
    thead_cells = "".join(
        f'<th align="left">{_escape_cell(c)}</th>' for c in header
    )
    body_rows = []
    for row in rows:
        # Pad/truncate to header width so the grid stays rectangular.
        cells = (row + [""] * len(header))[: len(header)]
        body_rows.append(
            "<tr>"
            + "".join(f"<td>{_escape_cell(c)}</td>" for c in cells)
            + "</tr>"
        )
    return (
        '<table border="1" cellpadding="6" cellspacing="0" '
        'style="border-collapse:collapse;border:1px solid #ccc;">'
        f'<thead><tr style="background-color:#f2f2f2;">{thead_cells}</tr></thead>'
        f'<tbody>{"".join(body_rows)}</tbody>'
        "</table>"
    )


def convert_markdown_inline_to_html(content: str) -> str:
    """Convert common markdown inline/block syntax to Teams-safe HTML.

    Teams' Graph HTML body ignores markdown and collapses bare newlines, so
    plain text sent as HTML arrives as one unbroken paragraph. This converts
    the subset of markdown that shows up in practice -- bold, inline code,
    links, bullets, and blank-line paragraph breaks -- and turns remaining
    newlines into ``<br>``.

    Lines already containing HTML markup (notably the ``<table>`` blocks
    emitted by :func:`convert_markdown_tables_to_html`) are passed through
    untouched so this can run after the table conversion.
    """
    blocks = re.split(r"\n\s*\n", content)
    html_blocks: list[str] = []

    for block in blocks:
        stripped = block.strip()
        if not stripped:
            continue

        # Leave pre-rendered HTML (e.g. converted tables) exactly as-is.
        if stripped.startswith("<table") or stripped.startswith("<hr"):
            html_blocks.append(stripped)
            continue

        # Within a block, consecutive bullet lines become a <ul>; runs of
        # non-bullet lines become a <p> with <br> between them. This keeps a
        # lead-in sentence and its bullet list rendering correctly.
        lines = [ln for ln in stripped.splitlines() if ln.strip()]
        run: list[str] = []
        run_is_list = False

        def _flush() -> None:
            if not run:
                return
            if run_is_list:
                items = "".join("<li>" + _inline_md(x) + "</li>" for x in run)
                html_blocks.append("<ul>" + items + "</ul>")
            else:
                html_blocks.append(
                    "<p>" + "<br>".join(_inline_md(x) for x in run) + "</p>"
                )
            run.clear()

        for ln in lines:
            is_bullet = bool(re.match(r"^\s*[-*]\s+", ln))
            if is_bullet != run_is_list:
                _flush()
                run_is_list = is_bullet
            run.append(re.sub(r"^\s*[-*]\s+", "", ln) if is_bullet else ln)
        _flush()

    return "".join(html_blocks)


def _inline_md(text: str) -> str:
    """Apply inline markdown (bold, code, links) to a single line."""
    # Bare URLs are linkified only when not already inside a markdown link.
    text = re.sub(r"\[([^\]]+)\]\((https?://[^\s)]+)\)", r'<a href="\2">\1</a>', text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"`([^`]+)`", r"<code>\1</code>", text)
    text = re.sub(
        r'(?<!href=")(?<!>)(https?://[^\s<>"]+)(?![^<]*</a>)',
        r'<a href="\1">\1</a>',
        text,
    )
    return text


class TeamsClient:
    def __init__(
        self,
        adapter: CloudAdapter | None,
        graph_client: GraphServiceClient,
        teams_app_id: str,
        team_id: str = "",
        teams_channel_id: str = "",
        user_display_name: str | None = None,
        is_delegated: bool = False,
    ):
        self.adapter = adapter
        self.graph_client = graph_client
        self.teams_app_id = teams_app_id
        self.team_id = team_id
        self.teams_channel_id = teams_channel_id
        self.service_url = None
        self._user_display_name = user_display_name
        self.is_delegated = is_delegated
        if self.adapter is not None:
            self.adapter.on_turn_error = self.on_turn_error

    def _attribution_footer(self, html: bool = True) -> str:
        """Return an attribution line for messages posted via MCP.

        In delegated mode the signed-in user's display name is included.
        In app mode a generic label is used.
        """
        if self._user_display_name:
            label = (
                f"Posted via Teams MCP on behalf of "
                f"{self._user_display_name}"
            )
        else:
            label = "Posted via Teams MCP"
        if html:
            return (
                '<hr style="margin-top:8px">'
                f'<span style="font-size:10px;color:#888">'
                f"{label}</span>"
            )
        return f"\n\n---\n*{label}*"

    def _prepare_message_body(
        self, content: str, content_type: str
    ) -> tuple[str, BodyType]:
        """Normalize outbound message content and pick the Graph body type.

        Markdown pipe tables are converted to HTML ``<table>`` markup so
        Teams renders a real grid instead of literal ``| --- |`` text. If a
        "text" message contains a table, it is upgraded to HTML so the
        converted markup is honored. The attribution footer is appended in
        the matching format. Returns the final content and its BodyType.
        """
        effective_type = content_type
        body = content
        if content_type != "html" and contains_markdown_table(content):
            effective_type = "html"

        if effective_type == "html":
            body = convert_markdown_tables_to_html(body)
            body = convert_markdown_inline_to_html(body)
            body += self._attribution_footer(html=True)
            return body, BodyType.Html

        body += self._attribution_footer(html=False)
        return body, BodyType.Text

    def get_team_id(self):
        return self.team_id

    @staticmethod
    async def on_turn_error(context: TurnContext, error: Exception):
        LOGGER.error(f"Error {str(error)}")
        # await context.send_activity("An error occurred in the bot, please try again later")

    def _create_conversation_reference(
        self, channel_id: str | None = None
    ) -> ConversationReference:
        resolved_channel_id = channel_id or self.teams_channel_id
        service_url = "https://smba.trafficmanager.net/emea/"
        if self.service_url is not None:
            service_url = self.service_url
        return ConversationReference(
            bot=TeamsChannelAccount(id=self.teams_app_id, name="MCP Bot"),
            channel_id=resolved_channel_id,
            service_url=service_url,
            conversation=ConversationAccount(
                id=resolved_channel_id,
                is_group=True,
                conversation_type="channel",
                name="Teams channel",
            ),
        )

    async def _initialize(self):
        if self.adapter is None:
            return
        if not self.service_url:

            def context_callback(context: TurnContext):
                self.service_url = context.activity.service_url

            await self.adapter.continue_conversation(
                bot_app_id=self.teams_app_id,
                reference=self._create_conversation_reference(),
                callback=context_callback,
            )

    async def list_teams(self) -> list[TeamsTeam]:
        """List all teams the user/app can access.

        In delegated mode, uses /me/joinedTeams to return teams the
        signed-in user is a member of (with pagination).
        In app mode, uses /teams (application-level endpoint).

        Returns:
            List of teams with their IDs and display names
        """
        try:
            result = []
            if self.is_delegated:
                response = await self.graph_client.me.joined_teams.get()
            else:
                response = await self.graph_client.teams.get()
            while response and response.value:
                for team in response.value:
                    result.append(
                        TeamsTeam(
                            team_id=team.id or "",
                            display_name=team.display_name or "",
                            description=team.description,
                        )
                    )
                # Follow pagination if there are more pages
                if (
                    hasattr(response, "odata_next_link")
                    and response.odata_next_link
                ):
                    if self.is_delegated:
                        response = (
                            await self.graph_client.me.joined_teams
                            .with_url(response.odata_next_link)
                            .get()
                        )
                    else:
                        response = (
                            await self.graph_client.teams
                            .with_url(response.odata_next_link)
                            .get()
                        )
                else:
                    break
            return result
        except Exception as e:
            LOGGER.error(f"Error listing teams: {str(e)}")
            raise

    async def list_channels(
        self, team_id: str | None = None
    ) -> list[TeamsChannel]:
        """List all channels in a team.

        Args:
            team_id: Team ID to list channels for. Falls back to self.team_id.

        Returns:
            List of channels with their IDs and display names
        """
        resolved_team_id = team_id or self.team_id
        try:
            response = (
                await self.graph_client.teams.by_team_id(resolved_team_id)
                .channels.get()
            )
            result = []
            if response and response.value:
                for channel in response.value:
                    result.append(
                        TeamsChannel(
                            channel_id=channel.id or "",
                            display_name=channel.display_name or "",
                            description=channel.description,
                            team_id=resolved_team_id,
                        )
                    )
            return result
        except Exception as e:
            LOGGER.error(f"Error listing channels: {str(e)}")
            raise

    async def resolve_teams_link(self, url: str) -> dict:
        """Parse a Teams URL to extract team_id and channel_id.

        Teams URLs look like:
        https://teams.microsoft.com/l/channel/{channel_id}/{channelName}?groupId={team_id}&tenantId={tenant_id}

        Args:
            url: The Teams URL to parse

        Returns:
            Dict with 'team_id' and 'channel_id' keys
        """
        parsed = urlparse(url)
        query_params = parse_qs(parsed.query)

        team_id = query_params.get("groupId", [None])[0]

        # Channel ID is in the path: /l/channel/{channel_id}/{channelName}
        channel_id = None
        path_parts = parsed.path.strip("/").split("/")
        if len(path_parts) >= 3 and path_parts[0] == "l" and path_parts[1] == "channel":
            channel_id = unquote(path_parts[2])

        return {
            "team_id": team_id,
            "channel_id": channel_id,
        }

    async def create_channel(
        self,
        display_name: str,
        description: str | None = None,
        channel_type: str = "standard",
        team_id: str | None = None,
    ) -> TeamsChannel:
        """Create a new channel in a team.

        Args:
            display_name: Display name for the new channel
            description: Optional description for the channel
            channel_type: Channel type - "standard" or "private"
            team_id: Override team ID (falls back to self.team_id)

        Returns:
            The created channel details
        """
        resolved_team_id = team_id or self.team_id
        try:
            request_body = Channel(
                display_name=display_name,
                description=description,
                membership_type=_channel_membership_type(channel_type),
            )
            response = (
                await self.graph_client.teams.by_team_id(resolved_team_id)
                .channels.post(request_body)
            )
            if response is None:
                raise RuntimeError("Graph returned no channel for the create request")
            return TeamsChannel(
                channel_id=response.id or "",
                display_name=response.display_name or "",
                description=response.description,
                team_id=resolved_team_id,
            )
        except Exception as e:
            LOGGER.error(f"Error creating channel: {str(e)}")
            raise

    async def add_channel_members(
        self,
        channel_id: str,
        member_emails: list[str],
        team_id: str | None = None,
    ) -> list[TeamsMember]:
        """Add members to a channel by email address.

        Args:
            channel_id: Channel ID to add members to
            member_emails: List of email addresses to add
            team_id: Override team ID (falls back to self.team_id)

        Returns:
            List of members that were added
        """
        resolved_team_id = team_id or self.team_id
        added = []
        try:
            for email in member_emails:
                member = AadUserConversationMember(
                    odata_type="#microsoft.graph.aadUserConversationMember",
                    roles=["member"],
                    additional_data={
                        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{email}')",
                    },
                )
                await (
                    self.graph_client.teams.by_team_id(resolved_team_id)
                    .channels.by_channel_id(channel_id)
                    .members.post(member)
                )
                added.append(TeamsMember(name=email, email=email))
            return added
        except Exception as e:
            LOGGER.error(f"Error adding channel members: {str(e)}")
            raise

    async def start_thread(
        self,
        title: str,
        content: str,
        member_name: str | None = None,
        team_id: str | None = None,
        channel_id: str | None = None,
    ) -> TeamsThread:
        """Start a new thread in a channel.

        Args:
            title: Thread title
            content: Initial thread content
            member_name: Member name to mention in content
            team_id: Override team ID (falls back to self.team_id)
            channel_id: Override channel ID (falls back to self.teams_channel_id)

        Returns:
            Created thread details including ID
        """
        resolved_team_id = team_id or self.team_id
        resolved_channel_id = channel_id or self.teams_channel_id
        if self.adapter is None:
            return await self._start_thread_graph(
                title=title,
                content=content,
                member_name=member_name,
                team_id=resolved_team_id,
                channel_id=resolved_channel_id,
            )
        try:
            await self._initialize()

            result = TeamsThread(title=title, content=content, thread_id="")

            async def start_thread_callback(context: TurnContext):
                mention_member = None
                if member_name is not None:
                    name_lower = member_name.lower()
                    members = await TeamsInfo.get_team_members(context, resolved_team_id)
                    for member in members:
                        if member.name and member.name.lower() == name_lower:
                            mention_member = member
                    if mention_member is None:
                        dir_user = await self._find_user_in_directory(member_name)
                        if dir_user is not None:
                            mention_member = ChannelAccount(
                                id=f"29:{dir_user['id']}",
                                name=dir_user["display_name"],
                            )

                mentions = []
                if mention_member is not None:
                    result.content = (
                        f"# **{title}**\n<at>{mention_member.name}</at> {content}"
                    )
                    mention = Mention(
                        text=f"<at>{mention_member.name}</at>",
                        type="mention",
                        mentioned=ChannelAccount(
                            id=mention_member.id, name=mention_member.name
                        ),
                    )
                    mentions.append(mention)

                result.content += self._attribution_footer(html=False)
                response = await context.send_activity(
                    activity_or_text=Activity(
                        type=ActivityTypes.message,
                        topic_name=title,
                        text=result.content,
                        text_format=TextFormatTypes.markdown,
                        entities=mentions,
                    )
                )
                if response is not None:
                    result.thread_id = response.id

            await self.adapter.continue_conversation(
                bot_app_id=self.teams_app_id,
                reference=self._create_conversation_reference(
                    channel_id=resolved_channel_id
                ),
                callback=start_thread_callback,
            )

            return result
        except Exception as e:
            LOGGER.error(f"Error creating thread: {str(e)}")
            raise

    @staticmethod
    def _get_conversation_operations(context: TurnContext) -> ConversationsOperations:
        # Hack to get the connector client and reply to an existing activity
        connector_client = context.turn_state[BotAdapter.BOT_CONNECTOR_CLIENT_KEY]
        return connector_client.conversations  # pyright: ignore

    async def update_thread(
        self,
        thread_id: str,
        content: str,
        member_name: str | None = None,
        team_id: str | None = None,
        channel_id: str | None = None,
    ) -> TeamsMessage:
        """Add a message to an existing thread, mentioning a user optionally.

        Args:
            thread_id: Thread ID to update
            content: Message content to add
            member_name: Member name to mention (optional)
            team_id: Override team ID (falls back to self.team_id)
            channel_id: Override channel ID (falls back to self.teams_channel_id)

        Returns:
            Updated thread details
        """
        resolved_team_id = team_id or self.team_id
        resolved_channel_id = channel_id or self.teams_channel_id
        if self.adapter is None:
            return await self._update_thread_graph(
                thread_id=thread_id,
                content=content,
                member_name=member_name,
                team_id=resolved_team_id,
                channel_id=resolved_channel_id,
            )
        try:
            await self._initialize()

            result = TeamsMessage(thread_id=thread_id, content=content, message_id="")

            async def update_thread_callback(context: TurnContext):
                mention_member = None
                if member_name is not None:
                    name_lower = member_name.lower()
                    members = await TeamsInfo.get_team_members(context, resolved_team_id)
                    for member in members:
                        if member.name and member.name.lower() == name_lower:
                            mention_member = member
                    if mention_member is None:
                        dir_user = await self._find_user_in_directory(member_name)
                        if dir_user is not None:
                            mention_member = ChannelAccount(
                                id=f"29:{dir_user['id']}",
                                name=dir_user["display_name"],
                            )

                mentions = []
                if mention_member is not None:
                    result.content = f"<at>{mention_member.name}</at> {content}"
                    mention = Mention(
                        text=f"<at>{mention_member.name}</at>",
                        type="mention",
                        mentioned=ChannelAccount(
                            id=mention_member.id, name=mention_member.name
                        ),
                    )
                    mentions.append(mention)

                result.content += self._attribution_footer(html=False)
                reply = Activity(
                    type=ActivityTypes.message,
                    text=result.content,
                    from_property=TeamsChannelAccount(
                        id=self.teams_app_id, name="MCP Bot"
                    ),
                    conversation=ConversationAccount(id=thread_id),
                    entities=mentions,
                )
                #
                # Hack to get the connector client and reply to an existing activity
                #
                conversations = TeamsClient._get_conversation_operations(context)
                #
                # Hack to reply to conversation https://github.com/microsoft/botframework-sdk/issues/6626
                #
                conversation_id = (
                    f"{context.activity.conversation.id};messageid={thread_id}"  # pyright: ignore
                )
                response = await conversations.send_to_conversation(
                    conversation_id=conversation_id, activity=reply
                )

                if response is not None:
                    result.message_id = response.id  # pyright: ignore

            await self.adapter.continue_conversation(
                bot_app_id=self.teams_app_id,
                reference=self._create_conversation_reference(
                    channel_id=resolved_channel_id
                ),
                callback=update_thread_callback,
            )

            return result
        except Exception as e:
            LOGGER.error(f"Error updating thread: {str(e)}")
            raise

    async def get_member_by_id(
        self,
        member_id: str,
        team_id: str | None = None,
        channel_id: str | None = None,
    ) -> TeamsMember:
        resolved_team_id = team_id or self.team_id
        resolved_channel_id = channel_id or self.teams_channel_id
        if self.adapter is None:
            try:
                member = (
                    await self.graph_client.teams.by_team_id(resolved_team_id)
                    .members.by_conversation_member_id(member_id)
                    .get()
                )
                if member is not None:
                    return TeamsMember(
                        name=member.display_name or "",
                        email=getattr(member, "email", "") or "",
                    )
                return TeamsMember(name="", email="")
            except Exception as e:
                LOGGER.error(f"Error getting member by ID via Graph: {str(e)}")
                raise
        try:
            await self._initialize()

            result = TeamsMember(name="", email="")

            async def get_member_by_id_callback(context: TurnContext):
                member = await TeamsInfo.get_team_member(
                    context, resolved_team_id, member_id
                )
                result.name = member.name
                result.email = member.email

            await self.adapter.continue_conversation(
                bot_app_id=self.teams_app_id,
                reference=self._create_conversation_reference(
                    channel_id=resolved_channel_id
                ),
                callback=get_member_by_id_callback,
            )
            return result
        except Exception as e:
            LOGGER.error(f"Error updating thread: {str(e)}")
            raise

    async def read_threads(
        self,
        limit: int = 50,
        cursor: str | None = None,
        team_id: str | None = None,
        channel_id: str | None = None,
    ) -> PagedTeamsMessages:
        """Read all threads in configured teams channel.

        Args:
            cursor: The pagination cursor.
            limit: The pagination page size
            team_id: Override team ID (falls back to self.team_id)
            channel_id: Override channel ID (falls back to self.teams_channel_id)

        Returns:
            Paged team channel messages containing
        """
        resolved_team_id = team_id or self.team_id
        resolved_channel_id = channel_id or self.teams_channel_id
        try:
            query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                top=limit
            )
            request = RequestConfiguration(query_parameters=query)
            if cursor is not None:
                response = (
                    await self.graph_client.teams.by_team_id(resolved_team_id)
                    .channels.by_channel_id(resolved_channel_id)
                    .messages.with_url(cursor)
                    .get(request_configuration=request)
                )
            else:
                response = (
                    await self.graph_client.teams.by_team_id(resolved_team_id)
                    .channels.by_channel_id(resolved_channel_id)
                    .messages.get(request_configuration=request)
                )

            result = PagedTeamsMessages(
                cursor=response.odata_next_link,  # pyright: ignore
                limit=limit,
                total=response.odata_count,  # pyright: ignore
                items=[],
            )
            if response.value is not None:  # pyright: ignore
                for message in response.value:  # pyright: ignore
                    extracted = _extract_attachments(message.attachments)
                    result.items.append(
                        TeamsMessage(
                            message_id=message.id,  # pyright: ignore
                            content=message.body.content,  # pyright: ignore
                            thread_id=message.id,  # pyright: ignore
                            attachments=extracted,
                        )
                    )

            return result
        except Exception as e:
            LOGGER.error(f"Error reading thread: {str(e)}")
            raise

    async def read_thread_replies(
        self,
        thread_id: str,
        limit: int = 50,
        cursor: str | None = None,
        team_id: str | None = None,
        channel_id: str | None = None,
    ) -> PagedTeamsMessages:
        """Read all replies in a thread.

        Args:
            thread_id: Thread ID to read
            cursor: The pagination cursor
            limit: The pagination page size
            team_id: Override team ID (falls back to self.team_id)
            channel_id: Override channel ID (falls back to self.teams_channel_id)

        Returns:
            List of thread messages
        """
        resolved_team_id = team_id or self.team_id
        resolved_channel_id = channel_id or self.teams_channel_id
        try:
            params = RepliesRequestBuilder.RepliesRequestBuilderGetQueryParameters(
                top=limit
            )
            request = RequestConfiguration(query_parameters=params)

            if cursor is not None:
                replies = (
                    await self.graph_client.teams.by_team_id(resolved_team_id)
                    .channels.by_channel_id(resolved_channel_id)
                    .messages.by_chat_message_id(thread_id)
                    .replies.with_url(cursor)
                    .get(request_configuration=request)
                )
            else:
                replies = (
                    await self.graph_client.teams.by_team_id(resolved_team_id)
                    .channels.by_channel_id(resolved_channel_id)
                    .messages.by_chat_message_id(thread_id)
                    .replies.get(request_configuration=request)
                )

            result = PagedTeamsMessages(
                cursor=cursor,
                limit=limit,
                total=replies.odata_count,  # pyright: ignore
                items=[],
            )

            if replies is not None and replies.value is not None:
                for reply in replies.value:
                    extracted = _extract_attachments(reply.attachments)
                    result.items.append(
                        TeamsMessage(
                            message_id=reply.id,  # pyright: ignore
                            content=reply.body.content,  # pyright: ignore
                            thread_id=reply.reply_to_id,  # pyright: ignore
                            attachments=extracted,
                        )
                    )

            return result
        except Exception as e:
            LOGGER.error(f"Error reading thread: {str(e)}")
            raise

    async def read_message(
        self,
        message_id: str,
        team_id: str | None = None,
        channel_id: str | None = None,
    ) -> ChatMessage | None:
        resolved_team_id = team_id or self.team_id
        resolved_channel_id = channel_id or self.teams_channel_id
        try:
            query = ChatMessageItemRequestBuilder.ChatMessageItemRequestBuilderGetQueryParameters()
            request = RequestConfiguration(query_parameters=query)
            response = (
                await self.graph_client.teams.by_team_id(resolved_team_id)
                .channels.by_channel_id(resolved_channel_id)
                .messages.by_chat_message_id(chat_message_id=message_id)
                .get(request_configuration=request)
            )
            return response
        except Exception as e:
            LOGGER.error(f"Error reading thread: {str(e)}")
            raise

    async def list_members(
        self,
        team_id: str | None = None,
        channel_id: str | None = None,
    ) -> list[TeamsMember]:
        """List all members in the configured team.

        Args:
            team_id: Override team ID (falls back to self.team_id)
            channel_id: Override channel ID (falls back to self.teams_channel_id)

        Returns:
            List of team member details
        """
        resolved_team_id = team_id or self.team_id
        resolved_channel_id = channel_id or self.teams_channel_id
        if self.adapter is None:
            return await self._list_members_graph(team_id=resolved_team_id)
        try:
            await self._initialize()
            result = []

            async def list_members_callback(context: TurnContext):
                members = await TeamsInfo.get_team_members(context, resolved_team_id)
                for member in members:
                    result.append(TeamsMember(name=member.name, email=member.email))

            await self.adapter.continue_conversation(
                bot_app_id=self.teams_app_id,
                reference=self._create_conversation_reference(
                    channel_id=resolved_channel_id
                ),
                callback=list_members_callback,
            )
            return result
        except Exception as e:
            LOGGER.error(f"Error listing members: {str(e)}")
            raise

    async def _list_members_graph(
        self,
        team_id: str,
    ) -> list[TeamsMember]:
        """List team members via Graph API (delegated auth mode).

        Args:
            team_id: Team ID to list members for.

        Returns:
            List of team member details
        """
        try:
            response = (
                await self.graph_client.teams.by_team_id(team_id)
                .members.get()
            )
            result = []
            if response and response.value:
                for member in response.value:
                    result.append(
                        TeamsMember(
                            name=member.display_name or "",
                            email=getattr(member, "email", "") or "",
                        )
                    )
            return result
        except Exception as e:
            LOGGER.error(f"Error listing members via Graph: {str(e)}")
            raise

    async def _find_member_by_name_graph(
        self,
        member_name: str,
        team_id: str,
    ) -> AadUserConversationMember | None:
        """Find a team member by display name via Graph API.

        Returns the raw AadUserConversationMember so the caller can access
        the user_id needed for ChatMessageMention. Uses case-insensitive
        matching.

        Args:
            member_name: Display name to search for.
            team_id: Team ID to search in.

        Returns:
            The matching member or None.
        """
        try:
            response = (
                await self.graph_client.teams.by_team_id(team_id)
                .members.get()
            )
            if response and response.value:
                name_lower = member_name.lower()
                for member in response.value:
                    if (
                        member.display_name
                        and member.display_name.lower() == name_lower
                    ):
                        return member  # pyright: ignore[reportReturnType]
            return None
        except Exception as e:
            LOGGER.error(f"Error finding member by name via Graph: {str(e)}")
            raise

    async def _find_user_in_directory(
        self,
        display_name: str,
    ) -> dict | None:
        """Search the organization directory for a user by display name.

        Falls back to directory search when user is not found in team members.
        This allows @mentioning users who are in the org but not in the team.

        Args:
            display_name: Display name to search for.

        Returns:
            Dict with 'id' and 'display_name' keys, or None if not found.
        """
        try:
            from msgraph.generated.users.users_request_builder import (
                UsersRequestBuilder,
            )

            query_params = (
                UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                    filter=f"displayName eq '{display_name}'",
                    select=["id", "displayName", "mail"],
                    top=1,
                )
            )
            config = RequestConfiguration(
                query_parameters=query_params,
            )
            response = await self.graph_client.users.get(
                request_configuration=config,
            )
            if response and response.value and len(response.value) > 0:
                user = response.value[0]
                return {
                    "id": user.id or "",
                    "display_name": user.display_name or display_name,
                }
            LOGGER.warning(
                f"User '{display_name}' not found in org directory"
            )
            return None
        except Exception as e:
            LOGGER.warning(
                f"Directory search for '{display_name}' failed: {str(e)}"
            )
            return None

    async def _start_thread_graph(
        self,
        title: str,
        content: str,
        member_name: str | None,
        team_id: str,
        channel_id: str,
    ) -> TeamsThread:
        """Start a new thread via Graph API (delegated auth mode).

        Args:
            title: Thread title
            content: Thread content
            member_name: Optional member name to mention
            team_id: Team ID
            channel_id: Channel ID

        Returns:
            Created thread details
        """
        try:
            # Threads are always sent as HTML; convert any markdown tables
            # so Teams renders a real grid instead of literal pipe text.
            content = convert_markdown_tables_to_html(content)
            body_content = f"<h1><b>{title}</b></h1>\n{content}"
            mentions = []

            if member_name is not None:
                found = await self._find_member_by_name_graph(member_name, team_id)
                resolved_name = member_name
                user_id = None
                if found is not None:
                    user_id = getattr(found, "user_id", None) or ""
                    resolved_name = getattr(found, "display_name", member_name) or member_name
                else:
                    # Fallback: search org directory
                    dir_user = await self._find_user_in_directory(member_name)
                    if dir_user is not None:
                        user_id = dir_user["id"]
                        resolved_name = dir_user["display_name"]

                if user_id is not None:
                    mention = ChatMessageMention(
                        id=0,
                        mention_text=resolved_name,
                        mentioned=ChatMessageMentionedIdentitySet(
                            user=Identity(
                                id=user_id,
                                display_name=resolved_name,
                            ),
                        ),
                    )
                    mentions.append(mention)
                    body_content = (
                        f"<h1><b>{title}</b></h1>\n"
                        f"<at id=\"0\">{resolved_name}</at> {content}"
                    )

            body_content += self._attribution_footer(html=True)

            message = ChatMessage(
                subject=title,
                body=ItemBody(
                    content_type=BodyType.Html,
                    content=body_content,
                ),
                mentions=mentions if mentions else None,
            )

            response = (
                await self.graph_client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.post(message)
            )

            thread_id = response.id if response else ""
            return TeamsThread(
                thread_id=thread_id or "",
                title=title,
                content=body_content,
            )
        except Exception as e:
            LOGGER.error(f"Error creating thread via Graph: {str(e)}")
            raise

    async def _update_thread_graph(
        self,
        thread_id: str,
        content: str,
        member_name: str | None,
        team_id: str,
        channel_id: str,
    ) -> TeamsMessage:
        """Reply to an existing thread via Graph API (delegated auth mode).

        Args:
            thread_id: Thread ID to reply to
            content: Reply content
            member_name: Optional member name to mention
            team_id: Team ID
            channel_id: Channel ID

        Returns:
            The reply message details
        """
        try:
            # Replies are always sent as HTML; convert any markdown tables
            # so Teams renders a real grid instead of literal pipe text.
            content = convert_markdown_tables_to_html(content)
            body_content = content
            mentions = []

            if member_name is not None:
                found = await self._find_member_by_name_graph(member_name, team_id)
                resolved_name = member_name
                user_id = None
                if found is not None:
                    user_id = getattr(found, "user_id", None) or ""
                    resolved_name = getattr(found, "display_name", member_name) or member_name
                else:
                    # Fallback: search org directory
                    dir_user = await self._find_user_in_directory(member_name)
                    if dir_user is not None:
                        user_id = dir_user["id"]
                        resolved_name = dir_user["display_name"]

                if user_id is not None:
                    mention = ChatMessageMention(
                        id=0,
                        mention_text=resolved_name,
                        mentioned=ChatMessageMentionedIdentitySet(
                            user=Identity(
                                id=user_id,
                                display_name=resolved_name,
                            ),
                        ),
                    )
                    mentions.append(mention)
                    body_content = f"<at id=\"0\">{resolved_name}</at> {content}"

            body_content += self._attribution_footer(html=True)

            reply = ChatMessage(
                body=ItemBody(
                    content_type=BodyType.Html,
                    content=body_content,
                ),
                mentions=mentions if mentions else None,
            )

            response = (
                await self.graph_client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.by_chat_message_id(thread_id)
                .replies.post(reply)
            )

            message_id = response.id if response else ""
            return TeamsMessage(
                thread_id=thread_id,
                content=body_content,
                message_id=message_id or "",
            )
        except Exception as e:
            LOGGER.error(f"Error replying to thread via Graph: {str(e)}")
            raise

    async def get_member_by_name(
        self,
        name: str,
        team_id: str | None = None,
        channel_id: str | None = None,
    ) -> TeamsMember | None:
        members = await self.list_members(team_id=team_id, channel_id=channel_id)
        name_lower = name.lower()
        for member in members:
            if member.name and member.name.lower() == name_lower:
                return member

    async def find_available_times(
        self,
        attendee_emails: list[str],
        date: str,
        start_hour: int = 8,
        end_hour: int = 17,
        duration_minutes: int = 60,
        exclude_start_hour: int | None = 12,
        exclude_end_hour: int | None = 13,
        timezone: str = "Central Standard Time",
    ) -> AvailableTimesResult:
        """Find common available time slots for multiple attendees.

        Uses the Graph API getSchedule endpoint to query free/busy info,
        then computes contiguous free slots of the requested duration.

        Args:
            attendee_emails: List of email addresses to check availability for
            date: Date to check in YYYY-MM-DD format
            start_hour: Start of business hours (default 8)
            end_hour: End of business hours (default 17)
            duration_minutes: Required meeting duration in minutes (default 60)
            exclude_start_hour: Start of excluded time block like lunch (default 12, None to disable)
            exclude_end_hour: End of excluded time block (default 13, None to disable)
            timezone: Timezone for the query (default "Central Standard Time")

        Returns:
            AvailableTimesResult with common free time slots
        """
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.users.item.calendar.get_schedule.get_schedule_post_request_body import (
            GetSchedulePostRequestBody,
        )

        try:
            request_body = GetSchedulePostRequestBody()
            request_body.schedules = attendee_emails

            start_time = DateTimeTimeZone()
            start_time.date_time = f"{date}T{start_hour:02d}:00:00"
            start_time.time_zone = timezone
            request_body.start_time = start_time

            end_time = DateTimeTimeZone()
            end_time.date_time = f"{date}T{end_hour:02d}:00:00"
            end_time.time_zone = timezone
            request_body.end_time = end_time

            request_body.availability_view_interval = 15

            # Use /me for delegated, /users/{email} for app mode
            if self.adapter is None:
                result = await self.graph_client.me.calendar.get_schedule.post(request_body)
            else:
                # In app mode, use the first attendee as the calling user context
                result = await (
                    self.graph_client.users.by_user_id(attendee_emails[0])
                    .calendar.get_schedule.post(request_body)
                )

            # Parse availability views to find common free slots
            total_slots = (end_hour - start_hour) * 4  # 15-min intervals
            # Initialize all slots as free
            common_free = [True] * total_slots

            if result and result.value:
                for schedule in result.value:
                    view = schedule.availability_view or ""
                    for i, char in enumerate(view):
                        if i < total_slots and char != "0":
                            common_free[i] = False

            # Exclude lunch/break period
            if exclude_start_hour is not None and exclude_end_hour is not None:
                for hour in range(exclude_start_hour, exclude_end_hour):
                    for quarter in range(4):
                        idx = (hour - start_hour) * 4 + quarter
                        if 0 <= idx < total_slots:
                            common_free[idx] = False

            # Find contiguous free blocks of sufficient duration
            required_slots = duration_minutes // 15
            available_slots: list[TimeSlot] = []

            i = 0
            while i <= total_slots - required_slots:
                if all(common_free[i : i + required_slots]):
                    slot_start_minutes = start_hour * 60 + i * 15
                    slot_end_minutes = slot_start_minutes + duration_minutes
                    s_h, s_m = divmod(slot_start_minutes, 60)
                    e_h, e_m = divmod(slot_end_minutes, 60)
                    available_slots.append(
                        TimeSlot(
                            start=f"{date}T{s_h:02d}:{s_m:02d}:00",
                            end=f"{date}T{e_h:02d}:{e_m:02d}:00",
                        )
                    )
                    i += required_slots  # Skip past this block
                else:
                    i += 1

            return AvailableTimesResult(
                date=date,
                slots=available_slots,
                attendees=attendee_emails,
                duration_minutes=duration_minutes,
            )
        except Exception as e:
            LOGGER.error(f"Error finding available times: {str(e)}")
            raise

    async def create_meeting(
        self,
        subject: str,
        start_datetime: str,
        end_datetime: str,
        attendee_emails: list[str],
        body: str | None = None,
        timezone: str = "Central Standard Time",
        is_online: bool = True,
        recurrence_type: str | None = None,
        recurrence_interval: int = 1,
        days_of_week: list[str] | None = None,
        recurrence_end_date: str | None = None,
        organizer_email: str | None = None,
    ) -> MeetingResult:
        """Create a calendar event (optionally recurring) with a Teams meeting link.

        Args:
            subject: Meeting subject/title
            start_datetime: Start time in ISO format (e.g., "2026-03-04T09:00:00")
            end_datetime: End time in ISO format (e.g., "2026-03-04T09:30:00")
            attendee_emails: List of attendee email addresses
            body: Optional HTML body content for the meeting
            timezone: Timezone (default "Central Standard Time")
            is_online: Whether to create a Teams meeting link (default True)
            recurrence_type: Recurrence pattern type - "daily", "weekly", or None for one-time
            recurrence_interval: Interval between recurrences (default 1)
            days_of_week: Days for weekly recurrence (e.g., ["wednesday"])
            recurrence_end_date: End date for recurrence in YYYY-MM-DD format (None for no end)
            organizer_email: Email of the organizer (required for app mode, optional for delegated)

        Returns:
            MeetingResult with event details and Teams join URL
        """
        from msgraph.generated.models.attendee import Attendee
        from msgraph.generated.models.attendee_type import AttendeeType
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.models.email_address import (
            EmailAddress as GraphEmailAddress,
        )
        from msgraph.generated.models.event import Event
        from msgraph.generated.models.online_meeting_provider_type import (
            OnlineMeetingProviderType,
        )

        try:
            event = Event()
            event.subject = subject

            if body:
                event.body = ItemBody(
                    content_type=BodyType.Html,
                    content=body,
                )

            start = DateTimeTimeZone()
            start.date_time = start_datetime
            start.time_zone = timezone
            event.start = start

            end = DateTimeTimeZone()
            end.date_time = end_datetime
            end.time_zone = timezone
            event.end = end

            attendees = []
            for email in attendee_emails:
                attendee = Attendee()
                attendee.email_address = GraphEmailAddress()
                attendee.email_address.address = email
                attendee.type = AttendeeType.Required
                attendees.append(attendee)
            event.attendees = attendees

            if is_online:
                event.is_online_meeting = True
                event.online_meeting_provider = OnlineMeetingProviderType.TeamsForBusiness

            # Add recurrence if specified
            is_recurring = False
            if recurrence_type is not None:
                is_recurring = True
                import datetime

                from msgraph.generated.models.day_of_week import DayOfWeek
                from msgraph.generated.models.patterned_recurrence import (
                    PatternedRecurrence,
                )
                from msgraph.generated.models.recurrence_pattern import (
                    RecurrencePattern,
                )
                from msgraph.generated.models.recurrence_pattern_type import (
                    RecurrencePatternType,
                )
                from msgraph.generated.models.recurrence_range import RecurrenceRange
                from msgraph.generated.models.recurrence_range_type import (
                    RecurrenceRangeType,
                )

                day_map = {
                    "sunday": DayOfWeek.Sunday,
                    "monday": DayOfWeek.Monday,
                    "tuesday": DayOfWeek.Tuesday,
                    "wednesday": DayOfWeek.Wednesday,
                    "thursday": DayOfWeek.Thursday,
                    "friday": DayOfWeek.Friday,
                    "saturday": DayOfWeek.Saturday,
                }

                recurrence = PatternedRecurrence()
                pattern = RecurrencePattern()

                if recurrence_type == "daily":
                    pattern.type = RecurrencePatternType.Daily
                elif recurrence_type == "weekly":
                    pattern.type = RecurrencePatternType.Weekly
                    if days_of_week:
                        pattern.days_of_week = [day_map[d.lower()] for d in days_of_week]
                    pattern.first_day_of_week = DayOfWeek.Sunday

                pattern.interval = recurrence_interval
                recurrence.pattern = pattern

                range_ = RecurrenceRange()
                # Extract start date from start_datetime
                start_date_str = start_datetime.split("T")[0]
                range_.start_date = datetime.date.fromisoformat(start_date_str)

                if recurrence_end_date:
                    range_.type = RecurrenceRangeType.EndDate
                    range_.end_date = datetime.date.fromisoformat(recurrence_end_date)
                else:
                    range_.type = RecurrenceRangeType.NoEnd

                recurrence.range = range_
                event.recurrence = recurrence

            # Create the event
            if self.adapter is None:
                # Delegated mode: use /me
                response = await self.graph_client.me.events.post(event)
            else:
                # App mode: must specify organizer
                if not organizer_email:
                    organizer_email = attendee_emails[0] if attendee_emails else ""
                response = await (
                    self.graph_client.users.by_user_id(organizer_email)
                    .events.post(event)
                )

            join_url = None
            if response and response.online_meeting:
                join_url = response.online_meeting.join_url

            return MeetingResult(
                event_id=(response.id or "") if response else "",
                subject=subject,
                start=start_datetime,
                end=end_datetime,
                join_url=join_url,
                attendees=attendee_emails,
                is_recurring=is_recurring,
            )
        except Exception as e:
            LOGGER.error(f"Error creating meeting: {str(e)}")
            raise

    async def cancel_meeting(
        self,
        event_id: str,
        organizer_email: str | None = None,
        comment: str | None = None,
    ) -> dict:
        """Cancel (delete) a calendar event by its event ID.

        Args:
            event_id: The event ID returned from create_meeting
            organizer_email: Email of the organizer (required for app mode, optional for delegated)
            comment: Optional cancellation message sent to attendees

        Returns:
            dict with success status
        """
        try:
            if self.adapter is None:
                # Delegated mode: use /me
                if comment:
                    from msgraph.generated.users.item.events.item.cancel.cancel_post_request_body import (
                        CancelPostRequestBody,
                    )
                    body = CancelPostRequestBody()
                    body.comment = comment
                    await self.graph_client.me.events.by_event_id(event_id).cancel.post(body)
                else:
                    await self.graph_client.me.events.by_event_id(event_id).delete()
            else:
                # App mode: must specify organizer
                if not organizer_email:
                    raise ValueError("organizer_email is required in app auth mode")
                if comment:
                    from msgraph.generated.users.item.events.item.cancel.cancel_post_request_body import (
                        CancelPostRequestBody,
                    )
                    body = CancelPostRequestBody()
                    body.comment = comment
                    await (
                        self.graph_client.users.by_user_id(organizer_email)
                        .events.by_event_id(event_id)
                        .cancel.post(body)
                    )
                else:
                    await (
                        self.graph_client.users.by_user_id(organizer_email)
                        .events.by_event_id(event_id)
                        .delete()
                    )

            return {"success": True, "event_id": event_id, "message": "Meeting cancelled successfully"}
        except Exception as e:
            LOGGER.error(f"Error cancelling meeting: {str(e)}")
            raise

    @staticmethod
    def _parse_event(event) -> CalendarEvent:
        """Parse a Graph API Event object into a CalendarEvent model."""
        attendee_emails = []
        if event.attendees:
            for att in event.attendees:
                if att.email_address and att.email_address.address:
                    attendee_emails.append(att.email_address.address)

        organizer_email = ""
        if event.organizer and event.organizer.email_address:
            organizer_email = event.organizer.email_address.address or ""

        start_str = ""
        tz = ""
        if event.start:
            start_str = event.start.date_time or ""
            tz = event.start.time_zone or ""

        end_str = ""
        if event.end:
            end_str = event.end.date_time or ""

        join_url = None
        if event.online_meeting:
            join_url = event.online_meeting.join_url

        is_recurring = event.recurrence is not None
        series_master_id = getattr(event, "series_master_id", None)
        recurrence_desc = None
        if event.recurrence and event.recurrence.pattern:
            pattern = event.recurrence.pattern
            pattern_type = getattr(pattern, "type", None)
            interval = getattr(pattern, "interval", 1)
            days = getattr(pattern, "days_of_week", None)
            if pattern_type:
                type_name = pattern_type.value if hasattr(pattern_type, "value") else str(pattern_type)
                if days:
                    day_names = [d.value if hasattr(d, "value") else str(d) for d in days]
                    recurrence_desc = f"{type_name} every {interval} week(s) on {', '.join(day_names)}"
                else:
                    recurrence_desc = f"{type_name} every {interval} interval(s)"

        return CalendarEvent(
            event_id=event.id or "",
            subject=event.subject or "",
            start=start_str,
            end=end_str,
            timezone=tz,
            organizer=organizer_email,
            attendees=attendee_emails,
            body_preview=event.body_preview or "",
            is_online_meeting=event.is_online_meeting or False,
            join_url=join_url,
            is_recurring=is_recurring,
            series_master_id=series_master_id,
            recurrence_pattern=recurrence_desc,
            web_link=event.web_link or None,
        )

    async def list_events(
        self,
        user_email: str | None = None,
        subject_filter: str | None = None,
        start_date: str | None = None,
        end_date: str | None = None,
        top: int = 25,
        order_by: str = "start/dateTime desc",
    ) -> CalendarEventList:
        """List/search calendar events with optional filtering.

        Args:
            user_email: User whose calendar to query (required for app mode, optional for delegated)
            subject_filter: Filter events whose subject contains this text
            start_date: Filter events starting on or after this date (YYYY-MM-DD)
            end_date: Filter events starting on or before this date (YYYY-MM-DD)
            top: Maximum number of events to return (default 25)
            order_by: Sort order (default "start/dateTime desc")

        Returns:
            CalendarEventList with matching events
        """
        from msgraph.generated.users.item.events.events_request_builder import (
            EventsRequestBuilder,
        )

        try:
            # Build $filter clauses
            filter_parts = []
            if subject_filter:
                # Use contains for subject search
                escaped = subject_filter.replace("'", "''")
                filter_parts.append(f"contains(subject,'{escaped}')")
            if start_date:
                filter_parts.append(f"start/dateTime ge '{start_date}T00:00:00'")
            if end_date:
                filter_parts.append(f"start/dateTime le '{end_date}T23:59:59'")

            filter_str = " and ".join(filter_parts) if filter_parts else None

            query_params = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters(
                select=[
                    "id", "subject", "start", "end", "organizer", "attendees",
                    "bodyPreview", "isOnlineMeeting", "onlineMeeting",
                    "recurrence", "seriesMasterId", "webLink",
                ],
                filter=filter_str,
                orderby=[order_by] if order_by else None,
                top=top,
            )
            config = RequestConfiguration(query_parameters=query_params)

            if self.adapter is None:
                response = await self.graph_client.me.events.get(request_configuration=config)
            else:
                if not user_email:
                    raise ValueError("user_email is required in app auth mode")
                response = await (
                    self.graph_client.users.by_user_id(user_email)
                    .events.get(request_configuration=config)
                )

            events = []
            if response and response.value:
                for ev in response.value:
                    events.append(self._parse_event(ev))

            return CalendarEventList(
                events=events,
                total_returned=len(events),
            )
        except Exception as e:
            LOGGER.error(f"Error listing events: {str(e)}")
            raise

    async def update_event(
        self,
        event_id: str,
        user_email: str | None = None,
        subject: str | None = None,
        body: str | None = None,
        start_datetime: str | None = None,
        end_datetime: str | None = None,
        timezone: str | None = None,
        attendee_emails: list[str] | None = None,
        is_online: bool | None = None,
        recurrence_type: str | None = None,
        recurrence_interval: int | None = None,
        days_of_week: list[str] | None = None,
        recurrence_end_date: str | None = None,
        remove_recurrence: bool = False,
    ) -> CalendarEvent:
        """Update an existing calendar event.

        For recurring meetings, updating the series master event updates all
        occurrences. Pass the series master event ID to update the whole series.

        Args:
            event_id: The event ID to update (use series master ID for recurring series)
            user_email: User whose calendar contains the event (required for app mode)
            subject: New subject (None to keep current)
            body: New HTML body (None to keep current)
            start_datetime: New start time in ISO format (None to keep current)
            end_datetime: New end time in ISO format (None to keep current)
            timezone: New timezone (None to keep current)
            attendee_emails: New attendee list - replaces existing (None to keep current)
            is_online: Whether to enable Teams meeting link (None to keep current)
            recurrence_type: New recurrence type - "daily", "weekly" (None to keep current)
            recurrence_interval: Interval between recurrences (None to keep current)
            days_of_week: Days for weekly recurrence (None to keep current)
            recurrence_end_date: End date for recurrence in YYYY-MM-DD (None to keep current)
            remove_recurrence: Set True to remove recurrence from a recurring event

        Returns:
            Updated CalendarEvent
        """
        from msgraph.generated.models.attendee import Attendee
        from msgraph.generated.models.attendee_type import AttendeeType
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.models.email_address import (
            EmailAddress as GraphEmailAddress,
        )
        from msgraph.generated.models.event import Event
        from msgraph.generated.models.online_meeting_provider_type import (
            OnlineMeetingProviderType,
        )

        try:
            event = Event()

            if subject is not None:
                event.subject = subject

            if body is not None:
                event.body = ItemBody(
                    content_type=BodyType.Html,
                    content=body,
                )

            if start_datetime is not None:
                start = DateTimeTimeZone()
                start.date_time = start_datetime
                start.time_zone = timezone or "Central Standard Time"
                event.start = start

            if end_datetime is not None:
                end = DateTimeTimeZone()
                end.date_time = end_datetime
                end.time_zone = timezone or "Central Standard Time"
                event.end = end

            if attendee_emails is not None:
                attendees = []
                for email in attendee_emails:
                    attendee = Attendee()
                    attendee.email_address = GraphEmailAddress()
                    attendee.email_address.address = email
                    attendee.type = AttendeeType.Required
                    attendees.append(attendee)
                event.attendees = attendees

            if is_online is not None:
                event.is_online_meeting = is_online
                if is_online:
                    event.online_meeting_provider = OnlineMeetingProviderType.TeamsForBusiness

            if remove_recurrence:
                event.recurrence = None
            elif recurrence_type is not None:
                import datetime

                from msgraph.generated.models.day_of_week import DayOfWeek
                from msgraph.generated.models.patterned_recurrence import (
                    PatternedRecurrence,
                )
                from msgraph.generated.models.recurrence_pattern import (
                    RecurrencePattern,
                )
                from msgraph.generated.models.recurrence_pattern_type import (
                    RecurrencePatternType,
                )
                from msgraph.generated.models.recurrence_range import RecurrenceRange
                from msgraph.generated.models.recurrence_range_type import (
                    RecurrenceRangeType,
                )

                day_map = {
                    "sunday": DayOfWeek.Sunday,
                    "monday": DayOfWeek.Monday,
                    "tuesday": DayOfWeek.Tuesday,
                    "wednesday": DayOfWeek.Wednesday,
                    "thursday": DayOfWeek.Thursday,
                    "friday": DayOfWeek.Friday,
                    "saturday": DayOfWeek.Saturday,
                }

                recurrence = PatternedRecurrence()
                pattern = RecurrencePattern()

                if recurrence_type == "daily":
                    pattern.type = RecurrencePatternType.Daily
                elif recurrence_type == "weekly":
                    pattern.type = RecurrencePatternType.Weekly
                    if days_of_week:
                        pattern.days_of_week = [day_map[d.lower()] for d in days_of_week]
                    pattern.first_day_of_week = DayOfWeek.Sunday

                pattern.interval = recurrence_interval or 1
                recurrence.pattern = pattern

                range_ = RecurrenceRange()
                if start_datetime:
                    start_date_str = start_datetime.split("T")[0]
                    range_.start_date = datetime.date.fromisoformat(start_date_str)

                if recurrence_end_date:
                    range_.type = RecurrenceRangeType.EndDate
                    range_.end_date = datetime.date.fromisoformat(recurrence_end_date)
                else:
                    range_.type = RecurrenceRangeType.NoEnd

                recurrence.range = range_
                event.recurrence = recurrence

            # Perform the update
            if self.adapter is None:
                response = await self.graph_client.me.events.by_event_id(event_id).patch(event)
            else:
                if not user_email:
                    raise ValueError("user_email is required in app auth mode")
                response = await (
                    self.graph_client.users.by_user_id(user_email)
                    .events.by_event_id(event_id)
                    .patch(event)
                )

            return self._parse_event(response)
        except Exception as e:
            LOGGER.error(f"Error updating event: {str(e)}")
            raise

    # In-memory directory cache built from team memberships.
    # Populated lazily on first resolve_user() call per session.
    _directory_cache: list[dict] | None = None

    async def _build_directory_cache(self) -> list[dict]:
        """Build a user directory from the caller's Teams memberships.

        Fetches /me/joinedTeams, then iterates each team's member list
        to build a de-duplicated directory. No extra permissions needed
        beyond Team.ReadBasic.All (delegated).
        """
        if self._directory_cache is not None:
            return self._directory_cache

        LOGGER.info("Building directory cache from team memberships...")
        seen_emails: set[str] = set()
        directory: list[dict] = []

        try:
            teams = await self.graph_client.me.joined_teams.get()
            if not teams or not teams.value:
                LOGGER.warning("No joined teams found for /me")
                self._directory_cache = []
                return []

            for team in teams.value:
                try:
                    if not team.id:
                        continue
                    members = await (
                        self.graph_client.teams
                        .by_team_id(team.id)
                        .members.get()
                    )
                    if not members or not members.value:
                        continue
                    for m in members.value:
                        email = ""
                        # ConversationMember has no typed `email`; some member
                        # subtypes carry one, others expose it in additional_data.
                        member_email = getattr(m, "email", None)
                        if member_email:
                            email = member_email.lower()
                        elif (
                            m.additional_data
                            and "email" in m.additional_data
                        ):
                            email = m.additional_data["email"].lower()
                        if not email or email in seen_emails:
                            continue
                        seen_emails.add(email)
                        directory.append({
                            "display_name": m.display_name or "",
                            "email": email,
                            "user_id": (
                                m.additional_data.get("userId", "")
                                if m.additional_data else ""
                            ),
                        })
                except Exception:
                    LOGGER.debug(
                        f"Could not list members for team {team.display_name}"
                    )
                    continue
        except Exception as e:
            LOGGER.error(f"Failed to build directory cache: {e}")
            self._directory_cache = []
            return []

        LOGGER.info(f"Directory cache built: {len(directory)} unique users")
        self._directory_cache = directory
        return directory

    @staticmethod
    def _fuzzy_score(query: str, name: str, email: str) -> float:
        """Score how well a query matches a name/email (higher = better).

        Handles partial names, case insensitivity, word-order
        independence, and basic typo tolerance via character overlap.
        """
        q = query.lower().strip()
        n = name.lower()
        e = email.lower().split("@")[0]  # just the local part

        score = 0.0

        # Exact match on full name or email prefix
        if q == n or q == e:
            return 1.0

        q_parts = q.split()
        n_parts = n.split()

        # Each query word that starts a name word (prefix match)
        for qp in q_parts:
            for np in n_parts:
                if np.startswith(qp):
                    score += 0.4
                elif qp in np:
                    score += 0.25

        # Email prefix partial match
        if q in e:
            score += 0.3

        # Character overlap ratio for typo tolerance
        q_chars = set(q.replace(" ", ""))
        n_chars = set(n.replace(" ", ""))
        if q_chars and n_chars:
            overlap = len(q_chars & n_chars) / max(len(q_chars), len(n_chars))
            score += overlap * 0.3

        return min(score, 1.0)

    async def resolve_user(
        self,
        query: str,
        max_results: int = 5,
    ) -> ResolvedUserList:
        """Resolve a fuzzy name, partial name, or email to matching users.

        Builds a directory from the caller's Teams memberships and does
        local fuzzy matching. Supports partial names ("Alex"),
        full names ("Alex Rivera"), misspellings ("alx rivra"), and email
        prefixes ("alex.rivera").

        Args:
            query: Name, partial name, or email to search for.
            max_results: Maximum number of results to return (default 5).

        Returns:
            ResolvedUserList with matched users sorted by relevance.
        """
        directory = await self._build_directory_cache()
        if not directory:
            return ResolvedUserList(
                users=[], query=query, total_matches=0,
            )

        # Score every person in the directory
        scored: list[tuple[float, dict]] = []
        for entry in directory:
            s = self._fuzzy_score(query, entry["display_name"], entry["email"])
            if s > 0.15:  # minimum threshold
                scored.append((s, entry))

        # Sort by score descending, take top N
        scored.sort(key=lambda x: x[0], reverse=True)
        top = scored[:max_results]

        users = [
            ResolvedUser(
                display_name=entry["display_name"],
                email=entry["email"],
                user_id=entry.get("user_id", ""),
                relevance_score=round(score, 2),
            )
            for score, entry in top
        ]

        return ResolvedUserList(
            users=users,
            query=query,
            total_matches=len(users),
        )

    async def _resolve_recipient(self, recipient: str) -> str:
        """Resolve a recipient string to an email address.

        If the input looks like an email (contains @), returns it as-is.
        Otherwise, uses resolve_user() for fuzzy name matching.
        Raises ValueError if no match or ambiguous results.
        """
        if "@" in recipient:
            return recipient

        result = await self.resolve_user(recipient, max_results=3)
        if result.total_matches == 0:
            raise ValueError(
                f"Could not find anyone matching '{recipient}'. "
                f"Try using their full name or email address."
            )
        if result.total_matches == 1:
            match = result.users[0]
            LOGGER.info(
                f"Resolved '{recipient}' -> {match.display_name} ({match.email})"
            )
            return match.email

        # Multiple matches — check if the top result is clearly the best
        top = result.users[0]
        runner_up = result.users[1]
        # If the query closely matches the top result's name, use it
        query_lower = recipient.lower().strip()
        top_name_lower = top.display_name.lower()
        if (
            query_lower == top_name_lower
            or top_name_lower.startswith(query_lower)
            or query_lower in top_name_lower
        ):
            # Check that runner-up isn't equally good
            runner_name_lower = runner_up.display_name.lower()
            if query_lower not in runner_name_lower:
                LOGGER.info(
                    f"Resolved '{recipient}' -> {top.display_name} ({top.email})"
                )
                return top.email

        # Ambiguous — raise with options
        options = "\n".join(
            f"  - {u.display_name} ({u.email})"
            for u in result.users
        )
        raise ValueError(
            f"Multiple matches for '{recipient}':\n{options}\n"
            f"Please use a more specific name or the email address."
        )

    async def send_direct_message(
        self,
        recipient: str,
        content: str,
        content_type: str = "html",
        sender_email: str | None = None,
    ) -> DirectMessageResult:
        """Send a 1:1 direct message to a user.

        In delegated mode, the message is sent from the signed-in user.
        In app mode, the message is sent from the bot (with attribution).

        If a 1:1 chat already exists between the two users, the existing
        chat is reused (Graph API does not create duplicates).

        Args:
            recipient: Name or email of the recipient. Supports fuzzy name
                matching (e.g. "Alex", "Alex Rivera", "alx rivra") via the
                People API. If an email is provided, it's used directly.
            content: Message content (HTML or text)
            content_type: "html" or "text" (default "html")
            sender_email: Sender email (required for app mode to resolve user ID)

        Returns:
            DirectMessageResult with chat and message IDs
        """
        # Resolve name to email if needed
        recipient_email = await self._resolve_recipient(recipient)

        if self.adapter is None:
            return await self._send_dm_delegated(
                recipient_email=recipient_email,
                content=content,
                content_type=content_type,
            )
        else:
            return await self._send_dm_app(
                recipient_email=recipient_email,
                content=content,
                content_type=content_type,
                sender_email=sender_email,
            )

    async def _send_dm_delegated(
        self,
        recipient_email: str,
        content: str,
        content_type: str = "html",
    ) -> DirectMessageResult:
        """Send a 1:1 DM via Graph API in delegated auth mode.

        Creates (or retrieves existing) 1:1 chat between /me and the
        recipient, then posts a message. The message appears from the
        signed-in user.
        """
        from msgraph.generated.models.aad_user_conversation_member import (
            AadUserConversationMember,
        )
        from msgraph.generated.models.chat import Chat
        from msgraph.generated.models.chat_type import ChatType

        try:
            # Step 1: Resolve the calling user (/me) ID
            me = await self.graph_client.me.get()
            if me is None or not me.id:
                raise ValueError("Could not resolve current user via /me")
            my_id = me.id

            # Step 2: Create or get existing 1:1 chat
            # Use UPN/email directly in user@odata.bind to avoid needing
            # User.Read.All permission for recipient lookup.
            chat_body = Chat()
            chat_body.chat_type = ChatType.OneOnOne
            member_me = AadUserConversationMember(
                odata_type="#microsoft.graph.aadUserConversationMember",
                roles=["owner"],
                additional_data={
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{my_id}')",
                },
            )
            member_recipient = AadUserConversationMember(
                odata_type="#microsoft.graph.aadUserConversationMember",
                roles=["owner"],
                additional_data={
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{recipient_email}')",
                },
            )
            chat_body.members = [member_me, member_recipient]

            chat_response = await self.graph_client.chats.post(chat_body)
            if chat_response is None or not chat_response.id:
                raise ValueError("Failed to create/get 1:1 chat")
            chat_id = chat_response.id

            # Step 3: Send the message
            message_content, body_type = self._prepare_message_body(
                content, content_type
            )

            message = ChatMessage(
                body=ItemBody(
                    content_type=body_type,
                    content=message_content,
                ),
            )
            msg_response = await (
                self.graph_client.chats.by_chat_id(chat_id)
                .messages.post(message)
            )

            message_id = msg_response.id if msg_response else ""

            return DirectMessageResult(
                chat_id=chat_id,
                message_id=message_id or "",
                recipient=recipient_email,
                content_preview=content[:100],
            )
        except Exception as e:
            LOGGER.error(f"Error sending DM (delegated): {str(e)}")
            raise

    async def _send_dm_app(
        self,
        recipient_email: str,
        content: str,
        content_type: str = "html",
        sender_email: str | None = None,
    ) -> DirectMessageResult:
        """Send a 1:1 DM via Graph API in app auth mode.

        Uses application permissions to create (or retrieve existing) 1:1
        chat between the sender and recipient, then posts a message.
        The message appears from the app/bot.
        """
        from msgraph.generated.models.aad_user_conversation_member import (
            AadUserConversationMember,
        )
        from msgraph.generated.models.chat import Chat
        from msgraph.generated.models.chat_type import ChatType

        try:
            if not sender_email:
                raise ValueError("sender_email is required in app auth mode")

            # Step 1: Create or get existing 1:1 chat
            # Use UPN/email directly in user@odata.bind to avoid needing
            # User.Read.All permission for user lookups.
            chat_body = Chat()
            chat_body.chat_type = ChatType.OneOnOne
            member_sender = AadUserConversationMember(
                odata_type="#microsoft.graph.aadUserConversationMember",
                roles=["owner"],
                additional_data={
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{sender_email}')",
                },
            )
            member_recipient = AadUserConversationMember(
                odata_type="#microsoft.graph.aadUserConversationMember",
                roles=["owner"],
                additional_data={
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{recipient_email}')",
                },
            )
            chat_body.members = [member_sender, member_recipient]

            chat_response = await self.graph_client.chats.post(chat_body)
            if chat_response is None or not chat_response.id:
                raise ValueError("Failed to create/get 1:1 chat")
            chat_id = chat_response.id

            # Step 2: Send the message
            # In app mode, Graph API requires Teamwork.Migrate.All for chat
            # messages, so we post on behalf of the sender using their context.
            message_content, body_type = self._prepare_message_body(
                content, content_type
            )

            message = ChatMessage(
                body=ItemBody(
                    content_type=body_type,
                    content=message_content,
                ),
            )
            msg_response = await (
                self.graph_client.chats.by_chat_id(chat_id)
                .messages.post(message)
            )

            message_id = msg_response.id if msg_response else ""

            return DirectMessageResult(
                chat_id=chat_id,
                message_id=message_id or "",
                recipient=recipient_email,
                content_preview=content[:100],
            )
        except Exception as e:
            LOGGER.error(f"Error sending DM (app): {str(e)}")
            raise

    async def _find_user_in_directory_by_email(
        self,
        email: str,
    ) -> dict | None:
        """Search the organization directory for a user by email (mail or UPN).

        Args:
            email: Email address to search for.

        Returns:
            Dict with 'id' and 'display_name' keys, or None if not found.
        """
        try:
            from msgraph.generated.users.users_request_builder import (
                UsersRequestBuilder,
            )

            # Try mail first, then userPrincipalName
            query_params = (
                UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                    filter=f"mail eq '{email}' or userPrincipalName eq '{email}'",
                    select=["id", "displayName", "mail"],
                    top=1,
                )
            )
            config = RequestConfiguration(
                query_parameters=query_params,
            )
            response = await self.graph_client.users.get(
                request_configuration=config,
            )
            if response and response.value and len(response.value) > 0:
                user = response.value[0]
                return {
                    "id": user.id or "",
                    "display_name": user.display_name or email,
                }
            LOGGER.warning(f"User '{email}' not found in org directory")
            return None
        except Exception as e:
            LOGGER.warning(f"Directory search for '{email}' failed: {str(e)}")
            return None

    # ── Meeting Transcript Methods ──────────────────────────────────

    @staticmethod
    def _build_meeting_id_from_join_url(join_url: str) -> str | None:
        """Construct the online meeting ID from a Teams join URL.

        The meeting ID format is deterministic based on the organizer's
        OID and the thread ID embedded in the join URL. This avoids
        needing the OnlineMeetings.Read permission and works for any
        attendee, not just the organizer.

        Format: MSo{organizer_oid}_0#{thread_id}#0
        Then base64-encode the whole thing.
        """
        import base64
        import json
        import re
        from urllib.parse import unquote

        try:
            # Extract thread ID from the URL path
            # URL pattern: .../meetup-join/19%3ameeting_<id>%40thread.v2/...
            decoded_url = unquote(join_url)
            thread_match = re.search(
                r'19:meeting_[a-zA-Z0-9\-]+@thread\.v2',
                decoded_url,
            )
            if not thread_match:
                return None
            thread_id = thread_match.group(0)

            # Extract organizer OID from the context parameter
            context_match = re.search(r'context=(\{.*?\})', decoded_url)
            if not context_match:
                # Try URL-encoded version
                context_match = re.search(
                    r'context=%7[bB](.*?)%7[dD]', join_url
                )
                if context_match:
                    context_str = unquote(f'{{{context_match.group(1)}}}')
                else:
                    return None
            else:
                context_str = context_match.group(1)

            context = json.loads(context_str)
            organizer_oid = context.get("Oid", "")
            tenant_id = context.get("Tid", "")

            if not organizer_oid or not tenant_id:
                return None

            # Build the meeting ID: base64("1*{oid}*0**{thread_id}")
            # "MSo" is not a prefix — it's the base64 encoding of "1*"
            raw_id = f"1*{organizer_oid}*0**{thread_id}"
            meeting_id = base64.b64encode(raw_id.encode()).decode().rstrip("=")
            return meeting_id
        except Exception as e:
            LOGGER.debug(f"Could not build meeting ID from URL: {e}")
            return None

    async def _resolve_meeting_id(
        self,
        join_url: str | None = None,
        meeting_id: str | None = None,
        user_email: str | None = None,
    ) -> str:
        """Resolve a Teams join URL or meeting ID to an online meeting ID.

        Uses three strategies in order:
        1. Direct meeting_id if provided
        2. Construct meeting ID from join URL components (no API call needed)
        3. Fall back to Graph API filter (requires OnlineMeetings.Read, organizer only)

        Args:
            join_url: Teams meeting join URL (from calendar event join_url field).
            meeting_id: Direct online meeting ID if already known.
            user_email: User email for app-mode lookups.

        Returns:
            The online meeting ID string.
        """
        if meeting_id:
            return meeting_id

        if not join_url:
            raise ValueError(
                "Either join_url or meeting_id is required. "
                "Use list_events to find a meeting's join_url."
            )

        # Strategy 1: Build meeting ID from URL components (works for any attendee)
        built_id = self._build_meeting_id_from_join_url(join_url)
        if built_id:
            LOGGER.info(f"Built meeting ID from join URL: {built_id[:20]}...")
            return built_id

        # Strategy 2: Fall back to Graph API filter (organizer only)
        try:
            from msgraph.generated.users.item.online_meetings.online_meetings_request_builder import (
                OnlineMeetingsRequestBuilder,
            )

            safe_url = join_url.replace("'", "''")
            query_params = OnlineMeetingsRequestBuilder.OnlineMeetingsRequestBuilderGetQueryParameters(
                filter=f"joinWebUrl eq '{safe_url}'",
            )
            config = RequestConfiguration(query_parameters=query_params)

            if self.adapter is None:
                response = await self.graph_client.me.online_meetings.get(
                    request_configuration=config,
                )
            else:
                if not user_email:
                    raise ValueError(
                        "user_email is required in app auth mode to resolve meeting ID"
                    )
                response = await (
                    self.graph_client.users.by_user_id(user_email)
                    .online_meetings.get(request_configuration=config)
                )

            if response and response.value and len(response.value) > 0:
                mid = response.value[0].id
                if mid:
                    return mid

            raise ValueError(
                "Could not resolve meeting from join URL via API. "
                "The meeting may have expired (30-day limit)."
            )
        except ValueError:
            raise
        except Exception as e:
            LOGGER.error(f"Error resolving meeting ID from join URL: {e}")
            raise ValueError(
                f"Failed to resolve meeting ID: {str(e)}. "
                f"The join URL may be invalid or the meeting may have expired."
            )

    async def list_meeting_transcripts(
        self,
        join_url: str | None = None,
        meeting_id: str | None = None,
        user_email: str | None = None,
    ) -> TranscriptList:
        """List all transcripts for a completed Teams meeting.

        Args:
            join_url: Teams meeting join URL (from list_events join_url field).
            meeting_id: Online meeting ID if already known.
            user_email: Meeting organizer email (required in app auth mode).

        Returns:
            TranscriptList with transcript metadata.
        """
        resolved_id = await self._resolve_meeting_id(join_url, meeting_id, user_email)

        try:
            if self.adapter is None:
                response = await (
                    self.graph_client.me.online_meetings
                    .by_online_meeting_id(resolved_id)
                    .transcripts.get()
                )
            else:
                if not user_email:
                    raise ValueError("user_email is required in app auth mode")
                response = await (
                    self.graph_client.users.by_user_id(user_email)
                    .online_meetings.by_online_meeting_id(resolved_id)
                    .transcripts.get()
                )

            transcripts: list[TranscriptInfo] = []
            if response and response.value:
                for t in response.value:
                    organizer_name = ""
                    if (
                        t.meeting_organizer
                        and t.meeting_organizer.user
                        and t.meeting_organizer.user.display_name
                    ):
                        organizer_name = t.meeting_organizer.user.display_name

                    transcripts.append(
                        TranscriptInfo(
                            transcript_id=t.id or "",
                            meeting_id=t.meeting_id or resolved_id,
                            created=str(t.created_date_time or ""),
                            end=str(t.end_date_time or ""),
                            organizer=organizer_name,
                        )
                    )

            return TranscriptList(
                transcripts=transcripts,
                meeting_id=resolved_id,
                total=len(transcripts),
            )
        except ValueError:
            raise
        except Exception as e:
            LOGGER.error(f"Error listing transcripts: {e}")
            raise ValueError(
                f"Failed to list transcripts: {str(e)}. "
                f"Ensure transcription was enabled during the meeting "
                f"and the meeting has not expired (30-day limit)."
            )

    async def get_meeting_transcript(
        self,
        transcript_id: str,
        join_url: str | None = None,
        meeting_id: str | None = None,
        user_email: str | None = None,
    ) -> TranscriptInfo:
        """Get metadata for a single transcript.

        Args:
            transcript_id: The transcript ID from list_meeting_transcripts.
            join_url: Teams meeting join URL.
            meeting_id: Online meeting ID if already known.
            user_email: Meeting organizer email (required in app auth mode).

        Returns:
            TranscriptInfo with transcript metadata.
        """
        resolved_id = await self._resolve_meeting_id(join_url, meeting_id, user_email)

        try:
            if self.adapter is None:
                response = await (
                    self.graph_client.me.online_meetings
                    .by_online_meeting_id(resolved_id)
                    .transcripts.by_call_transcript_id(transcript_id)
                    .get()
                )
            else:
                if not user_email:
                    raise ValueError("user_email is required in app auth mode")
                response = await (
                    self.graph_client.users.by_user_id(user_email)
                    .online_meetings.by_online_meeting_id(resolved_id)
                    .transcripts.by_call_transcript_id(transcript_id)
                    .get()
                )

            if not response:
                raise ValueError(f"Transcript {transcript_id} not found")

            organizer_name = ""
            if (
                response.meeting_organizer
                and response.meeting_organizer.user
                and response.meeting_organizer.user.display_name
            ):
                organizer_name = response.meeting_organizer.user.display_name

            return TranscriptInfo(
                transcript_id=response.id or transcript_id,
                meeting_id=response.meeting_id or resolved_id,
                created=str(response.created_date_time or ""),
                end=str(response.end_date_time or ""),
                organizer=organizer_name,
            )
        except ValueError:
            raise
        except Exception as e:
            LOGGER.error(f"Error getting transcript: {e}")
            raise ValueError(f"Failed to get transcript: {str(e)}")

    async def get_transcript_content(
        self,
        transcript_id: str,
        join_url: str | None = None,
        meeting_id: str | None = None,
        user_email: str | None = None,
    ) -> TranscriptContent:
        """Get the actual transcript text content in VTT format.

        Returns the full transcript with speaker names and timestamps.

        Args:
            transcript_id: The transcript ID from list_meeting_transcripts.
            join_url: Teams meeting join URL.
            meeting_id: Online meeting ID if already known.
            user_email: Meeting organizer email (required in app auth mode).

        Returns:
            TranscriptContent with VTT-formatted text.
        """
        resolved_id = await self._resolve_meeting_id(join_url, meeting_id, user_email)

        try:
            # The SDK's content.get() sends the wrong Accept header.
            # We need to set Accept: text/vtt to get VTT-formatted text.
            # Use try_add to replace (not append to) defaults, and create
            # a fresh HeadersCollection to avoid mutating shared state.
            from kiota_abstractions.headers_collection import HeadersCollection
            config = RequestConfiguration()
            config.headers = HeadersCollection()
            config.headers.try_add("Accept", "text/vtt")

            if self.adapter is None:
                response = await (
                    self.graph_client.me.online_meetings
                    .by_online_meeting_id(resolved_id)
                    .transcripts.by_call_transcript_id(transcript_id)
                    .content.get(request_configuration=config)
                )
            else:
                if not user_email:
                    raise ValueError("user_email is required in app auth mode")
                response = await (
                    self.graph_client.users.by_user_id(user_email)
                    .online_meetings.by_online_meeting_id(resolved_id)
                    .transcripts.by_call_transcript_id(transcript_id)
                    .content.get(request_configuration=config)
                )

            if response is None:
                raise ValueError(
                    f"No content returned for transcript {transcript_id}. "
                    f"The transcript may still be processing."
                )

            # Response is bytes — decode to string
            if isinstance(response, bytes):
                content_text = response.decode("utf-8")
            elif isinstance(response, str):
                content_text = response
            else:
                # Could be a stream-like object
                content_text = str(response)

            return TranscriptContent(
                transcript_id=transcript_id,
                meeting_id=resolved_id,
                content=content_text,
            )
        except ValueError:
            raise
        except Exception as e:
            LOGGER.error(f"Error getting transcript content: {e}")
            raise ValueError(f"Failed to get transcript content: {str(e)}")

    async def list_chats(
        self,
        limit: int = 20,
        chat_type_filter: str | None = None,
    ) -> ChatList:
        """List recent chats (1:1, group, and meeting chats).

        Uses GET /me/chats with $expand=members,lastMessagePreview.
        Requires delegated auth with Chat.Read permission.

        Args:
            limit: Maximum number of chats to return (default 20, max 50)
            chat_type_filter: Optional filter: "oneOnOne", "group", or "meeting"

        Returns:
            ChatList with chat metadata, member names, and last message preview
        """
        try:
            limit = min(limit, 50)

            # Build query parameters
            from msgraph.generated.users.item.chats.chats_request_builder import (
                ChatsRequestBuilder,
            )

            params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
                top=limit,
                expand=["members", "lastMessagePreview"],
            )
            if chat_type_filter:
                params.filter = f"chatType eq '{chat_type_filter}'"

            request = RequestConfiguration(query_parameters=params)
            response = await self.graph_client.me.chats.get(
                request_configuration=request
            )

            if response is None or response.value is None:
                return ChatList(chats=[], total=0)

            chats: list[ChatInfo] = []
            for chat in response.value:
                # Extract member display names
                member_names: list[str] = []
                if chat.members:
                    for m in chat.members:
                        name = getattr(m, "display_name", None) or ""
                        if name:
                            member_names.append(name)

                # Extract last message preview
                preview = ""
                if chat.last_message_preview and chat.last_message_preview.body:
                    raw = chat.last_message_preview.body.content or ""
                    preview = _strip_html(raw)
                    if len(preview) > 200:
                        preview = preview[:200] + "..."

                last_updated = ""
                if chat.last_updated_date_time:
                    last_updated = chat.last_updated_date_time.isoformat()

                chats.append(ChatInfo(
                    chat_id=chat.id or "",
                    chat_type=str(chat.chat_type.value) if chat.chat_type else "",
                    topic=chat.topic or "",
                    last_updated=last_updated,
                    last_message_preview=preview,
                    members=member_names,
                ))

            return ChatList(chats=chats, total=len(chats))
        except Exception as e:
            LOGGER.error(f"Error listing chats: {e}")
            raise ValueError(f"Failed to list chats: {str(e)}")

    async def read_chat_messages(
        self,
        chat_id: str,
        limit: int = 30,
        cursor: str | None = None,
    ) -> ChatMessageList:
        """Read messages from a 1:1 or group chat.

        Uses GET /chats/{chat-id}/messages.
        Requires delegated auth with Chat.Read permission.

        Args:
            chat_id: The chat ID (from list_chats or send_direct_message)
            limit: Maximum number of messages to return (default 30, max 50)
            cursor: Pagination cursor (odata nextLink) for subsequent pages

        Returns:
            ChatMessageList with sender, content, and timestamps
        """
        try:
            limit = min(limit, 50)

            from msgraph.generated.chats.item.messages.messages_request_builder import (
                MessagesRequestBuilder as ChatMessagesRequestBuilder,
            )

            params = ChatMessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                top=limit,
                orderby=["createdDateTime desc"],
            )
            request = RequestConfiguration(query_parameters=params)

            if cursor:
                response = await (
                    self.graph_client.chats.by_chat_id(chat_id)
                    .messages.with_url(cursor)
                    .get(request_configuration=request)
                )
            else:
                response = await (
                    self.graph_client.chats.by_chat_id(chat_id)
                    .messages.get(request_configuration=request)
                )

            if response is None or response.value is None:
                return ChatMessageList(chat_id=chat_id, messages=[], total=0)

            messages: list[ChatMessageInfo] = []
            for msg in response.value:
                # Skip system/event messages (memberAdded, etc.)
                if msg.message_type and msg.message_type.value != "message":
                    continue

                sender = ""
                if msg.from_ and msg.from_.user:
                    sender = msg.from_.user.display_name or ""

                content = ""
                if msg.body and msg.body.content:
                    content = _strip_html(msg.body.content)

                extracted = _extract_attachments(msg.attachments)

                created = ""
                if msg.created_date_time:
                    created = msg.created_date_time.isoformat()

                messages.append(ChatMessageInfo(
                    message_id=msg.id or "",
                    sender=sender,
                    content=content,
                    created=created,
                    attachments=extracted,
                ))

            next_cursor = getattr(response, "odata_next_link", None)

            return ChatMessageList(
                chat_id=chat_id,
                messages=messages,
                total=len(messages),
                next_cursor=next_cursor,
            )
        except Exception as e:
            LOGGER.error(f"Error reading chat messages: {e}")
            raise ValueError(f"Failed to read chat messages: {str(e)}")
