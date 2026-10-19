# SPDX-FileCopyrightText: 2025 INDUSTRIA DE DISEÑO TEXTIL, S.A. (INDITEX, S.A.)
# SPDX-FileCopyrightText: 2026 Jason Hastain (modifications)
# SPDX-License-Identifier: Apache-2.0
import logging
import os
import sys
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass
from importlib import metadata

from botbuilder.integration.aiohttp import (
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
)
from dotenv import load_dotenv
from mcp.server.fastmcp import Context, FastMCP
from msgraph.graph_service_client import GraphServiceClient
from pydantic import Field

from .auth import DELEGATED_SCOPES, create_credential
from .config import BotConfiguration
from .teams import (
    AvailableTimesResult,
    CalendarEvent,
    CalendarEventList,
    ChatInfo,
    ChatList,
    ChatMessageInfo,
    ChatMessageList,
    DirectMessageResult,
    ExtractedAttachment,
    MeetingResult,
    PagedTeamsMessages,
    ResolvedUser,
    ResolvedUserList,
    TeamsChannel,
    TeamsClient,
    TeamsMember,
    TeamsMessage,
    TeamsTeam,
    TeamsThread,
    TimeSlot,
    TranscriptContent,
    TranscriptInfo,
    TranscriptList,
)

try:
    __version__ = metadata.version("mcp-teams-server")
except metadata.PackageNotFoundError:
    __version__ = "unknown"

# Load .env
load_dotenv()

# Config logging
logging.basicConfig(
    level=os.environ.get("MCP_LOGLEVEL", "ERROR"),
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stderr),
    ],
)

# Re-exported for consumers of this package; not referenced in this module.
__all__ = [
    "ChatInfo",
    "ChatMessageInfo",
    "ExtractedAttachment",
    "ResolvedUser",
    "TimeSlot",
]

LOGGER = logging.getLogger(__name__)


def _get_required_env_vars() -> list[str]:
    auth_mode = os.environ.get("TEAMS_AUTH_MODE", "app").lower()
    if auth_mode == "delegated":
        return ["TEAMS_APP_ID", "TEAMS_APP_TYPE", "TEAMS_APP_TENANT_ID"]
    else:
        return [
            "TEAMS_APP_ID",
            "TEAMS_APP_PASSWORD",
            "TEAMS_APP_TYPE",
            "TEAMS_APP_TENANT_ID",
        ]


REQUIRED_ENV_VARS = _get_required_env_vars()


@dataclass
class AppContext:
    client: TeamsClient


@asynccontextmanager
async def app_lifespan(server: FastMCP) -> AsyncIterator[AppContext]:
    """Manage application lifecycle with type-safe context"""

    bot_config = BotConfiguration()

    if bot_config.is_delegated:
        # Delegated mode: no bot adapter
        credentials = create_credential(bot_config)
        graph_client = GraphServiceClient(
            credentials=credentials, scopes=DELEGATED_SCOPES
        )
        # Resolve the signed-in user's display name for attribution
        user_display_name = None
        try:
            me = await graph_client.me.get()
            if me and me.display_name:
                user_display_name = me.display_name
        except Exception:
            LOGGER.warning("Could not resolve user display name")
        client = TeamsClient(
            None,  # No bot adapter
            graph_client,
            bot_config.APP_ID,
            bot_config.TEAM_ID,
            bot_config.TEAMS_CHANNEL_ID,
            user_display_name=user_display_name,
            is_delegated=True,
        )
    else:
        # App mode (existing behavior, unchanged)
        adapter = CloudAdapter(ConfigurationBotFrameworkAuthentication(bot_config))
        credentials = create_credential(bot_config)
        scopes = ["https://graph.microsoft.com/.default"]
        graph_client = GraphServiceClient(credentials=credentials, scopes=scopes)
        client = TeamsClient(
            adapter,
            graph_client,
            bot_config.APP_ID,
            bot_config.TEAM_ID,
            bot_config.TEAMS_CHANNEL_ID,
        )

    yield AppContext(client=client)


mcp = FastMCP(
    "mcp-teams-server",
    lifespan=app_lifespan,
    dependencies=[
        "aiohttp",
        "asyncio",
        "botbuilder-core",
        "botbuilder-integration-aiohttp",
        "dotenv",
        "msgraph-sdk",
        "multidict",
    ],
)


def _get_teams_client(ctx: Context) -> TeamsClient:
    return ctx.request_context.lifespan_context.client


_TEAM_ID_DESC = (
    "Team ID. Use list_teams or resolve_teams_link"
    " to discover. If omitted, uses the default."
)
_CHANNEL_ID_DESC = (
    "Channel ID. Use list_channels or resolve_teams_link"
    " to discover. If omitted, uses the default."
)


@mcp.tool(
    name="list_teams",
    description=(
        "List all teams the app can access."
        " Returns team IDs and display names."
    ),
)
async def list_teams(ctx: Context) -> list[TeamsTeam]:
    await ctx.debug("list_teams")
    client = _get_teams_client(ctx)
    return await client.list_teams()


@mcp.tool(
    name="list_channels",
    description="List all channels in a team. Returns channel IDs and display names.",
)
async def list_channels(
    ctx: Context,
    team_id: str | None = Field(
        description=_TEAM_ID_DESC,
        default=None,
    ),
) -> list[TeamsChannel]:
    await ctx.debug(f"list_channels with team_id={team_id}")
    client = _get_teams_client(ctx)
    return await client.list_channels(team_id)


@mcp.tool(
    name="resolve_teams_link",
    description="Parse a Microsoft Teams URL to extract team_id and channel_id.",
)
async def resolve_teams_link(
    ctx: Context,
    url: str = Field(description="The Microsoft Teams URL to parse"),
) -> dict:
    await ctx.debug(f"resolve_teams_link with url={url}")
    client = _get_teams_client(ctx)
    return await client.resolve_teams_link(url)


@mcp.tool(
    name="create_channel",
    description=(
        "Create a new channel in a team."
        " Returns the created channel with its ID."
    ),
)
async def create_channel(
    ctx: Context,
    display_name: str = Field(description="Display name for the new channel"),
    description: str | None = Field(
        description="Optional description for the channel", default=None
    ),
    channel_type: str = Field(
        description=(
            "Channel type: 'standard' (visible to all"
            " team members) or 'private' (invite-only)"
        ),
        default="standard",
    ),
    team_id: str | None = Field(
        description=_TEAM_ID_DESC,
        default=None,
    ),
) -> TeamsChannel:
    await ctx.debug(f"create_channel with display_name={display_name}")
    client = _get_teams_client(ctx)
    return await client.create_channel(display_name, description, channel_type, team_id)


@mcp.tool(
    name="add_channel_members",
    description="Add members to a channel by their email addresses.",
)
async def add_channel_members(
    ctx: Context,
    channel_id: str = Field(description="The channel ID to add members to"),
    member_emails: list[str] = Field(
        description="List of email addresses of members to add to the channel"
    ),
    team_id: str | None = Field(
        description=_TEAM_ID_DESC,
        default=None,
    ),
) -> list[TeamsMember]:
    await ctx.debug(
        f"add_channel_members with channel_id={channel_id},"
        f" members={member_emails}"
    )
    client = _get_teams_client(ctx)
    return await client.add_channel_members(channel_id, member_emails, team_id)


@mcp.tool(
    name="start_thread", description="Start a new thread with a given title and content"
)
async def start_thread(
    ctx: Context,
    title: str = Field(description="The thread title"),
    content: str = Field(description="The thread content"),
    member_name: str | None = Field(
        description="Member name to mention in the thread", default=None
    ),
    team_id: str | None = Field(
        description=_TEAM_ID_DESC,
        default=None,
    ),
    channel_id: str | None = Field(
        description=_CHANNEL_ID_DESC,
        default=None,
    ),
) -> TeamsThread:
    await ctx.debug(f"start_thread with title={title} and content={content}")
    client = _get_teams_client(ctx)
    return await client.start_thread(title, content, member_name, team_id, channel_id)


@mcp.tool(
    name="update_thread", description="Update an existing thread with new content"
)
async def update_thread(
    ctx: Context,
    thread_id: str = Field(
        description="The thread ID as a string in the format '1743086901347'"
    ),
    content: str = Field(description="The content to update in the thread"),
    member_name: str | None = Field(
        description="Member name to mention in the thread", default=None
    ),
    team_id: str | None = Field(
        description=_TEAM_ID_DESC,
        default=None,
    ),
    channel_id: str | None = Field(
        description=_CHANNEL_ID_DESC,
        default=None,
    ),
) -> TeamsMessage:
    await ctx.debug(f"update_thread with thread_id={thread_id} and content={content}")
    client = _get_teams_client(ctx)
    return await client.update_thread(
        thread_id, content, member_name, team_id, channel_id
    )


@mcp.tool(name="read_thread", description="Read replies in a thread")
async def read_thread(
    ctx: Context,
    thread_id: str = Field(
        description="The thread ID as a string in the format '1743086901347'"
    ),
    team_id: str | None = Field(
        description=_TEAM_ID_DESC,
        default=None,
    ),
    channel_id: str | None = Field(
        description=_CHANNEL_ID_DESC,
        default=None,
    ),
) -> PagedTeamsMessages:
    await ctx.debug(f"read_thread with thread_id={thread_id}")
    client = _get_teams_client(ctx)
    return await client.read_thread_replies(
        thread_id, 50, team_id=team_id, channel_id=channel_id
    )


@mcp.tool(name="list_threads", description="List threads in channel with pagination")
async def list_threads(
    ctx: Context,
    limit: int = Field(
        description="Maximum number of items to retrieve or page size", default=50
    ),
    cursor: str | None = Field(
        description="Pagination cursor for the next page of results", default=None
    ),
    team_id: str | None = Field(
        description=_TEAM_ID_DESC,
        default=None,
    ),
    channel_id: str | None = Field(
        description=_CHANNEL_ID_DESC,
        default=None,
    ),
) -> PagedTeamsMessages:
    await ctx.debug(f"list_threads with cursor={cursor} and limit={limit}")
    client = _get_teams_client(ctx)
    return await client.read_threads(
        limit, cursor, team_id=team_id, channel_id=channel_id
    )


@mcp.tool(name="get_member_by_name", description="Get a member by its name")
async def get_member_by_name(
    ctx: Context,
    name: str = Field(description="Member name"),
    team_id: str | None = Field(
        description=_TEAM_ID_DESC,
        default=None,
    ),
    channel_id: str | None = Field(
        description=_CHANNEL_ID_DESC,
        default=None,
    ),
):
    await ctx.debug(f"get_member_by_name with name={name}")
    client = _get_teams_client(ctx)
    return await client.get_member_by_name(name, team_id=team_id, channel_id=channel_id)


@mcp.tool(name="list_members", description="List all members in the team")
async def list_members(
    ctx: Context,
    team_id: str | None = Field(
        description=_TEAM_ID_DESC,
        default=None,
    ),
    channel_id: str | None = Field(
        description=_CHANNEL_ID_DESC,
        default=None,
    ),
) -> list[TeamsMember]:
    await ctx.debug("list_members")
    client = _get_teams_client(ctx)
    return await client.list_members(team_id=team_id, channel_id=channel_id)


@mcp.tool(
    name="find_available_times",
    description=(
        "Find common available meeting times for a group of people."
        " Queries calendar free/busy data and returns time slots"
        " when all attendees are free."
    ),
)
async def find_available_times(
    ctx: Context,
    attendee_emails: list[str] = Field(
        description="List of email addresses to check availability for"
    ),
    date: str = Field(
        description="Date to check availability in YYYY-MM-DD format"
    ),
    start_hour: int = Field(
        description="Start of availability window (hour, 0-23). Default 8 (8am).",
        default=8,
    ),
    end_hour: int = Field(
        description="End of availability window (hour, 0-23). Default 17 (5pm).",
        default=17,
    ),
    duration_minutes: int = Field(
        description="Required meeting duration in minutes. Default 60.",
        default=60,
    ),
    exclude_start_hour: int | None = Field(
        description="Start hour of excluded block (e.g., lunch). Default 12. Set to null to disable.",
        default=12,
    ),
    exclude_end_hour: int | None = Field(
        description="End hour of excluded block. Default 13. Set to null to disable.",
        default=13,
    ),
    timezone: str = Field(
        description='Timezone for the query. Default "Central Standard Time".',
        default="Central Standard Time",
    ),
) -> AvailableTimesResult:
    await ctx.debug(f"find_available_times for {attendee_emails} on {date}")
    client = _get_teams_client(ctx)
    return await client.find_available_times(
        attendee_emails=attendee_emails,
        date=date,
        start_hour=start_hour,
        end_hour=end_hour,
        duration_minutes=duration_minutes,
        exclude_start_hour=exclude_start_hour,
        exclude_end_hour=exclude_end_hour,
        timezone=timezone,
    )


@mcp.tool(
    name="create_meeting",
    description=(
        "Create a calendar event with an optional Teams meeting link."
        " Supports one-time and recurring meetings (daily, weekly)."
        " Sends invitations to all attendees."
    ),
)
async def create_meeting(
    ctx: Context,
    subject: str = Field(description="Meeting subject/title"),
    start_datetime: str = Field(
        description='Start time in ISO format, e.g., "2026-03-04T09:00:00"'
    ),
    end_datetime: str = Field(
        description='End time in ISO format, e.g., "2026-03-04T09:30:00"'
    ),
    attendee_emails: list[str] = Field(
        description="List of attendee email addresses"
    ),
    body: str | None = Field(
        description="Optional HTML body content for the meeting invite",
        default=None,
    ),
    timezone: str = Field(
        description='Timezone. Default "Central Standard Time".',
        default="Central Standard Time",
    ),
    is_online: bool = Field(
        description="Create a Teams meeting link. Default true.",
        default=True,
    ),
    recurrence_type: str | None = Field(
        description='Recurrence type: "daily", "weekly", or null for one-time. Default null.',
        default=None,
    ),
    recurrence_interval: int = Field(
        description="Interval between recurrences (e.g., 1 = every week). Default 1.",
        default=1,
    ),
    days_of_week: list[str] | None = Field(
        description='Days for weekly recurrence, e.g., ["wednesday"]. Required if recurrence_type is "weekly".',
        default=None,
    ),
    recurrence_end_date: str | None = Field(
        description="End date for recurrence in YYYY-MM-DD format. Null for no end date.",
        default=None,
    ),
    organizer_email: str | None = Field(
        description="Organizer email. Required in app auth mode, optional in delegated mode.",
        default=None,
    ),
) -> MeetingResult:
    await ctx.debug(f"create_meeting: {subject} for {attendee_emails}")
    client = _get_teams_client(ctx)
    return await client.create_meeting(
        subject=subject,
        start_datetime=start_datetime,
        end_datetime=end_datetime,
        attendee_emails=attendee_emails,
        body=body,
        timezone=timezone,
        is_online=is_online,
        recurrence_type=recurrence_type,
        recurrence_interval=recurrence_interval,
        days_of_week=days_of_week,
        recurrence_end_date=recurrence_end_date,
        organizer_email=organizer_email,
    )


@mcp.tool(
    name="cancel_meeting",
    description=(
        "Cancel a calendar event by its event ID."
        " Optionally sends a cancellation message to attendees."
        " Works for both one-time and recurring meetings."
    ),
)
async def cancel_meeting(
    ctx: Context,
    event_id: str = Field(description="The event ID returned from create_meeting"),
    organizer_email: str | None = Field(
        description="Organizer email. Required in app auth mode, optional in delegated mode.",
        default=None,
    ),
    comment: str | None = Field(
        description="Optional cancellation message sent to all attendees.",
        default=None,
    ),
) -> dict:
    await ctx.debug(f"cancel_meeting: {event_id}")
    client = _get_teams_client(ctx)
    return await client.cancel_meeting(
        event_id=event_id,
        organizer_email=organizer_email,
        comment=comment,
    )


@mcp.tool(
    name="list_events",
    description=(
        "List or search calendar events."
        " Supports filtering by subject text, date range, and sorting."
        " Returns event details including subject, times, attendees,"
        " recurrence info, and Teams meeting join URLs."
    ),
)
async def list_events(
    ctx: Context,
    user_email: str | None = Field(
        description="User whose calendar to query. Required in app auth mode, optional in delegated mode.",
        default=None,
    ),
    subject_filter: str | None = Field(
        description="Filter events whose subject contains this text.",
        default=None,
    ),
    start_date: str | None = Field(
        description="Filter events starting on or after this date (YYYY-MM-DD).",
        default=None,
    ),
    end_date: str | None = Field(
        description="Filter events starting on or before this date (YYYY-MM-DD).",
        default=None,
    ),
    top: int = Field(
        description="Maximum number of events to return. Default 25.",
        default=25,
    ),
    order_by: str = Field(
        description='Sort order. Default "start/dateTime desc".',
        default="start/dateTime desc",
    ),
) -> CalendarEventList:
    await ctx.debug(f"list_events: subject={subject_filter}, start={start_date}, end={end_date}")
    client = _get_teams_client(ctx)
    return await client.list_events(
        user_email=user_email,
        subject_filter=subject_filter,
        start_date=start_date,
        end_date=end_date,
        top=top,
        order_by=order_by,
    )


@mcp.tool(
    name="update_event",
    description=(
        "Update an existing calendar event."
        " Can modify subject, body, start/end times, attendees,"
        " recurrence, and online meeting settings."
        " For recurring meetings, updating the series master event"
        " updates all occurrences in the series."
    ),
)
async def update_event(
    ctx: Context,
    event_id: str = Field(
        description="The event ID to update. For recurring series, use the series master event ID."
    ),
    user_email: str | None = Field(
        description="User whose calendar contains the event. Required in app auth mode.",
        default=None,
    ),
    subject: str | None = Field(
        description="New subject/title. Omit to keep current.",
        default=None,
    ),
    body: str | None = Field(
        description="New HTML body content. Omit to keep current.",
        default=None,
    ),
    start_datetime: str | None = Field(
        description='New start time in ISO format (e.g., "2026-03-04T09:00:00"). Omit to keep current.',
        default=None,
    ),
    end_datetime: str | None = Field(
        description='New end time in ISO format (e.g., "2026-03-04T09:30:00"). Omit to keep current.',
        default=None,
    ),
    timezone: str | None = Field(
        description="Timezone for start/end times. Omit to keep current.",
        default=None,
    ),
    attendee_emails: list[str] | None = Field(
        description="New attendee list (replaces existing). Omit to keep current.",
        default=None,
    ),
    is_online: bool | None = Field(
        description="Enable/disable Teams meeting link. Omit to keep current.",
        default=None,
    ),
    recurrence_type: str | None = Field(
        description='New recurrence type: "daily", "weekly". Omit to keep current.',
        default=None,
    ),
    recurrence_interval: int | None = Field(
        description="Interval between recurrences. Omit to keep current.",
        default=None,
    ),
    days_of_week: list[str] | None = Field(
        description='Days for weekly recurrence (e.g., ["wednesday"]). Omit to keep current.',
        default=None,
    ),
    recurrence_end_date: str | None = Field(
        description="End date for recurrence (YYYY-MM-DD). Omit to keep current.",
        default=None,
    ),
    remove_recurrence: bool = Field(
        description="Set to true to remove recurrence from a recurring event, making it one-time.",
        default=False,
    ),
) -> CalendarEvent:
    await ctx.debug(f"update_event: {event_id}")
    client = _get_teams_client(ctx)
    return await client.update_event(
        event_id=event_id,
        user_email=user_email,
        subject=subject,
        body=body,
        start_datetime=start_datetime,
        end_datetime=end_datetime,
        timezone=timezone,
        attendee_emails=attendee_emails,
        is_online=is_online,
        recurrence_type=recurrence_type,
        recurrence_interval=recurrence_interval,
        days_of_week=days_of_week,
        recurrence_end_date=recurrence_end_date,
        remove_recurrence=remove_recurrence,
    )


@mcp.tool(
    name="resolve_user",
    description=(
        "Resolve a person's name to their email and identity."
        " Supports fuzzy matching: partial names ('Alex'), full names"
        " ('Alex Rivera'), misspellings ('alx rivra'), and nicknames."
        " Results are ranked by relevance to the signed-in user."
        " Use this to look up people before sending messages or"
        " scheduling meetings."
    ),
)
async def resolve_user(
    ctx: Context,
    query: str = Field(
        description=(
            "Name or partial name to search for."
            " Examples: 'Alex', 'Alex Rivera', 'alx rivra', 'dana'."
        ),
    ),
    max_results: int = Field(
        description="Maximum number of results to return.",
        default=5,
    ),
) -> ResolvedUserList:
    await ctx.debug(f"resolve_user: query={query}")
    client = _get_teams_client(ctx)
    return await client.resolve_user(query=query, max_results=max_results)


@mcp.tool(
    name="send_direct_message",
    description=(
        "Send a 1:1 direct message to a user in Microsoft Teams."
        " Creates or reuses an existing 1:1 chat between the caller"
        " and the recipient. Accepts a name (e.g. 'Alex Rivera', 'Alex')"
        " or email address — names are automatically resolved via fuzzy"
        " matching. In delegated mode, the message appears from the"
        " signed-in user. Supports HTML and plain text content."
    ),
)
async def send_direct_message(
    ctx: Context,
    recipient: str = Field(
        description=(
            "Name or email of the recipient. Supports fuzzy name matching:"
            " 'Alex', 'Alex Rivera', 'alx rivra', or 'alex.rivera@example.com'."
        ),
    ),
    content: str = Field(
        description="Message content to send. Supports HTML formatting."
    ),
    content_type: str = Field(
        description='Content format: "html" (default) or "text".',
        default="html",
    ),
    sender_email: str | None = Field(
        description=(
            "Sender email address. Required in app auth mode to identify"
            " who is sending. In delegated mode this is ignored (uses /me)."
        ),
        default=None,
    ),
) -> DirectMessageResult:
    await ctx.debug(f"send_direct_message: to={recipient}")
    client = _get_teams_client(ctx)
    return await client.send_direct_message(
        recipient=recipient,
        content=content,
        content_type=content_type,
        sender_email=sender_email,
    )


_MEETING_JOIN_URL_DESC = (
    "Teams meeting join URL. Get this from list_events (the join_url field)"
    " or copy it from a calendar invite. Looks like:"
    " https://teams.microsoft.com/l/meetup-join/19%3a..."
)
_MEETING_ID_DESC = (
    "Online meeting ID (base64 string). Get this from a previous"
    " list_meeting_transcripts call. Use join_url instead if you"
    " have the meeting link."
)
_TRANSCRIPT_ID_DESC = (
    "Transcript ID (base64 string like 'MSMjMCMj...')."
    " Get this from list_meeting_transcripts."
)


@mcp.tool(
    name="list_meeting_transcripts",
    description=(
        "List all transcripts for a completed Teams meeting."
        " Requires the meeting join URL (from list_events) or"
        " an online meeting ID. Transcription must have been"
        " enabled during the meeting. Transcripts are available"
        " for ~30 days after the meeting ends."
    ),
)
async def list_meeting_transcripts(
    ctx: Context,
    join_url: str | None = Field(
        description=_MEETING_JOIN_URL_DESC,
        default=None,
    ),
    meeting_id: str | None = Field(
        description=_MEETING_ID_DESC,
        default=None,
    ),
    user_email: str | None = Field(
        description=(
            "Meeting organizer email. Required in app auth mode,"
            " ignored in delegated mode."
        ),
        default=None,
    ),
) -> TranscriptList:
    await ctx.debug(f"list_meeting_transcripts: join_url={join_url}, meeting_id={meeting_id}")
    client = _get_teams_client(ctx)
    return await client.list_meeting_transcripts(
        join_url=join_url,
        meeting_id=meeting_id,
        user_email=user_email,
    )


@mcp.tool(
    name="get_meeting_transcript",
    description=(
        "Get metadata for a single transcript (creation time,"
        " end time, organizer). Use list_meeting_transcripts"
        " first to get the transcript_id."
    ),
)
async def get_meeting_transcript(
    ctx: Context,
    transcript_id: str = Field(description=_TRANSCRIPT_ID_DESC),
    join_url: str | None = Field(
        description=_MEETING_JOIN_URL_DESC,
        default=None,
    ),
    meeting_id: str | None = Field(
        description=_MEETING_ID_DESC,
        default=None,
    ),
    user_email: str | None = Field(
        description=(
            "Meeting organizer email. Required in app auth mode,"
            " ignored in delegated mode."
        ),
        default=None,
    ),
) -> TranscriptInfo:
    await ctx.debug(f"get_meeting_transcript: transcript_id={transcript_id}")
    client = _get_teams_client(ctx)
    return await client.get_meeting_transcript(
        transcript_id=transcript_id,
        join_url=join_url,
        meeting_id=meeting_id,
        user_email=user_email,
    )


@mcp.tool(
    name="get_transcript_content",
    description=(
        "Get the full transcript text for a completed Teams meeting."
        " Returns VTT-formatted content with timestamps and speaker"
        " names. Use list_meeting_transcripts first to get the"
        " transcript_id. Example output:"
        '\n  00:00:16.246 --> 00:00:17.726'
        '\n  <v Dana Okafor>Let\'s start with the recap.</v>'
    ),
)
async def get_transcript_content(
    ctx: Context,
    transcript_id: str = Field(description=_TRANSCRIPT_ID_DESC),
    join_url: str | None = Field(
        description=_MEETING_JOIN_URL_DESC,
        default=None,
    ),
    meeting_id: str | None = Field(
        description=_MEETING_ID_DESC,
        default=None,
    ),
    user_email: str | None = Field(
        description=(
            "Meeting organizer email. Required in app auth mode,"
            " ignored in delegated mode."
        ),
        default=None,
    ),
) -> TranscriptContent:
    await ctx.debug(f"get_transcript_content: transcript_id={transcript_id}")
    client = _get_teams_client(ctx)
    return await client.get_transcript_content(
        transcript_id=transcript_id,
        join_url=join_url,
        meeting_id=meeting_id,
        user_email=user_email,
    )


@mcp.tool(
    name="list_chats",
    description=(
        "List recent 1:1 and group chats for the signed-in user."
        " Returns chat metadata including member names and a preview"
        " of the last message. Use the chat_id from results with"
        " read_chat_messages to read full conversation history."
    ),
)
async def list_chats(
    ctx: Context,
    limit: int = Field(
        description="Maximum number of chats to return (1-50).",
        default=20,
    ),
    chat_type: str | None = Field(
        description=(
            "Filter by chat type: 'oneOnOne' for 1:1 chats,"
            " 'group' for group chats, 'meeting' for meeting chats."
            " Omit to return all types."
        ),
        default=None,
    ),
) -> ChatList:
    await ctx.debug(f"list_chats: limit={limit}, chat_type={chat_type}")
    client = _get_teams_client(ctx)
    return await client.list_chats(limit=limit, chat_type_filter=chat_type)


@mcp.tool(
    name="read_chat_messages",
    description=(
        "Read messages from a 1:1 or group chat. Returns messages"
        " with sender names, content (HTML stripped to plain text),"
        " and timestamps. Get the chat_id from list_chats or"
        " send_direct_message. Supports pagination via cursor."
    ),
)
async def read_chat_messages(
    ctx: Context,
    chat_id: str = Field(
        description=(
            "The chat ID to read messages from. Get this from"
            " list_chats (the chat_id field) or from"
            " send_direct_message (the chat_id in the result)."
        ),
    ),
    limit: int = Field(
        description="Maximum number of messages to return (1-50).",
        default=30,
    ),
    cursor: str | None = Field(
        description=(
            "Pagination cursor from a previous read_chat_messages"
            " call (the next_cursor field). Omit for the first page."
        ),
        default=None,
    ),
) -> ChatMessageList:
    await ctx.debug(f"read_chat_messages: chat_id={chat_id}, limit={limit}")
    client = _get_teams_client(ctx)
    return await client.read_chat_messages(
        chat_id=chat_id, limit=limit, cursor=cursor
    )


def _check_required_environment():
    exit_code = None
    for var in REQUIRED_ENV_VARS:
        value = os.environ.get(var)
        if value is None:
            LOGGER.info(f"Required ENV {var} not present")
            exit_code = 1
    if exit_code is not None:
        sys.exit(exit_code)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="MCP Teams Server to allow Microsoft Teams interaction"
    )

    default_transport = os.environ.get("MCP_TRANSPORT", "stdio")

    parser.add_argument(
        "-t",
        "--transport",
        nargs=1,
        type=str,
        help="MCP Server Transport: stdio or sse",
        default=default_transport,
        choices=["stdio", "sse"],
    )

    args = parser.parse_args()

    LOGGER.info(
        f'Starting MCP Teams Server "{__version__}" with transport "{args.transport}"'
    )
    _check_required_environment()
    mcp.run(transport=args.transport)


if __name__ == "__main__":
    main()
